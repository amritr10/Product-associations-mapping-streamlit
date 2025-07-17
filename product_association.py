import streamlit as st
import pandas as pd
import re
import io
import json
from datetime import datetime
from xlsxwriter.utility import xl_col_to_name
import plotly.express as px


# ---------------------------
# Helper Functions
# ---------------------------

def read_excel(file_input):
    """Read Excel or CSV into DataFrame."""
    name = getattr(file_input, "name", "").lower()
    try:
        if name.endswith(".csv"):
            return pd.read_csv(file_input)
        else:
            return pd.read_excel(file_input)
    except Exception as e:
        st.error(f"Failed to read file: {e}")
        return pd.DataFrame()

def filter_and_sum(df, search_terms, exclusion_terms=None,
                   column_name='Description', sum_columns=['IA','FA']):
    """Filter by wildcard search_terms (+ exclusions), sum IA/FA."""
    def pattern(t):
        return '^' + re.escape(t.strip()).replace(r'\*','.*') + '$'
    pats = [pattern(t) for t in search_terms if t.strip()]
    if not pats:
        return pd.DataFrame(), {c:0.0 for c in sum_columns}
    rx = re.compile("|".join(pats), flags=re.IGNORECASE)
    mask = df[column_name].astype(str).str.match(rx)
    if exclusion_terms:
        ex_pats = [pattern(t) for t in exclusion_terms if t.strip()]
        if ex_pats:
            ex_rx = re.compile("|".join(ex_pats), flags=re.IGNORECASE)
            mask &= ~df[column_name].astype(str).str.match(ex_rx)
    sub = df[mask].copy()
    sums = {c: float(sub[c].sum()) if c in sub else 0.0 for c in sum_columns}
    return sub, sums

def calculate_balances(parent_sums, dependent_sums, col, multiple=1.0):
    return parent_sums.get(col,0.0)*multiple - dependent_sums.get(col,0.0)

def get_safe_sheet_name(name, existing):
    """Truncate to ≤31 chars and add suffix if needed."""
    max_len = 31
    base = name[:max_len]
    if base not in existing:
        return base
    for i in range(1,1000):
        cand = f"{base[:max_len-len(str(i))-3]} ({i})"
        if cand not in existing:
            return cand
    return f"Sheet{len(existing)+1}"

def save_to_excel_bytes(
    mappings_data, mapping_to_parent, dependent_to_mappings,
    parent_dfs, dependent_dfs,
    parent_names, dependent_names,
    parent_to_mappings,
    parent_sheet_names, dependent_sheet_names,
    parent_sums, dependent_sums
):
    """Produce an XLSX in memory, with multiple sheets and hyperlinks."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        wb = writer.book

        # Formats
        header_fmt   = wb.add_format({'bold':True,'fg_color':'#D7E4BC','border':1,'text_wrap':True})
        yellow_hdr   = wb.add_format({'bold':True,'fg_color':'yellow','border':1,'text_wrap':True})
        blue_hdr     = wb.add_format({'bold':True,'fg_color':'blue','font_color':'white','border':1,'text_wrap':True})
        red_font_hdr = wb.add_format({'bold':True,'fg_color':'#D7E4BC','font_color':'red','border':1,'text_wrap':True})
        link_fmt     = wb.add_format({'font_color':'blue','underline':True})
        balance_fmt  = wb.add_format({'font_color':'red'})
        heading_fmt  = wb.add_format({'bold':True,'font_size':14})

        #
        # 1) Dependent Perspective
        #
        dep_sheet = "Dependent Perspective"
        ws_dep = wb.add_worksheet(dep_sheet)
        writer.sheets[dep_sheet] = ws_dep

        # Build header row
        base_cols = [
            "Dependent Group","Linked Parent Groups",
            "Total Dependent SKUs","Total Parent SKUs",
            "Balance IA","Balance FA",
            "Total IA Dependent Group","Total FA Dependent Group",
            "Total IA Mapped Parents","Total FA Mapped Parents"
        ]
        for i,pn in enumerate(parent_names, start=1):
            base_cols += [
              f"Parent {i} Name",
              f"Parent {i} Mapping Type",
              f"Parent {i} Multiple",
              f"Parent {i} Total IA",
              f"Parent {i} Total FA"
            ]
        for c, h in enumerate(base_cols):
            fmt = header_fmt
            if h in ("Dependent Group","Total IA Dependent Group","Total FA Dependent Group"):
                fmt = yellow_hdr
            elif h in ("Linked Parent Groups","Total IA Mapped Parents","Total FA Mapped Parents"):
                fmt = blue_hdr
            ws_dep.write(0, c, h, fmt)

        row = 1
        for dg in dependent_names:
            linked = dependent_to_mappings.get(dg, [])
            linked_str = ", ".join(linked) if linked else "None"
            ia_dep = dependent_sums[dg]['IA']
            fa_dep = dependent_sums[dg]['FA']
            total_dep_skus = len(dependent_dfs[dg])
            total_parent_skus = sum(len(parent_dfs[p]) for p in linked)

            # Dependent hyperlink
            ws_dep.write_formula(
                row, 0,
                f'=HYPERLINK("#\'{dependent_sheet_names[dg]}\'!A1","{dg}")',
                link_fmt
            )
            ws_dep.write(row,1, linked_str)
            ws_dep.write_number(row,2, total_dep_skus)
            ws_dep.write_number(row,3, total_parent_skus)
            ws_dep.write_number(row,6, ia_dep)
            ws_dep.write_number(row,7, fa_dep)
            ws_dep.write_number(row,8, sum(parent_sums[p]['IA'] for p in linked))
            ws_dep.write_number(row,9, sum(parent_sums[p]['FA'] for p in linked))

            ia_parts = []
            fa_parts = []
            base = 10
            for i,pn in enumerate(parent_names, start=1):
                mtype = ""
                mult  = 1.0
                if pn in linked:
                    for mg, df_m in mappings_data.items():
                        if mapping_to_parent.get(mg)==pn and dg in df_m['Dependent Group'].values:
                            r = df_m[df_m['Dependent Group']==dg].iloc[0]
                            mtype = r['Type']
                            try: mult = float(r['Multiple'])
                            except: mult = 1.0
                            break

                # Write the five parent columns
                ws_dep.write_formula(
                    row, base+(i-1)*5,
                    f'=HYPERLINK("#\'{parent_sheet_names[pn]}\'!A1","{pn}")',
                    link_fmt
                )
                ws_dep.write(row, base+(i-1)*5+1, mtype)
                ws_dep.write_number(row, base+(i-1)*5+2, mult)
                ws_dep.write_number(row, base+(i-1)*5+3, parent_sums[pn]['IA'])
                ws_dep.write_number(row, base+(i-1)*5+4, parent_sums[pn]['FA'])

                if mtype:
                    c_mult = xl_col_to_name(base+(i-1)*5+2)
                    c_ia   = xl_col_to_name(base+(i-1)*5+3)
                    c_fa   = xl_col_to_name(base+(i-1)*5+4)
                    ia_parts.append(f"{c_mult}{row+1}*{c_ia}{row+1}")
                    fa_parts.append(f"{c_mult}{row+1}*{c_fa}{row+1}")

            # Balance formulas
            if ia_parts:
                ws_dep.write_formula(
                    row,4,
                    f"=G{row+1} - (" + " + ".join(ia_parts) + ")",
                    balance_fmt
                )
            else:
                ws_dep.write(row,4,"", balance_fmt)

            if fa_parts:
                ws_dep.write_formula(
                    row,5,
                    f"=H{row+1} - (" + " + ".join(fa_parts) + ")",
                    balance_fmt
                )
            else:
                ws_dep.write(row,5,"", balance_fmt)

            row += 1

        # Adjust widths + color
        for c,h in enumerate(base_cols):
            w=20
            if "Linked Parent" in h: w=40
            elif "Name" in h:       w=30
            ws_dep.set_column(c,c,w)
        ws_dep.set_tab_color('green')

        #
        # 2) All Mappings
        #
        map_sheet = "All Mappings"
        ws_map = wb.add_worksheet(map_sheet)
        writer.sheets[map_sheet] = ws_map

        headers = [
            "Parent Group","Dependent Group","Mapping Type","Multiple",
            "Balance IA","Balance FA",
            "Total IA Parent Group","Total FA Parent Group",
            "Total IA Dependent Group","Total FA Dependent Group"
        ]
        for c,h in enumerate(headers):
            fmt = header_fmt
            if h in ("Balance IA","Balance FA"):
                fmt = red_font_hdr
            ws_map.write(0,c,h,fmt)

        r = 1
        for mg, df_m in mappings_data.items():
            pg = mapping_to_parent.get(mg,"")
            ws_map.write(r,0, mg, heading_fmt)
            r += 1
            for _, rd in df_m.iterrows():
                ws_map.write_formula(
                    r,0,
                    f'=HYPERLINK("#\'{parent_sheet_names[pg]}\'!A1","{pg}")',
                    link_fmt
                )
                dg = rd['Dependent Group']
                ws_map.write_formula(
                    r,1,
                    f'=HYPERLINK("#\'{dependent_sheet_names[dg]}\'!A1","{dg}")',
                    link_fmt
                )
                ws_map.write(r,2, rd['Type'])
                try:
                    mnum = float(rd['Multiple'])
                    ws_map.write_number(r,3,mnum)
                    ws_map.write_number(r,4,rd['Balance IA'])
                    ws_map.write_number(r,5,rd['Balance FA'])
                except:
                    ws_map.write(r,3,"")
                    ws_map.write(r,4,0)
                    ws_map.write(r,5,0)
                ws_map.write(r,6, parent_sums[pg]['IA'])
                ws_map.write(r,7, parent_sums[pg]['FA'])
                ws_map.write(r,8, dependent_sums[dg]['IA'])
                ws_map.write(r,9, dependent_sums[dg]['FA'])
                r += 1
            r += 1
        for c,_ in enumerate(headers):
            ws_map.set_column(c,c,20)
        ws_map.set_tab_color('red')

        #
        # 3) Parent sheets
        #
        for pn, df_p in parent_dfs.items():
            sn = parent_sheet_names[pn]
            total_row = {
                'Description': 'Total Parent Group',
                'IA': parent_sums[pn]['IA'],
                'FA': parent_sums[pn]['FA']
            }
            dfp = pd.concat([df_p, pd.DataFrame([total_row])], ignore_index=True)
            dfp.to_excel(writer, sheet_name=sn, index=False, startrow=1)
            ws3 = writer.sheets[sn]
            ws3.write_formula(
                0,0,
                f'=HYPERLINK("#\'{dep_sheet}\'!A1","Go to Dependent Perspective")',
                link_fmt
            )
            for c,col in enumerate(dfp.columns):
                ws3.write(1,c,col,header_fmt)
                width = max(dfp[col].astype(str).map(len).max(), len(col)) + 2
                ws3.set_column(c,c,min(width,50))
            ws3.set_tab_color('blue')

        #
        # 4) Dependent sheets
        #
        for dg, df_d in dependent_dfs.items():
            sn = dependent_sheet_names[dg]
            total_row = {
                'Description': 'Total Dependent Group',
                'IA': dependent_sums[dg]['IA'],
                'FA': dependent_sums[dg]['FA']
            }
            dfd = pd.concat([df_d, pd.DataFrame([total_row])], ignore_index=True)
            dfd.to_excel(writer, sheet_name=sn, index=False, startrow=1)
            ws4 = writer.sheets[sn]
            ws4.write_formula(
                0,0,
                f'=HYPERLINK("#\'{dep_sheet}\'!A1","Go to Dependent Perspective")',
                link_fmt
            )
            for c,col in enumerate(dfd.columns):
                ws4.write(1,c,col,header_fmt)
                width = max(dfd[col].astype(str).map(len).max(), len(col)) + 2
                ws4.set_column(c,c,min(width,50))
            ws4.set_tab_color('yellow')

    output.seek(0)
    return output.getvalue()

def flatten_mappings(mappings_data, mapping_to_parent, dependent_sums):
    rows = []
    for mg, df_m in mappings_data.items():
        pg = mapping_to_parent.get(mg,"")
        for _, r in df_m.iterrows():
            dg = r['Dependent Group']
            rows.append({
                "Mapping Group": mg,
                "Parent Group": pg,
                "Dependent Group": dg,
                "Type": r['Type'],
                "Multiple": float(r['Multiple']) if pd.notna(r['Multiple']) else 1.0,
                "Dependent IA": dependent_sums[dg]['IA'],
                "Dependent FA": dependent_sums[dg]['FA'],
                "Balance IA": r['Balance IA'],
                "Balance FA": r['Balance FA']
            })
    return pd.DataFrame(rows)

def build_network_dot(flat_df):
    dot = ["digraph G {", "rankdir=LR;"]
    for pg in flat_df['Parent Group'].unique():
        dot.append(f'"{pg}" [shape=box, style=filled, fillcolor=lightblue];')
    for dg in flat_df['Dependent Group'].unique():
        dot.append(f'"{dg}" [shape=ellipse, style=filled, fillcolor=lightgreen];')
    for _, r in flat_df.iterrows():
        style = "solid" if r['Type']=="Objective" else "dashed"
        color = "blue"   if r['Type']=="Objective" else "gray"
        dot.append(
          f'"{r["Parent Group"]}" -> "{r["Dependent Group"]}" '
          f'[label="{r["Multiple"]}", style="{style}", color="{color}"];'
        )
    dot.append("}")
    return "\n".join(dot)

def apply_config(config):
    """Load JSON config into session_state, then rerun."""
    st.session_state["parent_count"]    = config.get("num_parent_groups",1)
    st.session_state["dependent_count"] = config.get("num_dependent_groups",1)

    # Parents
    for i, pg in enumerate(config.get("parent_groups",[])):
        st.session_state[f"parent_name_{i}"]   = pg.get("name","")
        st.session_state[f"parent_search_{i}"] = (
            ", ".join(pg.get("search_terms",[]))
            if isinstance(pg.get("search_terms"),list)
            else pg.get("search_terms","")
        )
        st.session_state[f"parent_excl_{i}"]   = (
            ", ".join(pg.get("exclusion_terms",[]))
            if isinstance(pg.get("exclusion_terms"),list)
            else pg.get("exclusion_terms","")
        )

    # Dependents
    for j, dg in enumerate(config.get("dependent_groups",[])):
        st.session_state[f"dep_name_{j}"]   = dg.get("name","")
        st.session_state[f"dep_search_{j}"] = (
            ", ".join(dg.get("search_terms",[]))
            if isinstance(dg.get("search_terms"),list)
            else dg.get("search_terms","")
        )
        st.session_state[f"dep_excl_{j}"]   = (
            ", ".join(dg.get("exclusion_terms",[]))
            if isinstance(dg.get("exclusion_terms"),list)
            else dg.get("exclusion_terms","")
        )

    # Mappings
    for parent_name, mapping in config.get("mapping_selections",{}).items():
        for i in range(st.session_state["parent_count"]):
            if st.session_state.get(f"parent_name_{i}") == parent_name:
                st.session_state[f"mapping_name_{i}"] = mapping.get("mapping_group_name","")
                for sel in mapping.get("selected_dependents",[]):
                    dep_name = sel.get("name","")
                    for j in range(st.session_state["dependent_count"]):
                        if st.session_state.get(f"dep_name_{j}") == dep_name:
                            st.session_state[f"map_{i}_{j}"] = True
                            st.session_state[f"obj_{i}_{j}"] = sel.get("Objective Mapping", False)
                            try:
                                st.session_state[f"mult_{i}_{j}"] = float(sel.get("multiple",1.0))
                            except:
                                st.session_state[f"mult_{i}_{j}"] = 1.0
                break

    st.rerun()


# ──────────────────────────────────────────────────────────────────────────────
# Helpers for “Delete Group” buttons — must live at top of file, before Configuration
# ──────────────────────────────────────────────────────────────────────────────
def delete_parent_group(idx):
    """Remove parent #idx (and all its mapping keys), shift everything above it down."""
    sc = st.session_state
    pc = sc.get("parent_count", 0)
    dc = sc.get("dependent_count", 0)
    # 1) delete the keys for this exact index
    for key in (
        f"parent_name_{idx}",
        f"parent_search_{idx}",
        f"parent_excl_{idx}",
        f"mapping_name_{idx}"
    ):
        sc.pop(key, None)
    # delete all map/obj/mult for this parent
    for j in range(dc):
        for prefix in ("map", "obj", "mult"):
            sc.pop(f"{prefix}_{idx}_{j}", None)
    # 2) shift every parent i>idx down to i-1
    for i in range(idx+1, pc):
        # shift name/search/excl/mapping_name
        for base in ("parent_name", "parent_search", "parent_excl", "mapping_name"):
            old = f"{base}_{i}"
            new = f"{base}_{i-1}"
            if old in sc:
                sc[new] = sc.pop(old)
        # shift map/obj/mult for each dependent
        for j in range(dc):
            for prefix in ("map", "obj", "mult"):
                old = f"{prefix}_{i}_{j}"
                new = f"{prefix}_{i-1}_{j}"
                if old in sc:
                    sc[new] = sc.pop(old)
    # 3) decrement the count
    sc["parent_count"] = pc - 1

def delete_dependent_group(idx):
    """Remove dependent #idx (and all its mapping keys), shift everything above it down."""
    sc = st.session_state
    pc = sc.get("parent_count", 0)
    dc = sc.get("dependent_count", 0)
    # 1) delete the keys for this exact index
    for key in (
        f"dep_name_{idx}",
        f"dep_search_{idx}",
        f"dep_excl_{idx}"
    ):
        sc.pop(key, None)
    # delete all map/obj/mult for this dependent
    for i in range(pc):
        for prefix in ("map", "obj", "mult"):
            sc.pop(f"{prefix}_{i}_{idx}", None)
    # 2) shift every dependent j>idx down to j-1
    for j in range(idx+1, dc):
        # shift name/search/excl
        for base in ("dep_name", "dep_search", "dep_excl"):
            old = f"{base}_{j}"
            new = f"{base}_{j-1}"
            if old in sc:
                sc[new] = sc.pop(old)
        # shift map/obj/mult for each parent
        for i in range(pc):
            for prefix in ("map", "obj", "mult"):
                old = f"{prefix}_{i}_{j}"
                new = f"{prefix}_{i}_{j-1}"
                if old in sc:
                    sc[new] = sc.pop(old)
    # 3) decrement the count
    sc["dependent_count"] = dc - 1

# ---------------------------
# Main App
# ---------------------------

st.set_page_config(page_title="Product & Dependent Mapper", layout="wide")
st.sidebar.title("Navigation")

# Sidebar: data upload
uploaded_file = st.sidebar.file_uploader("Upload your data file", type=["xlsx","xls","csv"])

# Sidebar: config import
config_upload = st.sidebar.file_uploader("Upload config JSON (optional)", type=["json"])
if config_upload is not None:
    try:
        cfg = json.load(config_upload)
        # st.sidebar.success("Config JSON loaded.")
        if st.sidebar.button("Apply Config to Form"):
            apply_config(cfg)
    except Exception as e:
        st.sidebar.error(f"Invalid JSON: {e}")
# ───────────────────────────────────────────────────────────────
# Sidebar: Clear Form button (only when the form is “dirty”)
# ───────────────────────────────────────────────────────────────
def _form_is_filled():
    for k in st.session_state.keys():
        if k.startswith((
            "parent_count","dependent_count",
            "parent_name_","parent_search_","parent_excl_",
            "dep_name_","dep_search_","dep_excl_",
            "mapping_name_","map_","obj_","mult_"
        )):
            # if any of those keys exist & are not default
            return True
    return False

if _form_is_filled():
    if st.sidebar.button("🔄 Clear Form"):
        # delete all form‐related keys
        for k in list(st.session_state.keys()):
            if k.startswith((
                "parent_count","dependent_count",
                "parent_name_","parent_search_","parent_excl_",
                "dep_name_","dep_search_","dep_excl_",
                "mapping_name_","map_","obj_","mult_"
            )) or k in ("last_config","config_prefilled","report_ready"):
                del st.session_state[k]
        # force a rerun so the form empties out
        st.rerun()
# Page selector
page = st.sidebar.radio("Go to", ["Configuration","Report"])

# Ensure flag
if "report_ready" not in st.session_state:
    st.session_state.report_ready = False

# === Configuration Page ===
if page == "Configuration":
    st.header("1) Configuration")
    if not uploaded_file:
        st.warning("Please upload a data file in the sidebar first.")
        st.stop()

    df = read_excel(uploaded_file)
    if df.empty:
        st.error("No data loaded from your file.")
        st.stop()

    has_ia_fa = {"IA","FA"}.issubset(df.columns)
    st.session_state["has_ia_fa"] = has_ia_fa
    if not has_ia_fa:
        st.warning("⚠️ No IA/FA columns detected – shortage balances and full-Excel export will be disabled.")

    # ─────────────────────────────────────────────
    # 0) Prefill the form from last‐generated config
    # ─────────────────────────────────────────────
    if "last_config" in st.session_state and not st.session_state.get("config_prefilled", False):
        # mark that we’ve applied it (so we don’t loop)
        st.session_state["config_prefilled"] = True
        # this will populate ALL your parent_/dep_/map_ widgets from last_config
        apply_config(st.session_state["last_config"])
        # apply_config() calls st.rerun() for us

    with st.expander("Data Preview (first 50 rows)", expanded=False):
        st.dataframe(df.head(50))




    # Now pull counts out of state
    parent_count    = st.session_state.get("parent_count", 1)
    dependent_count = st.session_state.get("dependent_count", 1)



    # Parent Groups
    st.subheader("Define Parent Groups")
    parent_list = []
    # for i in range(parent_count):
    #     with st.expander(f"Parent Group #{i+1}", expanded=(i==0)):

    for i in range(parent_count):
        # read the live name from session_state, fallback to P_i+1
        exp_label = st.session_state.get(f"parent_name_{i}", f"P_{i+1}")
        with st.expander(exp_label, expanded=(i==0)):
            c0 = st.columns([1,5])[0]
            # delete button at top right of the expander
            if c0.button("🗑️ Delete Group", key=f"delete_parent_{i}"):
                delete_parent_group(i)
                st.rerun()
            c1,c2,c3 = st.columns(3)
            name       = c1.text_input(
                "Name",
                value=st.session_state.get(f"parent_name_{i}",f"P_{i+1}"),
                key=f"parent_name_{i}"
            )
            search_str = c2.text_input(
                "Search Terms (comma-separated)",
                value=st.session_state.get(f"parent_search_{i}",""),
                key=f"parent_search_{i}"
            )
            excl_str   = c3.text_input(
                "Exclusion Terms (comma-separated)",
                value=st.session_state.get(f"parent_excl_{i}",""),
                key=f"parent_excl_{i}"
            )
            search = [s.strip() for s in search_str.split(",") if s.strip()]
            excl   = [s.strip() for s in excl_str.split(",") if s.strip()]
            if not search:
                st.warning("❗ No search terms: this group will return zero rows.")
            parent_list.append({"name":name.strip(), "search":search, "excl":excl})


    # moved Add-Parent button here
    if st.button("➕ Add Product Group"):
        st.session_state.parent_count = parent_count + 1
        idx = parent_count
        st.session_state[f"parent_name_{idx}"]   = f"P_{idx+1}"
        st.session_state[f"parent_search_{idx}"] = ""
        st.session_state[f"parent_excl_{idx}"]   = ""
        st.rerun()

    # Dependent Groups
    st.subheader("Define Dependent Groups")
    dependent_list = []
    # for j in range(dep_count):
    #     with st.expander(f"Dependent Group #{j+1}", expanded=(j==0)):

    for j in range(dependent_count):
        exp_label = st.session_state.get(f"dep_name_{j}", f"D_{j+1}")
        with st.expander(exp_label, expanded=(j==0)):
            c0 = st.columns([1,5])[0]
            if c0.button("🗑️ Delete Group", key=f"delete_dependent_{j}"):
                delete_dependent_group(j)
                st.rerun()
            c1,c2,c3 = st.columns(3)
            name       = c1.text_input(
                "Name",
                value=st.session_state.get(f"dep_name_{j}",f"D_{j+1}"),
                key=f"dep_name_{j}"
            )
            search_str = c2.text_input(
                "Search Terms (comma-separated)",
                value=st.session_state.get(f"dep_search_{j}",""),
                key=f"dep_search_{j}"
            )
            excl_str   = c3.text_input(
                "Exclusion Terms (comma-separated)",
                value=st.session_state.get(f"dep_excl_{j}",""),
                key=f"dep_excl_{j}"
            )
            search = [s.strip() for s in search_str.split(",") if s.strip()]
            excl   = [s.strip() for s in excl_str.split(",") if s.strip()]
            if not search:
                st.warning("❗ No search terms: this group will return zero rows.")
            dependent_list.append({"name":name.strip(), "search":search, "excl":excl})


    # moved Add-Dependent button here
    if st.button("➕ Add Dependent Group"):
        st.session_state.dependent_count = dependent_count + 1
        idx = dependent_count
        st.session_state[f"dep_name_{idx}"]   = f"D_{idx+1}"
        st.session_state[f"dep_search_{idx}"] = ""
        st.session_state[f"dep_excl_{idx}"]   = ""
        st.rerun()

    # Mappings
    st.subheader("Define Mappings")
    mapping_tabs = st.tabs([p["name"] or f"P_{i+1}" for i,p in enumerate(parent_list)])
    mapping_selections = []
    for i, tab in enumerate(mapping_tabs):
        with tab:
            pg_name = parent_list[i]["name"]
            mg = st.text_input(
                f"Mapping Group Name for '{pg_name}'",
                value=st.session_state.get(f"mapping_name_{i}",f"{pg_name} Mapping"),
                key=f"mapping_name_{i}"
            )
            st.markdown(f"**Map Dependents → {pg_name}**")
            sel_list = []
            for j, d in enumerate(dependent_list):
                c1,c2,c3 = st.columns([1,1,1])
                with c1:
                    m = st.checkbox(
                        d["name"],
                        value=st.session_state.get(f"map_{i}_{j}", False),
                        key=f"map_{i}_{j}"
                    )
                with c2:
                    obj = st.checkbox(
                        "Objective",
                        value=st.session_state.get(f"obj_{i}_{j}", False),
                        key=f"obj_{i}_{j}"
                    )
                with c3:
                    raw = st.session_state.get(f"mult_{i}_{j}", 1.0)
                    try:
                        default_mult = float(raw)
                    except:
                        default_mult = 1.0
                    mult = st.number_input(
                        "Multiple",
                        min_value=0.1, max_value=100.0,
                        step=0.1,
                        value=default_mult,
                        key=f"mult_{i}_{j}"
                    )
                sel_list.append({
                    "dep_idx": j,
                    "selected": m,
                    "objective": obj,
                    "multiple": mult
                })
            mapping_selections.append({
                "parent_idx": i,
                "mapping_name": mg,
                "selections": sel_list
            })

    # ─────────────────────────────────────────────
    # Build a dict of the current form configuration
    # ─────────────────────────────────────────────
    cur_config = {
        "num_parent_groups":    parent_count,
        "num_dependent_groups": dependent_count,
        "parent_groups": [
            {"name": p["name"], 
             "search_terms": p["search"], 
             "exclusion_terms": p["excl"]}
            for p in parent_list
        ],
        "dependent_groups": [
            {"name": d["name"], 
             "search_terms": d["search"], 
             "exclusion_terms": d["excl"]}
            for d in dependent_list
        ],
        "mapping_selections": {
            parent_list[blk["parent_idx"]]["name"]: {
                "mapping_group_name": blk["mapping_name"],
                "selected_dependents": [
                    {
                        "name": dependent_list[sel["dep_idx"]]["name"],
                        "Objective Mapping": sel["objective"],
                        "multiple": sel["multiple"]
                    }
                    for sel in blk["selections"] if sel["selected"]
                ]
            }
            for blk in mapping_selections
        }
    }


    # ─────────────────────────────────────────────────────────
    # Sidebar “Generate Report” button + status messaging
    # ─────────────────────────────────────────────────────────
    gen_clicked = st.sidebar.button("Generate Report")

    if gen_clicked:
        # # Generate Report
            # Process parent groups
            parent_dfs, parent_sums, parent_names = {}, {}, []
            for p in parent_list:
                parent_names.append(p["name"])
                dfp, sp = filter_and_sum(df, p["search"], p["excl"])
                parent_dfs[p["name"]] = dfp
                parent_sums[p["name"]] = sp

            # Process dependent groups
            dependent_dfs, dependent_sums, dependent_names = {}, {}, []
            for d in dependent_list:
                dependent_names.append(d["name"])
                dfd, sd = filter_and_sum(df, d["search"], d["excl"])
                dependent_dfs[d["name"]] = dfd
                dependent_sums[d["name"]] = sd

            # Overlap Analysis
            parent_models = {name: set(df['Description']) for name, df in parent_dfs.items() if not df.empty}
            dependent_models = {name: set(df['Description']) for name, df in dependent_dfs.items() if not df.empty}

            overlapping_models = {}
            for pname, pmodels in parent_models.items():
                for dname, dmodels in dependent_models.items():
                    intersection = pmodels.intersection(dmodels)
                    if intersection:
                        if pname not in overlapping_models:
                            overlapping_models[pname] = {}
                        overlapping_models[pname][dname] = intersection
            
            overlapping_parent_groups = set(overlapping_models.keys())
            parent_only_groups = set(parent_names) - overlapping_parent_groups

            # Process mappings
            mapping_to_parent = {}
            parent_to_mappings  = {}
            dependent_to_mappings= {}
            mappings_data       = {}
            for block in mapping_selections:
                i  = block["parent_idx"]
                pg = parent_list[i]["name"]
                mg = block["mapping_name"]
                entries = []
                for sel in block["selections"]:
                    if sel["selected"]:
                        dg = dependent_list[sel["dep_idx"]]["name"]
                        mult = sel["multiple"]
                        try:
                            bal_ia = calculate_balances(parent_sums[pg], dependent_sums[dg], 'IA', mult)
                            bal_fa = calculate_balances(parent_sums[pg], dependent_sums[dg], 'FA', mult)
                        except:
                            bal_ia = bal_fa = 0.0
                        entries.append({
                            "Dependent Group": dg,
                            "Type": "Objective" if sel["objective"] else "Subjective",
                            "Multiple": mult,
                            "Balance IA": bal_ia,
                            "Balance FA": bal_fa
                        })
                        dependent_to_mappings.setdefault(dg, []).append(pg)

                mappings_data[mg] = pd.DataFrame(entries)
                mapping_to_parent[mg] = pg
                parent_to_mappings.setdefault(pg, []).append(mg)

            # Safe sheet names
            existing = {"Dependent Perspective","All Mappings"}
            parent_sheet_names = {}
            for pn in parent_names:
                sn = get_safe_sheet_name(pn, existing)
                parent_sheet_names[pn] = sn
                existing.add(sn)
            dependent_sheet_names = {}
            for dn in dependent_names:
                sn = get_safe_sheet_name(dn, existing)
                dependent_sheet_names[dn] = sn
                existing.add(sn)

            # ———————— only build Excel if IA/FA are present ————————
            if st.session_state.get("has_ia_fa", False):
                excel_bytes = save_to_excel_bytes(
                    mappings_data, mapping_to_parent, dependent_to_mappings,
                    parent_dfs, dependent_dfs,
                    parent_names, dependent_names,
                    parent_to_mappings,
                    parent_sheet_names, dependent_sheet_names,
                    parent_sums, dependent_sums
                )
            else:
                excel_bytes = None


            # Persist to session_state
            st.session_state.update({
                "report_ready": True,
                "df": df,
                "parent_names": parent_names,
                "dependent_names": dependent_names,
                "parent_dfs": parent_dfs,
                "dependent_dfs": dependent_dfs,
                "parent_sums": parent_sums,
                "dependent_sums": dependent_sums,
                "mappings_data": mappings_data,
                "mapping_to_parent": mapping_to_parent,
                "parent_to_mappings": parent_to_mappings,
                "dependent_to_mappings": dependent_to_mappings,
                "parent_sheet_names": parent_sheet_names,
                "dependent_sheet_names": dependent_sheet_names,
                "excel_bytes": excel_bytes,
                "cfg_parent_list": parent_list,
                "cfg_dependent_list": dependent_list,
                "cfg_mapping_selections": mapping_selections,
                "overlapping_models": overlapping_models,
                "overlapping_parent_groups": overlapping_parent_groups,
                "parent_only_groups": parent_only_groups
            })


    # Always show *one* message in the sidebar:
    if gen_clicked:
        if st.session_state["has_ia_fa"]:
            st.sidebar.success("✅ Report + Excel generated! Switch to Report page.")
        else:
            st.sidebar.success("✅ Report generated! (Excel export disabled.)")
    elif "last_config" in st.session_state and cur_config != st.session_state["last_config"]:
        st.sidebar.warning("⚠ You have unsaved changes. Click **Generate Report** to refresh the report with your edits.")



# === Report Page ===
elif page == "Report":
    st.header("2) Report")
    if not st.session_state.report_ready:
        st.warning("Generate the report first on the Configuration page.")
        st.stop()


    # ─────────────────────────────────────────────────
    # REHYDRATE last_config from the saved cfg_* lists
    # ─────────────────────────────────────────────────
    if (
        "cfg_parent_list" in st.session_state
        and "cfg_dependent_list" in st.session_state
        and "cfg_mapping_selections" in st.session_state
    ):
        p_list  = st.session_state["cfg_parent_list"]
        d_list  = st.session_state["cfg_dependent_list"]
        map_sel = st.session_state["cfg_mapping_selections"]

        # build a fresh config dict exactly matching your apply_config format
        new_cfg = {
            "num_parent_groups":    len(p_list),
            "num_dependent_groups": len(d_list),
            "parent_groups": [
                {
                    "name": pg["name"],
                    "search_terms": pg["search"],
                    "exclusion_terms": pg["excl"],
                }
                for pg in p_list
            ],
            "dependent_groups": [
                {
                    "name": dg["name"],
                    "search_terms": dg["search"],
                    "exclusion_terms": dg["excl"],
                }
                for dg in d_list
            ],
            "mapping_selections": {
                # key = actual parent_name
                p_list[blk["parent_idx"]]["name"]: {
                    "mapping_group_name": blk["mapping_name"],
                    "selected_dependents": [
                        {
                            "name": d_list[sel["dep_idx"]]["name"],
                            "Objective Mapping": sel["objective"],
                            "multiple": sel["multiple"],
                        }
                        for sel in blk["selections"] if sel["selected"]
                    ],
                }
                for blk in map_sel
            },
        }

        # now replace last_config so that Configuration→prefill works
        st.session_state["last_config"]     = new_cfg
        st.session_state["config_prefilled"] = False

    # Retrieve
    parent_names      = st.session_state.parent_names
    dependent_names   = st.session_state.dependent_names
    parent_dfs        = st.session_state.parent_dfs
    dependent_dfs     = st.session_state.dependent_dfs
    parent_sums       = st.session_state.parent_sums
    dependent_sums    = st.session_state.dependent_sums
    mappings_data     = st.session_state.mappings_data
    mapping_to_parent = st.session_state.mapping_to_parent
    excel_bytes       = st.session_state.excel_bytes
    overlapping_models = st.session_state.overlapping_models
    overlapping_parent_groups = st.session_state.overlapping_parent_groups
    parent_only_groups = st.session_state.parent_only_groups

    # ———————— conditional Excel download ————————
    if st.session_state.get("has_ia_fa", False) and excel_bytes:
        st.download_button(
            "Download Full Excel Report",
            data=excel_bytes,
            file_name=f"Product_Dependent_Report_{datetime.now():%Y-%m-%d}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("ℹ️ Skipping full-Excel export (no IA/FA in source data).")

    # Flatten for visuals
    flat = flatten_mappings(mappings_data, mapping_to_parent, dependent_sums)


    # Network Graph with Selections
    st.subheader("Parent ↔ Dependent Network")
    # Product selection per group
    parent_selections = {}
    for pn in parent_names:
        with st.expander(f"Products for {pn}", expanded=False):
            all_opts = parent_dfs[pn]['Description'].tolist()
            select_all = st.checkbox("Select All", value=True, key=f"select_all_{pn}")
            default_opts = all_opts if select_all else []
            parent_selections[pn] = st.multiselect(
                f"Select products for {pn}",
                options=all_opts,
                default=default_opts,
                key=f"select_{pn}"
            )
    dependent_selections = {}
    for dn in dependent_names:
        with st.expander(f"Products for {dn}", expanded=False):
            all_opts = dependent_dfs[dn]['Description'].tolist()
            select_all = st.checkbox("Select All", value=True, key=f"select_all_{dn}")
            default_opts = all_opts if select_all else []
            dependent_selections[dn] = st.multiselect(
                f"Select products for {dn}",
                options=all_opts,
                default=default_opts,
                key=f"select_{dn}"
            )
    # Build network graph with dynamic labels and edge types
    # Build network graph with titled clusters
    # Create a flat set of all overlapping models for easy filtering
    all_overlapping_flat = set()
    if overlapping_models:
        for pname, dmap in overlapping_models.items():
            for dname, models in dmap.items():
                all_overlapping_flat.update(models)

    dot_lines = [
        "digraph G {",
        "  rankdir=LR;",
        "  node [fontname=helvetica];",
        "  edge [fontname=helvetica];",
        "  compound=true;",
        "",
        # 1) Parent cluster (only for non-overlapping parents)
        "  subgraph cluster_parents {",
        '    label="Parent Groups";',
        "    style=filled;",
        "    color=lightgrey;",
        "    node [style=filled, shape=box, fillcolor=lightblue];",
    ]
    for pn in parent_names:
        if pn in overlapping_parent_groups:
            continue # Skip overlapping parents, they are rendered as hybrid nodes
        selected_for_pn = set(parent_selections.get(pn, []))
        display_models = sorted(list(selected_for_pn))
        items = "<BR/>".join(display_models) if display_models else ""
        html_label = f"<B>{pn}</B>{'<BR/>' + items if items else ''}"
        dot_lines.append(f'    "{pn}" [label=<{html_label}>];')
    dot_lines.append("  }")

    dot_lines.extend([
        "",
        # 2) Dependent cluster
        "  subgraph cluster_dependents {",
        '    label="Dependent Groups";',
        "    style=filled; color=whitesmoke; node [style=filled];",
    ])

    # Create a container for each dependent group
    for dn in dependent_names:
        dot_lines.extend([
            f'  subgraph "cluster_{dn}" {{',
            f'    label="{dn}";',
            f'    style="filled"; color="lightgreen";',
        ])

        # Node for non-overlapping models within the dependent group
        selected_for_dn = set(dependent_selections.get(dn, []))
        display_models_dn = sorted(list(selected_for_dn - all_overlapping_flat))
        items_dn = "<BR/>".join(display_models_dn) if display_models_dn else ""
        if not items_dn and any(dn in d for p, d in overlapping_models.items()):
             html_label_dn = f'<B>{dn}</B><BR/>(unique models)'
        else:
            html_label_dn = f"<B>{dn}</B>{'<BR/>' + items_dn if items_dn else ''}"
        dot_lines.append(f'    "{dn}" [shape=ellipse, fillcolor=lightgreen, label=<{html_label_dn}>];')

        # Create hybrid boxes for overlaps INSIDE this dependent group
        for opg in overlapping_parent_groups:
            if opg in overlapping_models and dn in overlapping_models[opg]:
                intersecting_models = overlapping_models[opg][dn]
                selected_parent_products = set(parent_selections.get(opg, []))
                display_models_hybrid = sorted(list(intersecting_models.intersection(selected_parent_products)))
                
                if display_models_hybrid:
                    hybrid_name = f"hybrid_{opg}_{dn}"
                    items_hybrid = "<BR/>".join(display_models_hybrid)
                    html_label_hybrid = f"<B>{opg} Models</B><BR/>{items_hybrid}"
                    dot_lines.append(
                        f'    "{hybrid_name}" [shape=box, style=filled, fillcolor="tomato", label=<{html_label_hybrid}>];'
                    )
        dot_lines.append("  }")

    dot_lines.append("  }")

    # 3) Edges
    for _, r in flat.iterrows():
        pg = r["Parent Group"]
        dg = r["Dependent Group"]
        style = "solid" if r["Type"] == "Objective" else "dashed"
        color = "blue" if r["Type"] == "Objective" else "gray"
        label = f'{r["Type"]} ({r["Multiple"]})'
        
        if pg not in overlapping_parent_groups:
            # Edge from a normal parent
            dot_lines.append(f'  "{pg}" -> "{dg}" [label="{label}", style="{style}", color="{color}"];')
        else:
            # Edge from a hybrid parent
            # The connection originates from the hybrid box inside the dependent it overlaps with
            if pg in overlapping_models:
                for d_overlap in overlapping_models[pg].keys():
                    hybrid_name = f"hybrid_{pg}_{d_overlap}"
                    dot_lines.append(f'  "{hybrid_name}" -> "{dg}" [label="{label}", style="{style}", color="{color}"];')

    dot_lines.append("}")
    dot = "\n".join(dot_lines)

    st.graphviz_chart(dot, use_container_width=True)


    # Raw Data Explorer
    st.subheader("Raw Data Explorer")
    tabs = st.tabs(["Parent Groups","Dependent Groups","Mappings"])
    with tabs[0]:
        for pn in parent_names:
            with st.expander(pn):
                st.dataframe(parent_dfs[pn])
                s = parent_sums[pn]
                st.write(f"**Totals:** IA {int(s['IA'])}  FA {int(s['FA'])}")
    with tabs[1]:
        for dn in dependent_names:
            with st.expander(dn):
                st.dataframe(dependent_dfs[dn])
                s = dependent_sums[dn]
                st.write(f"**Totals:** IA {int(s['IA'])}  FA {int(s['FA'])}")
    with tabs[2]:
        st.dataframe(flat)

    # ---- FIXED: Download Config JSON ----
    st.subheader("Export Current Configuration")
    if st.button("Download Config JSON"):
        p_list = st.session_state.cfg_parent_list
        d_list = st.session_state.cfg_dependent_list
        mapping_sel = st.session_state.cfg_mapping_selections

        cfg = {
            "num_parent_groups":    len(p_list),
            "num_dependent_groups": len(d_list),
            "parent_groups":        [],
            "dependent_groups":     [],
            "mapping_selections":   {}
        }

        for pg in p_list:
            cfg["parent_groups"].append({
                "name": pg["name"],
                "search_terms": pg["search"],
                "exclusion_terms": pg["excl"]
            })

        for dg in d_list:
            cfg["dependent_groups"].append({
                "name": dg["name"],
                "search_terms": dg["search"],
                "exclusion_terms": dg["excl"]
            })

        for block in mapping_sel:
            i = block["parent_idx"]
            parent_name = p_list[i]["name"]
            mg_name     = block["mapping_name"]
            selected_deps = []
            for sel in block["selections"]:
                if sel["selected"]:
                    j = sel["dep_idx"]
                    selected_deps.append({
                        "name": d_list[j]["name"],
                        "Objective Mapping": sel["objective"],
                        "multiple": sel["multiple"]
                    })
            cfg["mapping_selections"][parent_name] = {
                "mapping_group_name": mg_name,
                "selected_dependents": selected_deps
            }

        js = json.dumps(cfg, indent=4)
        st.download_button(
            label="🔽 Save Config JSON",
            data=js,
            file_name="config.json",
            mime="application/json"
        )

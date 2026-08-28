import json
import streamlit as st
import pandas as pd
from discount_engine import (
    list_platforms, parse_report, get_master_sheet_names, get_master_headers,
    parse_master_file, build_working_sheet, build_mismatch_sheet,
    build_upload_sheet, write_output_workbook, parse_exclusion_skus,
    format_for_display,
)

st.set_page_config(page_title="Discount Promo Builder", layout="wide")
st.title("Marketplace Discount Promotion Upload File Builder")
st.caption("Report + Masterfile -> working sheet + RRP mismatch tab + ready-to-upload tier files. Works for any account (EWG, DBC, FYW, GSK...) and any supported marketplace.")

if "tiers" not in st.session_state:
    st.session_state.tiers = [{"name": "BAU", "col": "F"}]

# ---------------------------------------------------------------------------
# Step 0: Platform + Account + saved mapping presets
# ---------------------------------------------------------------------------
st.header("0. Platform & account")

col_a, col_b = st.columns(2)
with col_a:
    platform = st.selectbox("Marketplace", list_platforms())
with col_b:
    account_name = st.text_input("Account name (used for the output filename, e.g. EWG, DBC, FYW, GSK)", value="")

with st.expander("Load a saved column mapping (from a previous run for this account + platform)"):
    preset_file = st.file_uploader("Upload mapping .json", type=["json"], key="preset_upload")
    if preset_file:
        try:
            preset = json.loads(preset_file.getvalue())
            st.session_state.tiers = preset["tiers"]
            st.session_state["_preset_sheet_name"] = preset.get("sheet_name")
            st.session_state["_preset_sku_col"] = preset.get("sku_col")
            st.session_state["_preset_rrp_col"] = preset.get("rrp_col")
            st.success(f"Loaded mapping: {preset.get('sheet_name')}, SKU={preset.get('sku_col')}, RRP={preset.get('rrp_col')}, tiers={preset['tiers']}. Re-select these below if they didn't auto-fill.")
        except Exception as e:
            st.error(f"Could not read that mapping file: {e}")

# ---------------------------------------------------------------------------
# Step 1: Marketplace report
# ---------------------------------------------------------------------------
st.header(f"1. {platform} report")
report_file = st.file_uploader(f"Upload {platform} report (.xlsx)", type=["xlsx"], key="report_file")

report_df = None
if report_file:
    try:
        report_df = parse_report(report_file.getvalue(), platform)
        st.success(f"Parsed {len(report_df)} products from the {platform} report.")
        with st.expander(f"Preview {platform} report"):
            st.dataframe(report_df.head(20), use_container_width=True)
    except Exception as e:
        st.error(f"Could not parse the {platform} report: {e}")

# ---------------------------------------------------------------------------
# Step 2: Masterfile
# ---------------------------------------------------------------------------
st.header("2. Masterfile")
master_file = st.file_uploader("Upload Masterfile (.xlsx)", type=["xlsx"], key="master_file")

master_df = None
tiers = {}
if master_file:
    try:
        sheet_names = get_master_sheet_names(master_file.getvalue())
        preset_sheet = st.session_state.get("_preset_sheet_name")
        sheet_idx = sheet_names.index(preset_sheet) if preset_sheet in sheet_names else 0
        sheet_name = st.selectbox("Which tab to refer to?", sheet_names, index=sheet_idx)

        headers = get_master_headers(master_file.getvalue(), sheet_name)
        header_labels = [f"{letter} - {val}" for letter, val in headers]
        letter_only = [letter for letter, _ in headers]

        preset_sku = st.session_state.get("_preset_sku_col")
        preset_rrp = st.session_state.get("_preset_rrp_col")
        default_sku_idx = letter_only.index(preset_sku) if preset_sku in letter_only else 0
        default_rrp_idx = letter_only.index(preset_rrp) if preset_rrp in letter_only else min(4, len(headers) - 1)

        col1, col2 = st.columns(2)
        with col1:
            sku_idx = st.selectbox("SKU column", range(len(headers)),
                                    format_func=lambda i: header_labels[i], key="sku_col_sel",
                                    index=default_sku_idx)
        with col2:
            rrp_idx = st.selectbox("RRP column", range(len(headers)),
                                    format_func=lambda i: header_labels[i], key="rrp_col_sel",
                                    index=default_rrp_idx)

        sku_col = letter_only[sku_idx]
        rrp_col = letter_only[rrp_idx]

        st.subheader("Discount tiers")
        st.caption("Add one row per tier (e.g. BAU, D-day, SWFS, 9.9). Pick the masterfile column that holds the SRP for each.")

        for i, tier in enumerate(st.session_state.tiers):
            c1, c2, c3 = st.columns([2, 2, 1])
            with c1:
                tier["name"] = st.text_input(f"Tier {i+1} name", value=tier["name"], key=f"tier_name_{i}")
            with c2:
                default_idx = letter_only.index(tier["col"]) if tier["col"] in letter_only else 0
                sel = st.selectbox(f"Tier {i+1} column", range(len(headers)),
                                    format_func=lambda i: header_labels[i], key=f"tier_col_{i}",
                                    index=default_idx)
                tier["col"] = letter_only[sel]
            with c3:
                st.write("")
                st.write("")
                if st.button("Remove", key=f"remove_tier_{i}") and len(st.session_state.tiers) > 1:
                    st.session_state.tiers.pop(i)
                    st.rerun()

        if st.button("+ Add another tier"):
            st.session_state.tiers.append({"name": f"TIER{len(st.session_state.tiers)+1}", "col": "G"})
            st.rerun()

        tiers = {t["name"]: t["col"] for t in st.session_state.tiers if t["name"].strip()}
        master_df = parse_master_file(master_file.getvalue(), sheet_name, sku_col, rrp_col, tiers)

        with st.expander("Preview Masterfile mapping"):
            st.dataframe(master_df.head(20), use_container_width=True)

        preset_json = json.dumps({
            "platform": platform,
            "sheet_name": sheet_name,
            "sku_col": sku_col,
            "rrp_col": rrp_col,
            "tiers": st.session_state.tiers,
        }, indent=2)
        st.download_button(
            f"Save this column mapping for {account_name or 'this account'} / {platform} (.json)",
            data=preset_json,
            file_name=f"{account_name or 'account'}_{platform}_discount_mapping.json",
            mime="application/json",
        )
        st.caption("Next time you run this account + platform, upload that .json in step 0 to skip re-picking columns.")

    except Exception as e:
        st.error(f"Could not parse Masterfile: {e}")

# ---------------------------------------------------------------------------
# Step 2.5: Exclusions
# ---------------------------------------------------------------------------
st.header("2.5. Exclude SKUs (optional)")
st.caption("Excluded SKUs stay in the working sheet (marked, for visibility) but are left out of every TO UPLOAD tab.")

exc_col1, exc_col2 = st.columns(2)
with exc_col1:
    exclusion_text = st.text_area(
        "Type or paste SKUs to exclude",
        placeholder="e.g. 4975479497216, 4975479494895\nor one per line",
        key="exclusion_text",
    )
with exc_col2:
    exclusion_file = st.file_uploader("Or upload a file of SKUs to exclude (.xlsx or .csv)",
                                       type=["xlsx", "csv"], key="exclusion_file")

excluded_skus = parse_exclusion_skus(
    manual_text=exclusion_text or "",
    file_bytes=exclusion_file.getvalue() if exclusion_file else None,
    file_name=exclusion_file.name if exclusion_file else "",
)
if excluded_skus:
    st.info(f"{len(excluded_skus)} SKU(s) will be excluded: {', '.join(sorted(excluded_skus)[:20])}"
            + (" ..." if len(excluded_skus) > 20 else ""))

# ---------------------------------------------------------------------------
# Step 3: Generate
# ---------------------------------------------------------------------------
st.header("3. Generate")

if report_df is not None and master_df is not None:
    tier_names = list(tiers.keys())
    if st.button("Build working sheet", type="primary"):
        working = build_working_sheet(report_df, master_df, tier_names, platform, excluded_skus)
        mismatch = build_mismatch_sheet(working)
        st.session_state.working = working
        st.session_state.mismatch = mismatch
        st.session_state.tier_names = tier_names
        st.session_state.built_platform = platform

    if "working" in st.session_state and st.session_state.get("built_platform") == platform:
        working = st.session_state.working
        mismatch = st.session_state.mismatch
        tier_names = st.session_state.tier_names

        st.subheader("Working sheet")
        st.dataframe(format_for_display(working, tier_names), use_container_width=True)

        st.subheader(f"RRP mismatch ({len(mismatch)} rows)")
        if len(mismatch):
            st.warning(f"{len(mismatch)} product(s) have a Masterfile RRP that doesn't match the live {platform} price. Review before uploading.")
            st.dataframe(format_for_display(mismatch, tier_names), use_container_width=True)
        else:
            st.success("No RRP mismatches found.")

        for tier in tier_names:
            upload_df = build_upload_sheet(working, tier, platform)
            with st.expander(f"TO UPLOAD {tier} ({len(upload_df)} rows)"):
                st.dataframe(upload_df, use_container_width=True)

        out_bytes = write_output_workbook(working, mismatch, tier_names, working, platform)
        out_name = f"{account_name.strip() or 'Account'}_{platform}_Discount_Promo_Upload.xlsx".replace(" ", "_")
        st.download_button(
            "Download full workbook (.xlsx)",
            data=out_bytes,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info(f"Upload both the {platform} report and the Masterfile to continue.")

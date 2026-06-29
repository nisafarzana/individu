# =============================================================================
# Digital Pavement Condition Evaluation and Maintenance Decision Tool
# TCG633 Bridge & Road Maintenance — Individual Project
# Universiti Teknologi MARA, Cawangan Sarawak
#
# Computation logic (PCI weighting/severity factors, classification bands,
# IRI bands, and the PCI/IRI hybrid rule) is taken directly from the
# lecturer-provided files:
#   - TCG633_PCI_IRI_Model.xlsx      (Lookup, PCI_Compute, IRI_Compute sheets)
#   - TCG633_PCI_IRI_Pro_v2.xlsx     (Lookup, Settings_Summary "Hybrid" logic)
# Any place where the spreadsheet logic does not cover a case (e.g. a defect
# type or severity not in the Lookup table) is explicitly flagged as an
# ASSUMPTION in-app (see "Methodology & Assumptions" tab) rather than silently
# guessed.
# =============================================================================

import io
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st

# Plotly is an optional, "nice-to-have" charting dependency. If the deployment
# environment fails to install it for any reason (e.g. requirements.txt not
# picked up yet on Streamlit Cloud), the app must NOT crash — it should fall
# back to Streamlit's built-in charts instead.
try:
    import plotly.express as px
    PLOTLY_OK = True
except ModuleNotFoundError:
    PLOTLY_OK = False

# openpyxl powers BOTH reading uploaded .xlsx/.xls files AND writing the Excel
# download. If it fails to import (e.g. a broken Cloud build environment),
# the app must still work for CSV upload/download — it must never crash.
try:
    import openpyxl  # noqa: F401
    OPENPYXL_OK = True
except ModuleNotFoundError:
    OPENPYXL_OK = False

# -----------------------------------------------------------------------------
# PAGE CONFIG
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Pavement Condition Evaluation & Maintenance Decision Tool",
    page_icon="🛣️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# -----------------------------------------------------------------------------
# LOOK-UP CONSTANTS  (copied from the lecturer's Lookup sheet)
# -----------------------------------------------------------------------------

# PCI defect weighting factors — Lookup!A:B
DEFECT_WEIGHTS = {
    "Longitudinal Crack": 1.0,
    "Alligator (Fatigue) Crack": 1.6,
    "Potholes": 2.2,
    "Raveling": 1.2,
    "Depression/Sag": 1.4,
    "Patching (Failed)": 1.8,
    "Bleeding/Flushing": 1.0,
    "Rut/Rutting": 1.6,
}
DEFAULT_WEIGHT = 1.0  # ASSUMPTION: used only if an unrecognised defect type is supplied

# PCI severity factors — Lookup!D:E
SEVERITY_FACTORS = {"Low": 0.6, "Medium": 1.0, "High": 1.4}
DEFAULT_SEVERITY_FACTOR = 1.0  # ASSUMPTION: "Medium" weighting used if severity is unrecognised

# PCI condition bands — Lookup!G:J  (rank 1 = best ... 4 = worst)
PCI_RANK_LABEL = {1: "Very Good", 2: "Good / Satisfactory", 3: "Fair", 4: "Poor"}
PCI_RANK_RECO = {
    1: "Routine maintenance (cleaning, grass cutting, minor touch-ups)",
    2: "Preventive maintenance (crack sealing, local patching)",
    3: "Surface treatment / Overlay (localized)",
    4: "Major rehabilitation / Reconstruction assessment",
}

# IRI condition bands — Lookup!L:O (rank 1 = best ... 4 = worst)
IRI_RANK_LABEL = {1: "Very Good (Smooth)", 2: "Good", 3: "Fair", 4: "Poor (Rough)"}
IRI_RANK_RECO = {
    1: "Routine maintenance",
    2: "Preventive maintenance (localized patching/leveling)",
    3: "Surface treatment / thin overlay",
    4: "Structural overlay / rehabilitation",
}

CONDITION_COLORS = {
    "Very Good": "#2E7D32",
    "Very Good (Smooth)": "#2E7D32",
    "Good": "#1976D2",
    "Good / Satisfactory": "#1976D2",
    "Fair": "#F9A825",
    "Poor": "#C62828",
    "Poor (Rough)": "#C62828",
}
RANK_COLOR = {1: "#2E7D32", 2: "#1976D2", 3: "#F9A825", 4: "#C62828"}

# Supplementary, defect-level treatment guide.
# NOTE / ASSUMPTION: the lecturer's spreadsheet only issues a maintenance
# recommendation at the SECTION level (based on PCI/IRI classification).
# It does not provide a per-defect-type x severity action table. The table
# below is added as general pavement-maintenance practice guidance (common,
# non-proprietary treatments) so that the tool can also recommend an action
# "for each defect" as requested. It is clearly separated from the official
# section-level recommendation in every table/report this app produces.
DEFECT_TREATMENT_GUIDE = {
    "Longitudinal Crack": {
        "Low": "Monitor; seal during routine maintenance",
        "Medium": "Crack sealing",
        "High": "Crack sealing + localized patching",
    },
    "Alligator (Fatigue) Crack": {
        "Low": "Monitor / crack sealing",
        "Medium": "Partial-depth patching",
        "High": "Full-depth patching or overlay (structural distress)",
    },
    "Potholes": {
        "Low": "Patch at next routine maintenance round",
        "Medium": "Patch promptly",
        "High": "Immediate patching — safety hazard",
    },
    "Raveling": {
        "Low": "Monitor / fog seal",
        "Medium": "Surface (chip/slurry) seal",
        "High": "Thin overlay",
    },
    "Depression/Sag": {
        "Low": "Monitor drainage and surface",
        "Medium": "Localized levelling / patching",
        "High": "Investigate sub-base; structural repair",
    },
    "Patching (Failed)": {
        "Low": "Reseal patch edges",
        "Medium": "Remove and re-patch",
        "High": "Full-depth repair of patch area",
    },
    "Bleeding/Flushing": {
        "Low": "Apply sand/blotter material",
        "Medium": "Surface treatment",
        "High": "Overlay (loss of skid resistance)",
    },
    "Rut/Rutting": {
        "Low": "Monitor",
        "Medium": "Milling and overlay",
        "High": "Structural overlay / reconstruction",
    },
}
DEFAULT_TREATMENT = "Inspect on-site and patch/repair as needed (defect type not in guide)"

CANONICAL_COLS = ["Section", "Defect Type", "Severity", "Area Percentage (%)", "IRI"]

# Map of common header variants -> canonical column name
COLUMN_ALIASES = {
    "section": "Section", "section id": "Section", "sectionid": "Section",
    "section_id": "Section", "road section": "Section",
    "defect type": "Defect Type", "defect": "Defect Type", "defecttype": "Defect Type",
    "distress type": "Defect Type", "distress": "Defect Type",
    "severity": "Severity", "severity level": "Severity",
    "area percentage (%)": "Area Percentage (%)", "area percentage": "Area Percentage (%)",
    "area (%)": "Area Percentage (%)", "area affected (%)": "Area Percentage (%)",
    "area affected": "Area Percentage (%)", "area%": "Area Percentage (%)",
    "area_pct": "Area Percentage (%)", "area": "Area Percentage (%)",
    "iri": "IRI", "iri (m/km)": "IRI", "avg iri": "IRI", "average iri": "IRI",
    "iri(m/km)": "IRI", "roughness": "IRI",
}


# -----------------------------------------------------------------------------
# HELPER FUNCTIONS
# -----------------------------------------------------------------------------
def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Rename columns to the canonical set using a case/space-insensitive match."""
    rename_map = {}
    for col in df.columns:
        key = str(col).strip().lower()
        if key in COLUMN_ALIASES:
            rename_map[col] = COLUMN_ALIASES[key]
        elif str(col).strip() in CANONICAL_COLS:
            rename_map[col] = str(col).strip()
    return df.rename(columns=rename_map)


def pci_rank(pci: float) -> int:
    if pci >= 85:
        return 1
    if pci >= 70:
        return 2
    if pci >= 55:
        return 3
    return 4


def iri_rank(iri: float) -> int:
    if iri < 2:
        return 1
    if iri < 3:
        return 2
    if iri < 4:
        return 3
    return 4


def lookup_weight(defect_type, flags):
    key = str(defect_type).strip()
    for k, v in DEFECT_WEIGHTS.items():
        if k.lower() == key.lower():
            return v
    flags.append(f"Unrecognised defect type '{defect_type}' — default weighting {DEFAULT_WEIGHT} applied")
    return DEFAULT_WEIGHT


def lookup_severity(severity, flags):
    key = str(severity).strip()
    for k, v in SEVERITY_FACTORS.items():
        if k.lower() == key.lower():
            return v
    flags.append(f"Unrecognised severity '{severity}' — default factor {DEFAULT_SEVERITY_FACTOR} (Medium) applied")
    return DEFAULT_SEVERITY_FACTOR


def lookup_treatment(defect_type, severity):
    for k, v in DEFECT_TREATMENT_GUIDE.items():
        if k.lower() == str(defect_type).strip().lower():
            for sk, sv in v.items():
                if sk.lower() == str(severity).strip().lower():
                    return sv
            return DEFAULT_TREATMENT
    return DEFAULT_TREATMENT


@st.cache_data
def generate_sample_data() -> pd.DataFrame:
    """Synthetic demo dataset covering 10 road sections (S1-S10) spanning all
    four condition classes (Very Good -> Poor), in the exact 5-column input
    format required by the assignment brief. For teaching/demo purposes only
    — not the official lecturer dataset."""
    rows = [
        # Section, Defect Type, Severity, Area %, IRI (m/km)
        ("S1", "Longitudinal Crack", "Low", 4, 1.7),
        ("S1", "Bleeding/Flushing", "Low", 2, 1.7),
        ("S2", "Alligator (Fatigue) Crack", "Medium", 8, 2.4),
        ("S2", "Longitudinal Crack", "Medium", 5, 2.4),
        ("S3", "Potholes", "Medium", 6, 2.9),
        ("S3", "Raveling", "Low", 10, 2.9),
        ("S3", "Rut/Rutting", "Low", 4, 2.9),
        ("S4", "Patching (Failed)", "High", 9, 3.6),
        ("S4", "Alligator (Fatigue) Crack", "High", 7, 3.6),
        ("S5", "Potholes", "High", 14, 4.3),
        ("S5", "Depression/Sag", "High", 10, 4.3),
        ("S5", "Rut/Rutting", "Medium", 8, 4.3),
        ("S6", "Raveling", "Medium", 6, 1.9),
        ("S6", "Longitudinal Crack", "Low", 3, 1.9),
        ("S7", "Alligator (Fatigue) Crack", "Low", 5, 2.2),
        ("S7", "Bleeding/Flushing", "Medium", 4, 2.2),
        ("S8", "Potholes", "Low", 3, 2.6),
        ("S8", "Patching (Failed)", "Medium", 4, 2.6),
        ("S9", "Raveling", "Medium", 8, 3.3),
        ("S9", "Rut/Rutting", "Medium", 10, 3.3),
        ("S9", "Bleeding/Flushing", "High", 8, 3.3),
        ("S10", "Alligator (Fatigue) Crack", "High", 12, 4.6),
        ("S10", "Patching (Failed)", "High", 10, 4.6),
    ]
    return pd.DataFrame(rows, columns=CANONICAL_COLS)


def read_uploaded_file(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        if not OPENPYXL_OK:
            raise RuntimeError(
                "This server's Python environment is missing the 'openpyxl' package, "
                "so Excel (.xlsx/.xls) files can't be read right now. "
                "Workaround: open this file in Excel/Google Sheets and re-save/export it "
                "as .csv, then upload the .csv instead — CSV upload does not need openpyxl."
            )
        df = pd.read_excel(uploaded_file, sheet_name=0)
    return normalize_columns(df)


def compute_results(df: pd.DataFrame):
    """Core engine: replicates PCI_Compute / IRI_Compute / Settings_Summary
    logic from the lecturer's Excel model. Returns (detail_df, summary_df, flags)."""
    flags = []
    df = df.copy()

    for c in CANONICAL_COLS:
        if c not in df.columns:
            df[c] = np.nan

    df["Section"] = df["Section"].astype(str).str.strip()
    df = df[df["Section"].notna() & (df["Section"] != "") & (df["Section"].str.lower() != "nan")]

    df["Area Percentage (%)"] = pd.to_numeric(df["Area Percentage (%)"], errors="coerce")
    df["IRI"] = pd.to_numeric(df["IRI"], errors="coerce")

    defect_mask = (
        df["Defect Type"].notna()
        & df["Severity"].notna()
        & df["Area Percentage (%)"].notna()
    )
    detail = df[defect_mask].copy()

    if not detail.empty:
        detail["Weighting Factor"] = detail["Defect Type"].apply(lambda d: lookup_weight(d, flags))
        detail["Severity Factor"] = detail["Severity"].apply(lambda s: lookup_severity(s, flags))
        detail["Deduct Value"] = (
            detail["Area Percentage (%)"] * detail["Severity Factor"] * detail["Weighting Factor"]
        )
        detail["Suggested Defect Treatment"] = detail.apply(
            lambda r: lookup_treatment(r["Defect Type"], r["Severity"]), axis=1
        )
    else:
        for c in ["Weighting Factor", "Severity Factor", "Deduct Value", "Suggested Defect Treatment"]:
            detail[c] = np.nan

    all_sections = pd.Index(sorted(df["Section"].unique(), key=lambda x: (len(x), x)))

    sum_deduct = detail.groupby("Section")["Deduct Value"].sum() if not detail.empty else pd.Series(dtype=float)
    defect_count = detail.groupby("Section").size() if not detail.empty else pd.Series(dtype=int)
    avg_iri = df[df["IRI"].notna()].groupby("Section")["IRI"].mean()

    rows = []
    for sec in all_sections:
        has_pci = sec in sum_deduct.index
        has_iri = sec in avg_iri.index

        pci_val = max(0.0, 100.0 - min(100.0, sum_deduct.get(sec, 0.0))) if has_pci else np.nan
        iri_val = float(avg_iri.get(sec)) if has_iri else np.nan

        r_pci = pci_rank(pci_val) if has_pci else None
        r_iri = iri_rank(iri_val) if has_iri else None

        if has_pci and has_iri:
            combined_rank = max(r_pci, r_iri)
            combined_label = PCI_RANK_LABEL[combined_rank]
            combined_reco = PCI_RANK_RECO[combined_rank]
            basis = "Hybrid (PCI & IRI — worse of the two governs)"
        elif has_pci:
            combined_rank = r_pci
            combined_label = PCI_RANK_LABEL[combined_rank]
            combined_reco = PCI_RANK_RECO[combined_rank]
            basis = "PCI only (no IRI data)"
        elif has_iri:
            combined_rank = r_iri
            combined_label = IRI_RANK_LABEL[combined_rank]
            combined_reco = IRI_RANK_RECO[combined_rank]
            basis = "IRI only (no defect data)"
        else:
            combined_rank, combined_label, combined_reco, basis = None, "No Data", "—", "No data"

        rows.append({
            "Section": sec,
            "No. of Defects Recorded": int(defect_count.get(sec, 0)),
            "Sum Deduct Value": round(sum_deduct.get(sec, np.nan), 2) if has_pci else np.nan,
            "PCI": round(pci_val, 1) if has_pci else np.nan,
            "PCI Condition": PCI_RANK_LABEL[r_pci] if has_pci else "—",
            "Avg IRI (m/km)": round(iri_val, 2) if has_iri else np.nan,
            "IRI Condition": IRI_RANK_LABEL[r_iri] if has_iri else "—",
            "Combined Condition Rating": combined_label,
            "_rank": combined_rank,
            "Maintenance Recommendation": combined_reco,
            "Basis": basis,
        })

    summary = pd.DataFrame(rows)
    return detail, summary, flags


def df_to_excel_bytes(sheets: dict) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for name, d in sheets.items():
            d.to_excel(writer, sheet_name=name[:31], index=False)
    return buf.getvalue()


# -----------------------------------------------------------------------------
# LIGHT CUSTOM STYLING
# -----------------------------------------------------------------------------
st.markdown(
    """
    <style>
    .main > div {padding-top: 1.2rem;}
    .kpi-box {
        background: #ffffff; border: 1px solid #e6e6e6; border-radius: 10px;
        padding: 14px 18px; box-shadow: 0 1px 3px rgba(0,0,0,0.06);
    }
    .small-note {color:#666; font-size:0.85rem;}
    .badge {
        display:inline-block; padding:3px 10px; border-radius:14px;
        color:white; font-size:0.78rem; font-weight:600;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def condition_badge(label: str) -> str:
    color = CONDITION_COLORS.get(label, "#777")
    return f'<span class="badge" style="background:{color}">{label}</span>'


# -----------------------------------------------------------------------------
# HEADER
# -----------------------------------------------------------------------------
left, right = st.columns([0.8, 0.2])
with left:
    st.title("🛣️ Digital Pavement Condition Evaluation and Maintenance Decision Tool")
    st.caption(
        "TCG633 Bridge & Road Maintenance · Individual Project · "
        "Fakulti Kejuruteraan Awam, UiTM Cawangan Sarawak"
    )

st.divider()

# -----------------------------------------------------------------------------
# SIDEBAR — DATA INPUT
# -----------------------------------------------------------------------------
with st.sidebar:
    st.header("📥 Data Input")
    uploaded_file = st.file_uploader(
        "Upload pavement condition data (.csv, .xlsx, .xls)",
        type=["csv", "xlsx", "xls"],
        help="Required columns: Section, Defect Type, Severity, Area Percentage (%), IRI",
    )

    col_a, col_b = st.columns(2)
    with col_a:
        load_sample = st.button("📊 Load Sample Data", use_container_width=True)
    with col_b:
        clear_data = st.button("🗑️ Clear", use_container_width=True)

    if clear_data:
        st.session_state.pop("df_raw", None)
        st.session_state.pop("data_source", None)

    if uploaded_file is not None:
        try:
            st.session_state["df_raw"] = read_uploaded_file(uploaded_file)
            st.session_state["data_source"] = f"Uploaded file: {uploaded_file.name}"
        except Exception as e:
            st.error(
                f"File received, but it could not be parsed: {e}\n\n"
                "(The file chip above only confirms it was transferred to the app — "
                "this error means the server-side step that opens/reads it failed.)"
            )

    if load_sample:
        st.session_state["df_raw"] = generate_sample_data()
        st.session_state["data_source"] = "Built-in sample / demo dataset"

    st.divider()
    if "df_raw" in st.session_state:
        st.success(st.session_state.get("data_source", "Data loaded"))
        st.caption(f"{len(st.session_state['df_raw'])} rows loaded")
    else:
        st.info("Upload a file or load the sample dataset to begin.")

    st.divider()
    st.caption(
        "Required columns:\n"
        "- Section\n- Defect Type\n- Severity (Low/Medium/High)\n"
        "- Area Percentage (%)\n- IRI (m/km)\n\n"
        "A section may appear in several rows (one per defect). "
        "IRI may be repeated per row or given once per section."
    )

# -----------------------------------------------------------------------------
# MAIN TABS
# -----------------------------------------------------------------------------
tab_how, tab_data, tab_dash, tab_results, tab_charts, tab_method = st.tabs(
    [
        "🏠 How to Use",
        "📥 Upload & Preview",
        "📊 Dashboard",
        "📋 Detailed Results",
        "📈 Charts",
        "📐 Methodology & Assumptions",
    ]
)

# =============================================================================
# TAB: HOW TO USE
# =============================================================================
with tab_how:
    st.subheader("How to Use This Tool")
    st.markdown(
        """
**1. Prepare your data.**  Build a spreadsheet/CSV with one row per observed
defect (or per IRI sample) using these columns:

| Section | Defect Type | Severity | Area Percentage (%) | IRI |
|---|---|---|---|---|
| S1 | Potholes | High | 6 | 3.8 |
| S1 | Raveling | Low | 10 | 3.8 |
| S2 | Longitudinal Crack | Low | 4 | 1.9 |

- One **Section** can repeat across several rows — one row per defect found in that section.
- **Severity** must be `Low`, `Medium`, or `High`.
- **Area Percentage (%)** is the percentage of the sample area affected by that defect (e.g. `6` for 6%).
- **IRI** (m/km) can be repeated on every row of a section, or supplied once — the
  tool simply averages whatever IRI values it finds for that section.
- You only need defect columns **or** IRI — the tool will compute PCI-only,
  IRI-only, or a combined (hybrid) rating depending on what is available.

**2. Upload the file** using the **Data Input** panel on the left (CSV or Excel),
or click **Load Sample Data** to try the tool with a built-in demo dataset.

**3. Review the results:**
- **Dashboard** — KPI summary cards and the overall network condition.
- **Detailed Results** — section-level summary table and full defect-level
  computation table, plus the download button.
- **Charts** — PCI by section, IRI by section, defect distribution, and
  condition-rating distribution.

**4. Download** your analysed results as an Excel workbook or CSV from the
**Detailed Results** tab to attach to your technical report.

**5. Methodology & Assumptions** explains exactly how PCI, IRI, the combined
rating, and the maintenance recommendations are calculated — useful material
for your report and video presentation.
        """
    )
    st.info(
        "Tip for your video presentation: walk through these five steps live — "
        "upload data → dashboard → results table → charts → download — to cover "
        "the 'Demonstration' and 'Results' sections of Part B.",
        icon="🎬",
    )

# =============================================================================
# DATA LOADING / VALIDATION (shared across remaining tabs)
# =============================================================================
df_raw = st.session_state.get("df_raw")

with tab_data:
    st.subheader("Upload & Preview")
    if df_raw is None:
        st.warning("No data loaded yet. Use the sidebar to upload a file or load the sample dataset.")
    else:
        missing = [c for c in CANONICAL_COLS if c not in df_raw.columns]
        if "Section" in missing:
            st.error(
                "Your file must contain a 'Section' column. Detected columns: "
                + ", ".join(map(str, df_raw.columns))
            )
        else:
            if missing:
                st.warning(
                    f"Columns not found and treated as empty: {', '.join(missing)}. "
                    "The tool will still compute whatever indicator(s) the available data supports."
                )
            st.success(f"Data validated — {df_raw['Section'].nunique()} unique section(s), {len(df_raw)} row(s).")
            st.dataframe(df_raw, use_container_width=True, height=320)

# =============================================================================
# RUN ANALYSIS
# =============================================================================
detail_df, summary_df, calc_flags = (None, None, [])
if df_raw is not None and "Section" in df_raw.columns:
    detail_df, summary_df, calc_flags = compute_results(df_raw)

# =============================================================================
# TAB: DASHBOARD
# =============================================================================
with tab_dash:
    st.subheader("Summary Dashboard")
    if summary_df is None or summary_df.empty:
        st.info("Load data to see the dashboard.")
    else:
        if calc_flags:
            with st.expander(f"⚠️ {len(calc_flags)} data-quality flag(s) detected — click to view", expanded=False):
                for f in sorted(set(calc_flags)):
                    st.write("• " + f)

        total_sections = len(summary_df)
        avg_pci = summary_df["PCI"].mean()
        avg_iri = summary_df["Avg IRI (m/km)"].mean()
        poor_sections = int((summary_df["Combined Condition Rating"].isin(["Poor", "Poor (Rough)"])).sum())
        valid_ranks = summary_df["_rank"].dropna()
        overall_rank = int(round(valid_ranks.mean())) if not valid_ranks.empty else None
        overall_label = PCI_RANK_LABEL.get(overall_rank, "No Data") if overall_rank else "No Data"

        c1, c2, c3, c4, c5 = st.columns(5)
        with c1:
            st.metric("Total Road Sections", total_sections)
        with c2:
            st.metric("Average PCI", f"{avg_pci:.1f}" if pd.notna(avg_pci) else "N/A")
        with c3:
            st.metric("Average IRI (m/km)", f"{avg_iri:.2f}" if pd.notna(avg_iri) else "N/A")
        with c4:
            st.metric("Poor Sections", poor_sections)
        with c5:
            st.markdown("**Overall Network Condition**")
            st.markdown(condition_badge(overall_label), unsafe_allow_html=True)

        st.markdown("")
        colL, colR = st.columns([0.55, 0.45])
        with colL:
            st.markdown("##### Section Condition Overview")
            show_cols = ["Section", "PCI", "PCI Condition", "Avg IRI (m/km)", "IRI Condition",
                         "Combined Condition Rating", "Maintenance Recommendation"]
            st.dataframe(summary_df[show_cols], use_container_width=True, height=360)
        with colR:
            st.markdown("##### Condition Rating Distribution")
            dist = summary_df["Combined Condition Rating"].value_counts().reset_index()
            dist.columns = ["Condition", "Count"]
            if PLOTLY_OK:
                fig = px.pie(
                    dist, names="Condition", values="Count", hole=0.45,
                    color="Condition", color_discrete_map=CONDITION_COLORS,
                )
                fig.update_layout(margin=dict(t=10, b=10, l=10, r=10), height=340)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.bar_chart(dist.set_index("Condition")["Count"])

# =============================================================================
# TAB: DETAILED RESULTS
# =============================================================================
with tab_results:
    st.subheader("Detailed Results")
    if summary_df is None or summary_df.empty:
        st.info("Load data to see detailed results.")
    else:
        st.markdown("##### 1. Section-Level Summary (PCI, IRI, Combined Rating, Recommendation)")
        display_summary = summary_df.drop(columns=["_rank"])
        st.dataframe(display_summary, use_container_width=True, height=320)

        st.markdown("##### 2. Defect-Level Detail & Computation")
        if detail_df is None or detail_df.empty:
            st.info("No defect-level rows found in the uploaded data (Defect Type / Severity / Area % missing).")
        else:
            detail_show = detail_df[[
                "Section", "Defect Type", "Severity", "Area Percentage (%)",
                "Weighting Factor", "Severity Factor", "Deduct Value", "Suggested Defect Treatment"
            ]].reset_index(drop=True)
            st.dataframe(detail_show, use_container_width=True, height=320)
            st.caption(
                "Deduct Value = Area (%) × Severity Factor × Weighting Factor. "
                "'Suggested Defect Treatment' is supplementary general guidance per "
                "defect (see Methodology tab) — the official section rating/recommendation "
                "is the 'Combined Condition Rating' / 'Maintenance Recommendation' columns above."
            )

        st.markdown("##### 3. Download Analysed Results")
        c1, c2 = st.columns(2)
        ts = datetime.now().strftime("%Y%m%d_%H%M")
        with c1:
            st.download_button(
                "⬇️ Download Section Summary (CSV)",
                data=display_summary.to_csv(index=False).encode("utf-8"),
                file_name=f"pavement_section_summary_{ts}.csv",
                mime="text/csv",
                use_container_width=True,
            )
        with c2:
            if not OPENPYXL_OK:
                st.button(
                    "⬇️ Download Full Results (Excel) — unavailable",
                    disabled=True, use_container_width=True,
                    help="The 'openpyxl' package isn't available in this server environment right now.",
                )
                st.caption(
                    "⚠️ Excel download is temporarily unavailable on this server "
                    "(missing 'openpyxl' package) — use the CSV download on the left instead. "
                    "See the note at the bottom of this page for how to fix this on Streamlit Cloud."
                )
            else:
                try:
                    excel_bytes = df_to_excel_bytes({
                        "Section_Summary": display_summary,
                        "Defect_Detail": detail_df if detail_df is not None else pd.DataFrame(),
                        "Raw_Input": df_raw,
                    })
                    st.download_button(
                        "⬇️ Download Full Results (Excel, 3 sheets)",
                        data=excel_bytes,
                        file_name=f"pavement_analysis_results_{ts}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                except Exception as e:
                    st.caption(f"⚠️ Excel export failed ({e}). Use the CSV download on the left instead.")

        if not OPENPYXL_OK:
            with st.expander("ℹ️ Why is Excel upload/download unavailable? (click to view)"):
                st.markdown(
                    "This server's Python environment is missing the **openpyxl** package, "
                    "which both Excel *upload* and Excel *download* depend on. This is almost "
                    "always a deployment/server-side issue, not a problem with your data.\n\n"
                    "**If you're the app owner (Streamlit Community Cloud):** this is a known, "
                    "currently active platform issue where Cloud forces a very new Python version "
                    "(3.14) on new deployments, which breaks the install of several packages. "
                    "The fix is to **delete the app and redeploy it**, choosing an older, stable "
                    "Python version (e.g. 3.11 or 3.12) in the **'Advanced settings'** dialog "
                    "*before* clicking Deploy — Python version can only be set at deploy time, "
                    "not changed afterwards. Then reboot/redeploy with the same `requirements.txt`.\n\n"
                    "**In the meantime:** CSV upload and CSV download both work normally and need "
                    "no extra packages — use CSV instead of Excel for now."
                )

# =============================================================================
# TAB: CHARTS
# =============================================================================
with tab_charts:
    st.subheader("Charts")
    if summary_df is None or summary_df.empty:
        st.info("Load data to see charts.")
    else:
        if not PLOTLY_OK:
            st.warning(
                "Plotly isn't installed in this environment, so charts below are shown "
                "using Streamlit's simplified built-in charts (no colour-by-condition). "
                "Add `plotly` to requirements.txt and reboot the app for the full "
                "interactive charts.",
                icon="⚠️",
            )

        # --- PCI by Section ---
        st.markdown("##### PCI by Section")
        pci_chart_df = summary_df.dropna(subset=["PCI"])
        if pci_chart_df.empty:
            st.caption("No PCI data available to chart.")
        elif PLOTLY_OK:
            fig1 = px.bar(
                pci_chart_df, x="Section", y="PCI", color="PCI Condition",
                color_discrete_map=CONDITION_COLORS, text="PCI",
                category_orders={"Section": list(summary_df["Section"])},
            )
            fig1.add_hline(y=85, line_dash="dot", line_color="#2E7D32", annotation_text="Very Good ≥85")
            fig1.add_hline(y=70, line_dash="dot", line_color="#1976D2", annotation_text="Good ≥70")
            fig1.add_hline(y=55, line_dash="dot", line_color="#F9A825", annotation_text="Fair ≥55")
            fig1.update_traces(texttemplate="%{text:.1f}", textposition="outside")
            fig1.update_layout(yaxis_range=[0, 105], height=380, margin=dict(t=30, b=10))
            st.plotly_chart(fig1, use_container_width=True)
        else:
            st.bar_chart(pci_chart_df.set_index("Section")["PCI"])

        # --- IRI by Section ---
        st.markdown("##### IRI by Section")
        iri_chart_df = summary_df.dropna(subset=["Avg IRI (m/km)"])
        if iri_chart_df.empty:
            st.caption("No IRI data available to chart.")
        elif PLOTLY_OK:
            fig2 = px.bar(
                iri_chart_df, x="Section", y="Avg IRI (m/km)", color="IRI Condition",
                color_discrete_map=CONDITION_COLORS, text="Avg IRI (m/km)",
                category_orders={"Section": list(summary_df["Section"])},
            )
            fig2.add_hline(y=2, line_dash="dot", line_color="#2E7D32", annotation_text="Very Good <2")
            fig2.add_hline(y=3, line_dash="dot", line_color="#1976D2", annotation_text="Good <3")
            fig2.add_hline(y=4, line_dash="dot", line_color="#F9A825", annotation_text="Fair <4")
            fig2.update_traces(texttemplate="%{text:.2f}", textposition="outside")
            fig2.update_layout(height=380, margin=dict(t=30, b=10))
            st.plotly_chart(fig2, use_container_width=True)
        else:
            st.bar_chart(iri_chart_df.set_index("Section")["Avg IRI (m/km)"])

        colA, colB = st.columns(2)
        with colA:
            st.markdown("##### Defect Type Distribution")
            if detail_df is None or detail_df.empty:
                st.caption("No defect-level data available to chart.")
            else:
                defect_counts = (
                    detail_df.groupby(["Defect Type", "Severity"]).size().reset_index(name="Count")
                )
                if PLOTLY_OK:
                    fig3 = px.bar(
                        defect_counts, x="Defect Type", y="Count", color="Severity",
                        color_discrete_map={"Low": "#2E7D32", "Medium": "#F9A825", "High": "#C62828"},
                        barmode="stack",
                    )
                    fig3.update_layout(height=380, margin=dict(t=10, b=10), xaxis_tickangle=-30)
                    st.plotly_chart(fig3, use_container_width=True)
                else:
                    pivot = defect_counts.pivot_table(
                        index="Defect Type", columns="Severity", values="Count", fill_value=0
                    )
                    st.bar_chart(pivot)

        with colB:
            st.markdown("##### Condition Rating Distribution")
            dist2 = summary_df["Combined Condition Rating"].value_counts().reset_index()
            dist2.columns = ["Condition", "Count"]
            if PLOTLY_OK:
                fig4 = px.bar(
                    dist2, x="Condition", y="Count", color="Condition",
                    color_discrete_map=CONDITION_COLORS, text="Count",
                )
                fig4.update_layout(height=380, margin=dict(t=10, b=10), showlegend=False)
                st.plotly_chart(fig4, use_container_width=True)
            else:
                st.bar_chart(dist2.set_index("Condition")["Count"])

        st.markdown("##### Defects by Section (stacked)")
        if detail_df is None or detail_df.empty:
            st.caption("No defect-level data available to chart.")
        elif PLOTLY_OK:
            by_section = detail_df.groupby(["Section", "Defect Type"]).size().reset_index(name="Count")
            fig5 = px.bar(
                by_section, x="Section", y="Count", color="Defect Type", barmode="stack",
                category_orders={"Section": list(summary_df["Section"])},
            )
            fig5.update_layout(height=380, margin=dict(t=10, b=10))
            st.plotly_chart(fig5, use_container_width=True)
        else:
            by_section = detail_df.groupby(["Section", "Defect Type"]).size().reset_index(name="Count")
            pivot2 = by_section.pivot_table(index="Section", columns="Defect Type", values="Count", fill_value=0)
            st.bar_chart(pivot2)

# =============================================================================
# TAB: METHODOLOGY & ASSUMPTIONS
# =============================================================================
with tab_method:
    st.subheader("Methodology & Assumptions")
    st.markdown(
        """
This tool follows the computation logic encoded in the lecturer-provided
workbooks `TCG633_PCI_IRI_Model.xlsx` and `TCG633_PCI_IRI_Pro_v2.xlsx`
(see their `Lookup`, `PCI_Compute`, `IRI_Compute`, and `Settings_Summary`
sheets). It is a **teaching simplification** of pavement evaluation practice,
not the full ASTM D6433 deduct-curve procedure — this is stated explicitly
because the full ASTM method uses graphical, multi-defect corrected deduct
curves that are beyond the scope of the provided spreadsheet template.

### 1. PCI (Pavement Condition Index) — simplified deduct-value method

For every recorded defect:

```
Deduct Value = Area Affected (%) × Severity Factor × Weighting Factor
```

| Defect Type | Weighting Factor | | Severity | Factor |
|---|---|---|---|---|
| Longitudinal Crack | 1.0 | | Low | 0.6 |
| Alligator (Fatigue) Crack | 1.6 | | Medium | 1.0 |
| Potholes | 2.2 | | High | 1.4 |
| Raveling | 1.2 | | | |
| Depression/Sag | 1.4 | | | |
| Patching (Failed) | 1.8 | | | |
| Bleeding/Flushing | 1.0 | | | |
| Rut/Rutting | 1.6 | | | |

For each **Section**, all Deduct Values are summed, then:

```
PCI = MAX( 0, 100 − MIN(100, ΣDeduct Value) )
```

**PCI condition classes:**

| PCI Range | Condition | Recommendation |
|---|---|---|
| 85–100 | Very Good | Routine maintenance |
| 70–84 | Good / Satisfactory | Preventive maintenance (crack sealing, local patching) |
| 55–69 | Fair | Surface treatment / overlay (localized) |
| 0–54 | Poor | Major rehabilitation / reconstruction assessment |

### 2. IRI (International Roughness Index)

For each Section, IRI is the simple average of all IRI (m/km) readings
recorded for that section:

```
Average IRI = mean(IRI readings for that section)
```

**IRI condition classes:**

| IRI (m/km) | Condition | Recommendation |
|---|---|---|
| < 2 | Very Good (Smooth) | Routine maintenance |
| 2 – < 3 | Good | Preventive maintenance (localized patching/leveling) |
| 3 – < 4 | Fair | Surface treatment / thin overlay |
| ≥ 4 | Poor (Rough) | Structural overlay / rehabilitation |

### 3. Combined (Hybrid) Condition Rating

Following the "Hybrid mode" logic in `TCG633_PCI_IRI_Pro_v2.xlsx`
(`Settings_Summary` sheet): when **both** PCI and IRI are available for a
section, the **worse (more severe)** of the two condition classes governs
the combined rating and its recommendation. If only one indicator is
available for a section, that indicator alone determines the rating.

### 4. Assumptions made explicit by this tool

- **Unrecognised defect type** → a neutral weighting factor of **1.0** is
  substituted, and the section is flagged in the Dashboard tab. The lecturer's
  Lookup table only defines 8 defect types; anything else is an assumption.
- **Unrecognised severity** → treated as **Medium (factor 1.0)**, flagged.
- **Section with defect data but no IRI** → rated on **PCI only**.
- **Section with IRI but no defect rows** → rated on **IRI only**.
- **Per-defect "Suggested Defect Treatment" column** (Detailed Results tab) is
  **supplementary, general pavement-maintenance practice guidance** added by
  this tool — it is *not* part of the lecturer's spreadsheet, which only
  issues a recommendation at the section level. This is clearly separated in
  the UI from the official PCI/IRI-based "Maintenance Recommendation".
- Condition-class thresholds and recommendation wording are reproduced
  exactly as given in the Lookup sheet of the supplied workbooks; they are
  **editable in the spreadsheet** (and in the `app.py` constants if a
  different agency standard, e.g. a specific JKR manual edition, is required).
- This tool does not implement GPS/GIS mapping, image-based defect detection,
  or automated PDF report generation — these remain available as **bonus**
  extensions per the project brief, not required functionality.
        """
    )
    st.warning(
        "This is a teaching/learning tool built for an academic assignment. "
        "It should not be used for real-world pavement asset management "
        "decisions without validation against the full ASTM D6433 / JKR "
        "manual procedures by a qualified engineer.",
        icon="⚠️",
    )

# -----------------------------------------------------------------------------
# FOOTER
# -----------------------------------------------------------------------------
st.divider()
st.caption(
    "TCG633 Bridge & Road Maintenance · Digital Pavement Condition Evaluation and "
    "Maintenance Decision Tool · Built with Streamlit · "
    "Computation logic sourced from lecturer-provided Excel model."
)

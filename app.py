from datetime import datetime
from typing import Optional
from urllib.parse import quote

import pandas as pd
import plotly.express as px
import streamlit as st

# =========================================================
# Page config
# =========================================================
st.set_page_config(
    page_title="ABSOFOAM Dashboard",
    page_icon="📊",
    layout="wide"
)

# =========================================================
# Defaults
# =========================================================
DEFAULT_SHEET_NAME = "Data"
DEFAULT_GOOGLE_SHEET_ID = ""

if "gsheets" in st.secrets and "sheet_id" in st.secrets["gsheets"]:
    DEFAULT_GOOGLE_SHEET_ID = st.secrets["gsheets"]["sheet_id"]

# =========================================================
# Formatting helpers
# =========================================================
def format_number(value: Optional[float], decimals: int = 2) -> str:
    if pd.isna(value):
        return "-"
    return f"{value:.{decimals}f}"


def format_percent(value: Optional[float], decimals: int = 1) -> str:
    if pd.isna(value):
        return "-"
    return f"{value:.{decimals}%}"


# =========================================================
# Parsing helpers
# =========================================================
def parse_mixed_numeric(series: pd.Series) -> pd.Series:
    """
    Handles plain numbers, numeric strings, and percentage strings like '4.4%'.
    Returns decimal form for percentages, e.g. '4.4%' -> 0.044.
    """
    s = series.copy()

    # Already numeric
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce")

    s = s.astype(str).str.strip()

    # Track which entries are percentages
    is_percent = s.str.contains("%", na=False)

    # Remove commas and percent signs
    s = s.str.replace(",", "", regex=False).str.replace("%", "", regex=False)

    out = pd.to_numeric(s, errors="coerce")

    # Convert percent-formatted values to decimal
    out.loc[is_percent] = out.loc[is_percent] / 100

    return out


def compute_discrepancy(df: pd.DataFrame) -> pd.DataFrame:
    """
    Computes discrepancy as absolute percentage difference vs COA:
    abs(Inspection - COA) / COA
    """
    if (
        "Adhesiveness on Inspection Report" in df.columns
        and "Adhesiveness on COA" in df.columns
    ):
        inspection = pd.to_numeric(df["Adhesiveness on Inspection Report"], errors="coerce")
        coa = pd.to_numeric(df["Adhesiveness on COA"], errors="coerce")

        discrepancy = ((inspection - coa).abs() / coa.replace(0, pd.NA))
        df["Discrepancy"] = discrepancy

    return df


# =========================================================
# Data normalization
# =========================================================
def normalize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    df = df.dropna(axis=1, how="all")
    df.columns = [str(col).strip() for col in df.columns]

    if "Reference Code" in df.columns and "Reference code" not in df.columns:
        df = df.rename(columns={"Reference Code": "Reference code"})

    numeric_cols = [
        "Adhesiveness reading 1",
        "Adhesiveness reading 2",
        "Adhesiveness reading 3",
        "Adhesiveness on Inspection Report",
        "Adhesiveness on COA",
        "Year",
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    text_cols = [
        "Lot Number",
        "LOT#",
        "Product Range",
        "Reference code",
        "Remarks",
    ]
    for col in text_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace({"nan": pd.NA, "None": pd.NA, "": pd.NA})

    # Parse discrepancy robustly if present
    if "Discrepancy" in df.columns:
        df["Discrepancy"] = parse_mixed_numeric(df["Discrepancy"])

    # Recompute discrepancy if missing or effectively empty
    if "Discrepancy" not in df.columns or df["Discrepancy"].isna().all():
        df = compute_discrepancy(df)

    key_cols = [col for col in ["Product Range", "Lot Number"] if col in df.columns]
    if key_cols:
        df = df.dropna(subset=key_cols, how="all")

    if "Year" in df.columns:
        df["Year"] = pd.to_numeric(df["Year"], errors="coerce").astype("Int64")

    return df


# =========================================================
# Google Sheets loading
# =========================================================
@st.cache_data(show_spinner=False)
def load_data_from_gsheet(sheet_id: str, sheet_name: str) -> pd.DataFrame:
    encoded_sheet_name = quote(sheet_name)
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={encoded_sheet_name}"
    df = pd.read_csv(url)
    return normalize_dataframe(df)


# =========================================================
# Validation
# =========================================================
def validate_required_columns(df: pd.DataFrame) -> None:
    required_columns = [
        "Year",
        "Lot Number",
        "Product Range",
        "Adhesiveness on Inspection Report",
        "Adhesiveness on COA",
    ]
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        st.error(f"Colonnes manquantes / Missing required columns: {missing_columns}")
        st.stop()


# =========================================================
# Session-state defaults for filters
# =========================================================
def init_filter_state(df: pd.DataFrame) -> None:
    product_ranges = sorted(df["Product Range"].dropna().astype(str).unique().tolist())
    years = sorted(df["Year"].dropna().astype(int).unique().tolist()) if "Year" in df.columns else []

    if "selected_product_ranges" not in st.session_state:
        st.session_state.selected_product_ranges = product_ranges

    if "selected_years" not in st.session_state:
        st.session_state.selected_years = years

    reference_df = df.copy()
    if st.session_state.selected_product_ranges:
        reference_df = reference_df[
            reference_df["Product Range"].astype(str).isin(st.session_state.selected_product_ranges)
        ]

    reference_codes = []
    if "Reference code" in reference_df.columns:
        reference_codes = sorted(
            reference_df["Reference code"].dropna().astype(str).unique().tolist()
        )

    if "selected_reference_codes" not in st.session_state:
        st.session_state.selected_reference_codes = reference_codes

    # Keep references valid when product range changes
    st.session_state.selected_reference_codes = [
        x for x in st.session_state.selected_reference_codes if x in reference_codes
    ] or reference_codes


# =========================================================
# Header
# =========================================================
st.title("📊 ABSOFOAM – Adhesiveness Dashboard")
st.caption("Suivi interactif de l’adhésivité | Interactive adhesiveness monitoring")
st.caption(f"Dernière actualisation / Last refresh: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

# =========================================================
# Sidebar - Data source
# =========================================================
st.sidebar.header("Source des données / Data source")

sheet_name = st.sidebar.text_input(
    "Nom de feuille / Sheet name",
    value=DEFAULT_SHEET_NAME
)

secret_sheet_id = ""
if "gsheets" in st.secrets and "sheet_id" in st.secrets["gsheets"]:
    secret_sheet_id = st.secrets["gsheets"]["sheet_id"]

if secret_sheet_id:
    sheet_id = secret_sheet_id
    st.sidebar.success("Google Sheet connecté / Google Sheet connected")
else:
    sheet_id = st.sidebar.text_input(
        "Google Sheet ID",
        value=DEFAULT_GOOGLE_SHEET_ID,
        help="Paste the Google Sheet ID from the spreadsheet URL"
    )

df = None
load_error = None

try:
    if sheet_id:
        df = load_data_from_gsheet(sheet_id, sheet_name)
    else:
        st.info(
            "Veuillez coller le Google Sheet ID dans la barre latérale.\n\n"
            "Please paste the Google Sheet ID in the sidebar."
        )
except Exception as exc:
    load_error = str(exc)

if load_error:
    st.error(f"Erreur de chargement / Loading error:\n\n{load_error}")
    st.warning(
        "Checklist:\n"
        "1. Confirm the Google Sheet is shared as 'Anyone with the link - Viewer'\n"
        "2. Confirm the worksheet name is correct, for example 'Data'\n"
        "3. Confirm you pasted the Sheet ID only, not the full URL"
    )

if df is None:
    st.stop()

validate_required_columns(df)
init_filter_state(df)

# =========================================================
# Sidebar - Filters
# =========================================================
st.sidebar.header("Filtres / Filters")

product_ranges = sorted(df["Product Range"].dropna().astype(str).unique().tolist())
years = sorted(df["Year"].dropna().astype(int).unique().tolist()) if "Year" in df.columns else []

selected_product_ranges = st.sidebar.multiselect(
    "Gamme produit / Product Range",
    options=product_ranges,
    default=st.session_state.selected_product_ranges,
    key="selected_product_ranges"
)
if selected_product_ranges:
    st.sidebar.caption("Sélection actuelle / Current selection")
    for item in selected_product_ranges:
        st.sidebar.write(f"- {item}")

reference_df = df.copy()
if selected_product_ranges:
    reference_df = reference_df[
        reference_df["Product Range"].astype(str).isin(selected_product_ranges)
    ]

reference_codes = []
if "Reference code" in reference_df.columns:
    reference_codes = sorted(
        reference_df["Reference code"].dropna().astype(str).unique().tolist()
    )

# Keep references valid
current_ref_defaults = [x for x in st.session_state.selected_reference_codes if x in reference_codes] or reference_codes
st.session_state.selected_reference_codes = current_ref_defaults

selected_reference_codes = st.sidebar.multiselect(
    "Référence produit / Product Reference",
    options=reference_codes,
    default=st.session_state.selected_reference_codes,
    key="selected_reference_codes"
)

selected_years = st.sidebar.multiselect(
    "Année / Year",
    options=years,
    default=st.session_state.selected_years,
    key="selected_years"
)

metric_choice = st.sidebar.selectbox(
    "Métrique du graphique / Chart metric",
    options=["Inspection only", "COA only", "Both"],
    index=2
)

show_raw_data = st.sidebar.checkbox(
    "Afficher les données brutes / Show raw data",
    value=False
)

if st.sidebar.button("Réinitialiser les filtres / Reset filters"):
    st.session_state.selected_product_ranges = product_ranges
    st.session_state.selected_reference_codes = reference_codes
    st.session_state.selected_years = years
    st.rerun()

# =========================================================
# Filter data
# =========================================================
filtered_df = df.copy()

if selected_product_ranges:
    filtered_df = filtered_df[
        filtered_df["Product Range"].astype(str).isin(selected_product_ranges)
    ]

if selected_reference_codes and "Reference code" in filtered_df.columns:
    filtered_df = filtered_df[
        filtered_df["Reference code"].astype(str).isin(selected_reference_codes)
    ]

if selected_years and "Year" in filtered_df.columns:
    filtered_df = filtered_df[
        filtered_df["Year"].isin(selected_years)
    ]

if filtered_df.empty:
    st.warning("Aucune donnée pour les filtres sélectionnés / No data for the selected filters.")
    st.stop()

# Ensure discrepancy still exists after filtering
if "Discrepancy" not in filtered_df.columns or filtered_df["Discrepancy"].isna().all():
    filtered_df = compute_discrepancy(filtered_df)

# =========================================================
# Derived summaries
# =========================================================
total_rows = len(filtered_df)
total_lots = filtered_df["Lot Number"].nunique()
avg_inspection = filtered_df["Adhesiveness on Inspection Report"].mean()
avg_coa = filtered_df["Adhesiveness on COA"].mean()
avg_discrepancy = filtered_df["Discrepancy"].mean() if "Discrepancy" in filtered_df.columns else pd.NA
high_discrepancy_count = (
    (filtered_df["Discrepancy"] > 0.10).sum()
    if "Discrepancy" in filtered_df.columns else 0
)

chart_df = (
    filtered_df.groupby(["Year", "Lot Number"], as_index=False)[
        ["Adhesiveness on Inspection Report", "Adhesiveness on COA"]
    ]
    .mean()
    .sort_values(by=["Year", "Lot Number"])
)

chart_df["Lot Label"] = (
    chart_df["Year"].astype(str) + " | " + chart_df["Lot Number"].astype(str)
)

yearly_avg = (
    filtered_df.groupby("Year", as_index=False)[
        ["Adhesiveness on Inspection Report", "Adhesiveness on COA"]
    ]
    .mean()
)

yearly_long = yearly_avg.melt(
    id_vars="Year",
    value_vars=["Adhesiveness on Inspection Report", "Adhesiveness on COA"],
    var_name="Metric",
    value_name="Average Adhesiveness"
)

yearly_long["Metric"] = yearly_long["Metric"].replace({
    "Adhesiveness on Inspection Report": "Inspection Report",
    "Adhesiveness on COA": "COA"
})

# More robust discrepancy aggregation
discrepancy_by_year = pd.DataFrame()
if "Discrepancy" in filtered_df.columns:
    discrepancy_by_year = (
        filtered_df.dropna(subset=["Discrepancy"])
        .groupby("Year", as_index=False)["Discrepancy"]
        .mean()
        .sort_values("Year")
    )

# =========================================================
# Top summary row
# =========================================================
top_left, top_right = st.columns([4, 1])

with top_left:
    selected_products_text = ", ".join(selected_product_ranges) if selected_product_ranges else "All"
    st.markdown(f"**Gamme sélectionnée / Selected range:** {selected_products_text}")

with top_right:
    if st.button("🔄 Actualiser / Refresh"):
        st.cache_data.clear()
        st.rerun()

# =========================================================
# KPI row
# =========================================================
k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Inspection moyenne / Avg Inspection", format_number(avg_inspection, 2))
k2.metric("COA moyenne / Avg COA", format_number(avg_coa, 2))
k3.metric("Lots uniques / Unique Lots", f"{total_lots}")
k4.metric("Écart moyen / Avg Discrepancy", format_percent(avg_discrepancy, 1))
k5.metric("Lots > 10% écart / Lots > 10% discrepancy", int(high_discrepancy_count))

# =========================================================
# Tabs
# =========================================================
tab1, tab2, tab3 = st.tabs([
    "📈 Vue d'ensemble / Overview",
    "📊 Analyse / Analysis",
    "🧾 Données / Data"
])

# =========================================================
# TAB 1 - Overview
# =========================================================
with tab1:
    st.subheader("Tendance par lot / Trend by lot")

    if metric_choice == "Both":
        trend_long = chart_df.melt(
            id_vars=["Year", "Lot Number", "Lot Label"],
            value_vars=["Adhesiveness on Inspection Report", "Adhesiveness on COA"],
            var_name="Metric",
            value_name="Adhesiveness"
        )

        trend_long["Metric"] = trend_long["Metric"].replace({
            "Adhesiveness on Inspection Report": "Inspection Report",
            "Adhesiveness on COA": "COA"
        })

        fig_trend = px.line(
            trend_long,
            x="Lot Label",
            y="Adhesiveness",
            color="Metric",
            markers=True,
            hover_data=["Year", "Lot Number"],
            title="Inspection Report vs COA"
        )
    else:
        y_col = (
            "Adhesiveness on Inspection Report"
            if metric_choice == "Inspection only"
            else "Adhesiveness on COA"
        )

        chart_title = (
            "Inspection Report trend"
            if metric_choice == "Inspection only"
            else "COA trend"
        )

        fig_trend = px.line(
            chart_df,
            x="Lot Label",
            y=y_col,
            markers=True,
            hover_data=["Year", "Lot Number"],
            title=chart_title
        )

    fig_trend.update_layout(
        xaxis_title="Lot",
        yaxis_title="Adhesiveness",
        height=500
    )
    st.plotly_chart(fig_trend, use_container_width=True)

    c1, c2 = st.columns(2)

    with c1:
        st.subheader("Moyenne annuelle / Yearly average")
        fig_year = px.bar(
            yearly_long,
            x="Year",
            y="Average Adhesiveness",
            color="Metric",
            barmode="group",
            title="Average adhesiveness by year"
        )
        fig_year.update_layout(height=420)
        st.plotly_chart(fig_year, use_container_width=True)

    with c2:
        st.subheader("Écart annuel / Yearly discrepancy")
        if not discrepancy_by_year.empty:
            fig_disc = px.bar(
                discrepancy_by_year,
                x="Year",
                y="Discrepancy",
                title="Average discrepancy by year"
            )
            fig_disc.update_layout(
                height=420,
                yaxis_tickformat=".0%"
            )
            st.plotly_chart(fig_disc, use_container_width=True)
        else:
            st.info(
                "Aucune donnée exploitable pour l'écart / No usable discrepancy data found."
            )

# =========================================================
# TAB 2 - Analysis
# =========================================================
with tab2:
    st.subheader("Résumé analytique / Analytical summary")

    a1, a2 = st.columns(2)

    with a1:
        st.markdown("**Périmètre / Scope**")
        st.write(f"- Lignes filtrées / Filtered rows: **{total_rows}**")
        st.write(f"- Lots uniques / Unique lots: **{total_lots}**")
        st.write(f"- Années sélectionnées / Selected years: **{', '.join(map(str, selected_years)) if selected_years else 'All'}**")

    with a2:
        st.markdown("**Indicateurs / Indicators**")
        st.write(f"- Inspection moyenne / Avg Inspection: **{format_number(avg_inspection, 2)}**")
        st.write(f"- COA moyenne / Avg COA: **{format_number(avg_coa, 2)}**")
        st.write(f"- Écart moyen / Avg Discrepancy: **{format_percent(avg_discrepancy, 1)}**")

    if "Reference code" in filtered_df.columns:
        st.subheader("Répartition par référence / Breakdown by reference")

        ref_summary = (
            filtered_df.groupby("Reference code", as_index=False)
            .agg(
                Lots=("Lot Number", "nunique"),
                Avg_Inspection=("Adhesiveness on Inspection Report", "mean"),
                Avg_COA=("Adhesiveness on COA", "mean")
            )
            .sort_values("Lots", ascending=False)
        )

        fig_ref = px.bar(
            ref_summary.sort_values("Lots", ascending=False),
            x="Reference code",
            y="Lots",
            title="Number of lots by reference"
        )
        fig_ref.update_layout(height=420)
        st.plotly_chart(fig_ref, use_container_width=True)

        st.dataframe(ref_summary, use_container_width=True)

# =========================================================
# TAB 3 - Data
# =========================================================
with tab3:
    st.subheader("Données filtrées / Filtered data")

    display_df = filtered_df.rename(columns={
        "Product Range": "Gamme produit / Product Range",
        "Reference code": "Référence produit / Product Reference",
        "Lot Number": "Numéro de lot / Lot Number",
        "Adhesiveness on Inspection Report": "Inspection Report",
        "Adhesiveness on COA": "COA",
        "Discrepancy": "Écart / Discrepancy",
        "Year": "Année / Year",
    }).copy()

    if "Écart / Discrepancy" in display_df.columns:
        display_df["Écart / Discrepancy"] = display_df["Écart / Discrepancy"].apply(
            lambda x: format_percent(x, 1) if pd.notna(x) else "-"
        )

    if show_raw_data:
        st.dataframe(display_df, use_container_width=True)
    else:
        st.info(
            "Cochez l’option dans la barre latérale pour afficher les données brutes.\n\n"
            "Use the sidebar checkbox to display raw data."
        )

    download_df = filtered_df.copy()

    # Ensure LOT# is included if available
    columns_order = [
        "Year",
        "Product Range",
        "Reference code",
        "Lot Number",
        "LOT#",
        "Adhesiveness on Inspection Report",
        "Adhesiveness on COA",
        "Discrepancy"
    ]


    # Keep only existing columns (avoid errors)
    columns_order = [col for col in columns_order if col in download_df.columns]

    download_df = download_df[columns_order]
    download_df = download_df.rename(columns={
    "Product Range": "Product Range",
    "Reference code": "Product Reference",
    "Lot Number": "Shipment LOT (YYMM)",
    "LOT#": "Product LOT#",
    "Adhesiveness on Inspection Report": "Inspection",
    "Adhesiveness on COA": "COA",
    "Discrepancy": "Discrepancy"
    })

    csv_data = download_df.to_csv(index=False).encode("utf-8")

    st.download_button(
        "⬇️ Télécharger les données filtrées / Download filtered CSV",
        data=csv_data,
        file_name="absofoam_filtered_data.csv",
        mime="text/csv"
    )

# =========================================================
# Footer
# =========================================================
st.markdown("---")
st.caption("Connected to Google Sheets. Update the spreadsheet, then click Refresh in the dashboard.")

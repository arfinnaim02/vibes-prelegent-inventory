from __future__ import annotations

import io
import re
from pathlib import Path
from typing import Iterable

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


APP_DIR = Path(__file__).resolve().parent
DEFAULT_FILE = (
    APP_DIR
    / "data"
    / "Vibes_Current_Stock_Prelegant_Mall.xlsx"
)

st.set_page_config(
    page_title="VIBES Prelegent Inventory Portal",
    page_icon="◆",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ============================================================
# PREMIUM DESIGN
# ============================================================

st.markdown(
    """
<style>
:root {
    --ink:#102A43;
    --muted:#627D98;
    --line:#D9E2EC;
    --bg:#F4F7FA;
    --teal:#2A9D8F;
    --blue:#1F6E8C;
    --navy:#102A43;
}

.stApp {
    background:var(--bg);
}

.block-container {
    padding-top:1.2rem;
    padding-bottom:3rem;
    max-width:1680px;
}

[data-testid="stSidebar"] {
    background:#0B1F33;
}

[data-testid="stSidebar"] * {
    color:#F8FAFC;
}

[data-testid="stSidebar"] label {
    color:#F8FAFC!important;
}

.portal-title {
    font-size:2rem;
    font-weight:850;
    letter-spacing:-.04em;
    color:var(--ink);
    margin:.1rem 0;
}

.portal-sub {
    color:var(--muted);
    font-size:.95rem;
    margin-bottom:1rem;
}

.eyebrow {
    color:#486581;
    font-size:.72rem;
    text-transform:uppercase;
    letter-spacing:.12em;
    font-weight:800;
}

.section-title {
    font-size:1.22rem;
    font-weight:800;
    color:var(--ink);
    margin:1.15rem 0 .4rem;
}

.section-sub {
    color:var(--muted);
    margin:-.15rem 0 .8rem;
    font-size:.9rem;
}

.kpi-card {
    background:#fff;
    border:1px solid #E4E7EC;
    border-radius:17px;
    padding:17px;
    min-height:116px;
    box-shadow:0 2px 6px rgba(16,42,67,.04);
}

.kpi-label {
    font-size:.76rem;
    color:#627D98;
    text-transform:uppercase;
    letter-spacing:.06em;
    font-weight:750;
}

.kpi-value {
    margin-top:8px;
    font-size:1.72rem;
    font-weight:850;
    color:#102A43;
    letter-spacing:-.035em;
}

.kpi-note {
    margin-top:4px;
    font-size:.76rem;
    color:#829AB1;
}

.kpi-teal {
    border-top:4px solid #2A9D8F;
}

.kpi-blue {
    border-top:4px solid #1F6E8C;
}

.kpi-red {
    border-top:4px solid #E05252;
}

.kpi-amber {
    border-top:4px solid #E9A23B;
}

.kpi-dark {
    border-top:4px solid #102A43;
}

.kpi-purple {
    border-top:4px solid #7B61A8;
}

.notice {
    background:#fff;
    border:1px solid #D9E2EC;
    border-radius:13px;
    padding:12px 14px;
    color:#486581;
    font-size:.86rem;
}

.notice strong {
    color:#102A43;
}

.filter-summary {
    background:#EAF6FA;
    border:1px solid #B8E0EA;
    border-radius:13px;
    padding:11px 14px;
    color:#175267;
    font-size:.86rem;
    margin-bottom:10px;
}

.return-banner {
    background:linear-gradient(135deg,#FFF5F5,#FFF);
    border:1px solid #FED7D7;
    border-left:5px solid #E05252;
    border-radius:14px;
    padding:14px 16px;
    color:#9B2C2C;
    font-size:.88rem;
}

.return-banner strong {
    color:#742A2A;
}

div[data-testid="stDataFrame"] {
    border:1px solid #D9E2EC;
    border-radius:14px;
    overflow:hidden;
    background:white;
}

div[data-testid="stPlotlyChart"] {
    background:white;
    border:1px solid #E4E7EC;
    border-radius:16px;
    padding:5px;
}

.stTabs [data-baseweb="tab-list"] {
    gap:7px;
}

.stTabs [data-baseweb="tab"] {
    border-radius:10px;
    padding:8px 13px;
    background:#fff;
    border:1px solid #D9E2EC;
}

.stTabs [aria-selected="true"] {
    background:#102A43!important;
    color:#fff!important;
}
</style>
""",
    unsafe_allow_html=True,
)


# ============================================================
# CONFIGURATION
# ============================================================

CATEGORY_RULES = [
    (r"cap", "Accessories"),
    (r"punjabi", "Punjabi"),
    (r"hoodie", "Hoodies"),
    (r"sweatpant", "Sweatpants"),
    (r"jacket|funnel", "Jackets & Sets"),
    (r"short pant", "Shorts"),
    (r"formal pant", "Formal Pants"),
    (r"jeans|denim", "Jeans & Denim"),
    (r"polo", "Polo Shirts"),
    (r"shirt", "Shirts"),
    (r"drop shoulder", "T-Shirts"),
    (r"boxer", "Innerwear"),
]

SIZE_ALIASES = {
    "NA": "N/A",
    "N/A": "N/A",
    "FREE": "N/A",
    "FREE SIZE": "N/A",
    "XXL": "2XL",
    "2XL": "2XL",
    "XXXL": "3XL",
    "3XL": "3XL",
    "XXXXL": "4XL",
    "4XL": "4XL",
}

NUMERIC_STOCK_COLS = [
    "Stock",
    "WI Stock",
    "WI Out Stock",
    "Total Stock",
    "Return Stock",
]


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def fmt_int(value):
    return f"{int(round(float(value))):,}"


def fmt_money(value):
    return f"৳{float(value):,.0f}"


def kpi(label, value, note, tone="dark"):
    st.markdown(
        f"""
        <div class="kpi-card kpi-{tone}">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-note">{note}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def clean_label(value, fallback="N/A"):
    if pd.isna(value) or not str(value).strip():
        return fallback

    return re.sub(r"\s+", " ", str(value).strip())


def display_case(value):
    text = clean_label(value)

    return " ".join(
        word
        if word.isupper() and len(word) <= 4
        else word.capitalize()
        for word in text.split()
    )


def normalize_size(value):
    size = clean_label(value).upper()
    return SIZE_ALIASES.get(size, size)


def natural_size_key(value):
    size = normalize_size(value)

    alpha = {
        "N/A": 0,
        "XS": 10,
        "S": 20,
        "M": 30,
        "L": 40,
        "XL": 50,
        "2XL": 60,
        "3XL": 70,
        "4XL": 80,
        "5XL": 90,
    }

    if size in alpha:
        return 0, alpha[size], size

    combination = re.fullmatch(
        r"(XS|S|M|L|XL|2XL|3XL|4XL)-(\d+)",
        size,
    )

    if combination:
        return (
            2,
            int(combination.group(2)),
            alpha.get(combination.group(1), 999),
            size,
        )

    numbers = re.findall(r"\d+(?:\.\d+)?", size)

    if numbers:
        return 1, float(numbers[0]), size

    return 3, 9999, size


def order_sizes(values: Iterable[object]):
    sizes = {
        normalize_size(value)
        for value in values
        if pd.notna(value) and str(value).strip()
    }

    return sorted(sizes, key=natural_size_key)


def infer_category(product):
    for pattern, category in CATEGORY_RULES:
        if re.search(pattern, product.lower()):
            return category

    return "Other Apparel"


def parse_price(series):
    return pd.to_numeric(
        series.astype(str).str.replace(
            r"[^0-9.\-]",
            "",
            regex=True,
        ),
        errors="coerce",
    ).fillna(0)


def canonicalize_case(series):
    cleaned = series.map(clean_label)
    lookup = {}

    for value in cleaned:
        lookup.setdefault(
            value.casefold(),
            display_case(value),
        )

    return cleaned.map(
        lambda value: lookup[value.casefold()]
    )


def find_inventory_sheet(file_obj):
    excel_file = pd.ExcelFile(file_obj)

    required = {
        "Name",
        "Size",
        "Color",
        "SKU",
        "Total Stock",
    }

    for sheet in excel_file.sheet_names:
        candidate = pd.read_excel(
            file_obj,
            sheet_name=sheet,
        )

        columns = {
            str(column).strip()
            for column in candidate.columns
        }

        if required.issubset(columns):
            return candidate

    raise ValueError(
        "No sheet contains Name, Size, Color, SKU "
        "and Total Stock columns."
    )


def prepare_df(raw):
    df = raw.copy()

    df.columns = [
        str(column).strip()
        for column in df.columns
    ]

    df = df.dropna(how="all")

    required = [
        "Name",
        "Size",
        "Color",
        "SKU",
        "Total Stock",
    ]

    missing = [
        column
        for column in required
        if column not in df.columns
    ]

    if missing:
        raise ValueError(
            f"Missing required columns: {', '.join(missing)}"
        )

    df["SKU"] = (
        df["SKU"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    df["Name"] = (
        df["Name"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    df = df[
        df["SKU"].ne("")
        & ~df["Name"].str.casefold().eq("total")
    ].copy()

    df["Product"] = canonicalize_case(df["Name"])
    df["Category"] = df["Product"].map(infer_category)
    df["Size"] = df["Size"].map(normalize_size)
    df["Color"] = canonicalize_case(df["Color"])

    if "Brands" in df.columns:
        df["Brand"] = df["Brands"]
    else:
        df["Brand"] = "Vibes"

    df["Brand"] = (
        df["Brand"]
        .fillna("Vibes")
        .astype(str)
        .str.strip()
    )

    if "Status" in df.columns:
        df["Status"] = df["Status"]
    else:
        df["Status"] = "Active"

    df["Status"] = (
        df["Status"]
        .fillna("Active")
        .astype(str)
        .str.strip()
    )

    if "Price" not in df.columns:
        df["Price"] = 0

    df["Price"] = parse_price(df["Price"])

    for column in NUMERIC_STOCK_COLS:
        if column not in df.columns:
            df[column] = 0

        df[column] = pd.to_numeric(
            df[column],
            errors="coerce",
        ).fillna(0)

    # All current-stock calculations use Total Stock only.
    df["Stock Value"] = (
        df["Total Stock"] * df["Price"]
    )

    # Returns are calculated separately.
    df["Return Value"] = (
        df["Return Stock"] * df["Price"]
    )

    df["Return Status"] = "No Return"

    df.loc[
        df["Return Stock"].gt(0),
        "Return Status",
    ] = "Return Required"

    df["Stock Status"] = "Healthy (3+)"

    df.loc[
        df["Total Stock"].eq(0),
        "Stock Status",
    ] = "Out of Stock"

    df.loc[
        df["Total Stock"].between(1, 2),
        "Stock Status",
    ] = "Low Stock (1-2)"

    size_order = order_sizes(df["Size"])

    size_lookup = {
        size: index
        for index, size in enumerate(size_order)
    }

    df["Size Rank"] = df["Size"].map(size_lookup)

    return df.reset_index(drop=True)


@st.cache_data(show_spinner=False)
def load_inventory_from_path(path_str):
    return prepare_df(
        find_inventory_sheet(Path(path_str))
    )


@st.cache_data(show_spinner=False)
def load_inventory_from_bytes(file_bytes):
    return prepare_df(
        find_inventory_sheet(
            io.BytesIO(file_bytes)
        )
    )


def filter_inventory(
    data,
    search_text="",
    categories=None,
    products=None,
    sizes=None,
    colors=None,
    stock_states=None,
):
    result = data

    filters = [
        ("Category", categories),
        ("Product", products),
        ("Size", sizes),
        ("Color", colors),
        ("Stock Status", stock_states),
    ]

    for column, values in filters:
        if values:
            result = result[
                result[column].isin(values)
            ]

    if search_text.strip():
        search_mask = pd.Series(
            False,
            index=result.index,
        )

        for column in [
            "Product",
            "SKU",
            "Color",
            "Size",
            "Category",
        ]:
            search_mask |= (
                result[column]
                .astype(str)
                .str.contains(
                    search_text.strip(),
                    case=False,
                    regex=False,
                    na=False,
                )
            )

        result = result[search_mask]

    return result.copy()


def size_summary(data):
    if data.empty:
        return pd.DataFrame(
            columns=[
                "Size",
                "Total Stock",
                "Returns",
                "Variants",
                "Out of Stock",
            ]
        )

    summary = (
        data.groupby(
            "Size",
            as_index=False,
        )
        .agg(
            **{
                "Total Stock": (
                    "Total Stock",
                    "sum",
                ),
                "Returns": (
                    "Return Stock",
                    "sum",
                ),
                "Variants": (
                    "SKU",
                    "nunique",
                ),
                "Out of Stock": (
                    "Total Stock",
                    lambda values: int(
                        values.eq(0).sum()
                    ),
                ),
            }
        )
    )

    ordering = {
        value: index
        for index, value in enumerate(
            order_sizes(summary["Size"])
        )
    }

    summary["_order"] = summary["Size"].map(
        ordering
    )

    return (
        summary
        .sort_values("_order")
        .drop(columns="_order")
    )


def color_size_matrix(
    data,
    value_column="Total Stock",
):
    if data.empty:
        return pd.DataFrame()

    matrix = pd.pivot_table(
        data,
        index="Color",
        columns="Size",
        values=value_column,
        aggfunc="sum",
        fill_value=0,
    )

    matrix = matrix.reindex(
        columns=order_sizes(data["Size"]),
        fill_value=0,
    )

    matrix["Total"] = matrix.sum(axis=1)
    matrix.loc["TOTAL"] = matrix.sum(axis=0)

    return matrix.astype(int)


def product_size_matrix(data):
    if data.empty:
        return pd.DataFrame()

    matrix = pd.pivot_table(
        data,
        index=["Category", "Product"],
        columns="Size",
        values="Total Stock",
        aggfunc="sum",
        fill_value=0,
    )

    matrix = matrix.reindex(
        columns=order_sizes(data["Size"]),
        fill_value=0,
    )

    matrix["Grand Total"] = matrix.sum(axis=1)

    return (
        matrix
        .astype(int)
        .sort_values(
            "Grand Total",
            ascending=False,
        )
    )


def matrix_style(matrix):
    return (
        matrix.style
        .format("{:,.0f}")
        .background_gradient(
            cmap="BuGn",
            axis=None,
        )
        .set_properties(
            **{
                "text-align": "center",
                "font-weight": "600",
            }
        )
        .set_table_styles(
            [
                {
                    "selector": "th",
                    "props": [
                        (
                            "background-color",
                            "#1F6E8C",
                        ),
                        ("color", "white"),
                        ("font-weight", "800"),
                        ("text-align", "center"),
                    ],
                }
            ]
        )
    )


def bar_chart(
    data,
    x,
    y,
    title,
    orientation="v",
):
    if data.empty:
        return None

    if orientation == "h":
        figure = px.bar(
            data,
            x=y,
            y=x,
            orientation="h",
            text_auto=True,
        )
    else:
        figure = px.bar(
            data,
            x=x,
            y=y,
            text_auto=True,
        )

    figure.update_traces(
        marker_color="#1F6E8C",
        textposition="outside",
        cliponaxis=False,
    )

    figure.update_layout(
        template="plotly_white",
        title=title,
        height=400,
        margin=dict(
            l=20,
            r=20,
            t=55,
            b=35,
        ),
        showlegend=False,
    )

    return figure


def stock_status_chart(data):
    status = (
        data.groupby(
            "Stock Status",
            as_index=False,
        )["SKU"]
        .nunique()
        .rename(
            columns={"SKU": "Variants"}
        )
    )

    figure = px.pie(
        status,
        names="Stock Status",
        values="Variants",
        hole=.62,
        color="Stock Status",
        color_discrete_map={
            "Out of Stock": "#E05252",
            "Low Stock (1-2)": "#E9A23B",
            "Healthy (3+)": "#2A9D8F",
        },
    )

    figure.update_layout(
        title="SKU Stock Health",
        template="plotly_white",
        height=400,
        margin=dict(
            l=20,
            r=20,
            t=55,
            b=25,
        ),
        legend_title_text="",
    )

    return figure


def return_status_chart(data):
    affected = int(
        data.loc[
            data["Return Stock"].gt(0),
            "SKU",
        ].nunique()
    )

    clear = int(
        data["SKU"].nunique() - affected
    )

    status = pd.DataFrame(
        {
            "Status": [
                "Return Required",
                "No Return",
            ],
            "Variants": [
                affected,
                clear,
            ],
        }
    )

    figure = px.pie(
        status,
        names="Status",
        values="Variants",
        hole=.65,
        color="Status",
        color_discrete_map={
            "Return Required": "#E05252",
            "No Return": "#D9E2EC",
        },
    )

    figure.update_layout(
        title="Return-Affected SKU Share",
        template="plotly_white",
        height=400,
        margin=dict(
            l=20,
            r=20,
            t=55,
            b=25,
        ),
        legend_title_text="",
    )

    return figure


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:
    st.markdown("## VIBES × PRELEGENT")
    st.caption("Current Inventory Portal")
    st.markdown("---")

    upload = st.file_uploader(
        "Upload latest stock workbook",
        type=["xlsx"],
    )

    st.markdown("---")
    st.caption("Inventory definitions")
    st.caption("• Current inventory = Total Stock")
    st.caption("• Low stock = 1–2 total units")
    st.caption("• Returns = Return Stock")
    st.caption("• Size aliases are normalized")
    st.caption(
        "• Product names are grouped case-insensitively"
    )


# ============================================================
# LOAD INVENTORY
# ============================================================

try:
    if upload is not None:
        inventory = load_inventory_from_bytes(
            upload.getvalue()
        )

        data_label = upload.name

    else:
        inventory = load_inventory_from_path(
            str(DEFAULT_FILE)
        )

        data_label = DEFAULT_FILE.name

except Exception as exc:
    st.error(
        f"Could not load inventory: {exc}"
    )
    st.stop()


# ============================================================
# HEADER AND NAVIGATION
# ============================================================

st.markdown(
    '<div class="eyebrow">VIBES / PRELEGENT MALL</div>',
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="portal-title">
        Current Inventory Command Center
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    f"""
    <div class="portal-sub">
        Product, category, colour and size-level stock visibility
        · Data: <b>{data_label}</b>
    </div>
    """,
    unsafe_allow_html=True,
)

tabs = st.tabs(
    [
        "Dashboard",
        "Premium Search & Size Analysis",
        "Category Stock",
        "Return Analysis",
        "SKU Data",
    ]
)


# ============================================================
# DASHBOARD
# ============================================================

with tabs[0]:
    total_skus = inventory["SKU"].nunique()
    total_products = inventory["Product"].nunique()
    total_stock = inventory["Total Stock"].sum()

    out_count = inventory.loc[
        inventory["Total Stock"].eq(0),
        "SKU",
    ].nunique()

    low_count = inventory.loc[
        inventory["Total Stock"].between(1, 2),
        "SKU",
    ].nunique()

    total_returns = inventory["Return Stock"].sum()
    return_value = inventory["Return Value"].sum()

    return_skus = inventory.loc[
        inventory["Return Stock"].gt(0),
        "SKU",
    ].nunique()

    return_products = inventory.loc[
        inventory["Return Stock"].gt(0),
        "Product",
    ].nunique()

    dashboard_cards = st.columns(6)

    with dashboard_cards[0]:
        kpi(
            "SKU variants",
            fmt_int(total_skus),
            "Valid SKU rows",
            "dark",
        )

    with dashboard_cards[1]:
        kpi(
            "Products",
            fmt_int(total_products),
            "Case-normalized styles",
            "blue",
        )

    with dashboard_cards[2]:
        kpi(
            "Current stock",
            fmt_int(total_stock),
            "Based on Total Stock",
            "teal",
        )

    with dashboard_cards[3]:
        kpi(
            "Return stock",
            fmt_int(total_returns),
            "Tracked separately",
            "red",
        )

    with dashboard_cards[4]:
        kpi(
            "Low stock",
            fmt_int(low_count),
            "Variants with 1–2 units",
            "amber",
        )

    with dashboard_cards[5]:
        kpi(
            "Out of stock",
            fmt_int(out_count),
            "Zero-total-stock variants",
            "red",
        )

    st.markdown(
        "<br>",
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <div class="notice">
            <strong>Clean inventory view:</strong>
            duplicate capitalization is grouped, size labels such as
            XXL/2XL and XXXL/3XL are normalized, and the source Total
            row is excluded from calculations.
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <div class="section-title">
            Return Control Snapshot
        </div>
        """,
        unsafe_allow_html=True,
    )

    return_cards = st.columns(4)

    with return_cards[0]:
        kpi(
            "Return units",
            fmt_int(total_returns),
            "Marked Return Stock",
            "red",
        )

    with return_cards[1]:
        kpi(
            "Return value",
            fmt_money(return_value),
            "At listed price",
            "red",
        )

    with return_cards[2]:
        kpi(
            "Affected SKUs",
            fmt_int(return_skus),
            "Variants with returns",
            "amber",
        )

    with return_cards[3]:
        kpi(
            "Affected products",
            fmt_int(return_products),
            "Product styles",
            "purple",
        )

    st.markdown(
        f"""
        <div class="return-banner">
            <strong>Return visibility:</strong>
            {fmt_int(total_returns)} return units across
            {fmt_int(return_skus)} SKU variants are tracked
            separately and can be inspected in the dedicated
            Return Analysis page.
        </div>
        """,
        unsafe_allow_html=True,
    )

    by_category = (
        inventory.groupby(
            "Category",
            as_index=False,
        )
        .agg(
            **{
                "Total Stock": (
                    "Total Stock",
                    "sum",
                ),
                "Returns": (
                    "Return Stock",
                    "sum",
                ),
                "Products": (
                    "Product",
                    "nunique",
                ),
                "Variants": (
                    "SKU",
                    "nunique",
                ),
            }
        )
        .sort_values(
            "Total Stock",
            ascending=False,
        )
    )

    top_products = (
        inventory.groupby(
            ["Category", "Product"],
            as_index=False,
        )["Total Stock"]
        .sum()
        .sort_values(
            "Total Stock",
            ascending=False,
        )
        .head(12)
    )

    top_products["Label"] = (
        top_products["Product"]
        + " · "
        + top_products["Category"]
    )

    left, right = st.columns([1.2, 1])

    with left:
        category_figure = bar_chart(
            by_category,
            "Category",
            "Total Stock",
            "Current Stock by Category (Total Stock)",
        )

        if category_figure is not None:
            category_figure.update_xaxes(
                tickangle=-30
            )

            st.plotly_chart(
                category_figure,
                width="stretch",
            )

    with right:
        st.plotly_chart(
            stock_status_chart(inventory),
            width="stretch",
        )

    left, right = st.columns([1.2, 1])

    with left:
        st.plotly_chart(
            bar_chart(
                top_products.sort_values(
                    "Total Stock"
                ),
                "Label",
                "Total Stock",
                "Top Products by Current Stock",
                orientation="h",
            ),
            width="stretch",
        )

    with right:
        st.plotly_chart(
            bar_chart(
                size_summary(inventory),
                "Size",
                "Total Stock",
                "Current Stock by Size",
            ),
            width="stretch",
        )

    st.markdown(
        """
        <div class="section-title">
            Category Inventory Summary
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.dataframe(
        by_category,
        width="stretch",
        hide_index=True,
    )


# ============================================================
# PREMIUM SEARCH AND SIZE ANALYSIS
# ============================================================

with tabs[1]:
    st.markdown(
        """
        <div class="section-title">
            Premium Product Search & Size Analysis
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <div class="section-sub">
            Use the filters once. Every result, KPI, size chart,
            stock matrix and return analysis below automatically
            follows these filters.
        </div>
        """,
        unsafe_allow_html=True,
    )

    # --------------------------------------------------------
    # ONLY FILTER AREA ON THIS PAGE
    # --------------------------------------------------------

    query = st.text_input(
        "Search product / SKU / colour / size",
        placeholder=(
            "Example: Classic Polo, VB-PR-CP, Black, 2XL"
        ),
    )

    f1, f2, f3, f4, f5 = st.columns(
        [1, 1.3, 1, 1, 1.1]
    )

    with f1:
        selected_categories = st.multiselect(
            "Category",
            sorted(
                inventory["Category"].unique()
            ),
            placeholder="All categories",
        )

    if selected_categories:
        product_pool = inventory[
            inventory["Category"].isin(
                selected_categories
            )
        ]
    else:
        product_pool = inventory

    with f2:
        selected_products = st.multiselect(
            "Product",
            sorted(
                product_pool["Product"].unique()
            ),
            placeholder="All products",
        )

    with f3:
        selected_sizes = st.multiselect(
            "Size",
            order_sizes(product_pool["Size"]),
            placeholder="All sizes",
        )

    with f4:
        selected_colors = st.multiselect(
            "Colour",
            sorted(
                product_pool["Color"].unique()
            ),
            placeholder="All colours",
        )

    with f5:
        selected_stock = st.multiselect(
            "Stock status",
            [
                "Healthy (3+)",
                "Low Stock (1-2)",
                "Out of Stock",
            ],
            placeholder="All stock",
        )

    filtered = filter_inventory(
        inventory,
        search_text=query,
        categories=selected_categories,
        products=selected_products,
        sizes=selected_sizes,
        colors=selected_colors,
        stock_states=selected_stock,
    )

    st.markdown(
        f"""
        <div class="filter-summary">
            <b>{len(filtered):,}</b> matching SKU rows ·
            <b>{filtered["Product"].nunique():,}</b> products ·
            <b>{fmt_int(filtered["Total Stock"].sum())}</b>
            current units from Total Stock
        </div>
        """,
        unsafe_allow_html=True,
    )

    if filtered.empty:
        st.warning(
            "No products match the selected filters."
        )

    else:
        # All sections below use the same filtered DataFrame.
        product_data = filtered.copy()

        matching_products = sorted(
            product_data[
                "Product"
            ].dropna().unique()
        )

        if len(matching_products) == 1:
            analysis_title = matching_products[0]

        elif selected_products:
            analysis_title = (
                f"{len(matching_products)} "
                "Selected Products"
            )

        elif query.strip():
            analysis_title = (
                f'Search Results for “{query.strip()}”'
            )

        else:
            analysis_title = "All Filtered Products"

        # ----------------------------------------------------
        # FILTERED KPI SUMMARY
        # ----------------------------------------------------

        search_cards = st.columns(6)

        with search_cards[0]:
            kpi(
                "Matches",
                fmt_int(len(product_data)),
                "SKU rows",
                "dark",
            )

        with search_cards[1]:
            kpi(
                "Products",
                fmt_int(
                    product_data[
                        "Product"
                    ].nunique()
                ),
                "Matching styles",
                "blue",
            )

        with search_cards[2]:
            kpi(
                "Sizes",
                fmt_int(
                    product_data[
                        "Size"
                    ].nunique()
                ),
                "Normalized sizes",
                "purple",
            )

        with search_cards[3]:
            kpi(
                "Current stock",
                fmt_int(
                    product_data[
                        "Total Stock"
                    ].sum()
                ),
                "Based on Total Stock",
                "teal",
            )

        with search_cards[4]:
            kpi(
                "Returns",
                fmt_int(
                    product_data[
                        "Return Stock"
                    ].sum()
                ),
                "Return units",
                "red",
            )

        with search_cards[5]:
            kpi(
                "Stock value",
                fmt_money(
                    product_data[
                        "Stock Value"
                    ].sum()
                ),
                "Total Stock × Price",
                "amber",
            )

        # ----------------------------------------------------
        # FILTERED PRODUCT RESULTS
        # ----------------------------------------------------

        st.markdown(
            """
            <div class="section-title">
                Filtered Product Results
            </div>
            """,
            unsafe_allow_html=True,
        )

        result_columns = [
            "Category",
            "Product",
            "SKU",
            "Size",
            "Color",
            "Price",
            "Total Stock",
            "Return Stock",
            "Stock Status",
            "Brand",
        ]

        result = (
            product_data
            .sort_values(
                [
                    "Product",
                    "Size Rank",
                    "Color",
                ]
            )
            [result_columns]
        )

        st.dataframe(
            result,
            width="stretch",
            hide_index=True,
            height=560,
            column_config={
                "Price": (
                    st.column_config.NumberColumn(
                        format="৳%d"
                    )
                ),
                "Total Stock": (
                    st.column_config.NumberColumn(
                        format="%d"
                    )
                ),
                "Return Stock": (
                    st.column_config.NumberColumn(
                        format="%d"
                    )
                ),
            },
        )

        st.download_button(
            "Download filtered inventory (CSV)",
            result.to_csv(
                index=False
            ).encode("utf-8-sig"),
            "Vibes_Filtered_Inventory.csv",
            "text/csv",
        )

        # ----------------------------------------------------
        # PRODUCT AND SIZE ANALYSIS
        # NO SECOND FILTER OR PRODUCT SELECTOR
        # ----------------------------------------------------

        st.markdown(
            """
            <div class="section-title">
                Product & Size Analysis
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.markdown(
            """
            <div class="section-sub">
                This section uses exactly the same filtered
                results shown above. There is no separate
                product or size filter.
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.markdown(
            f"""
            <div class="filter-summary">
                Analysing <b>{analysis_title}</b> ·
                <b>{len(product_data):,}</b> SKU rows ·
                <b>{len(matching_products):,}</b> products ·
                <b>{fmt_int(product_data["Total Stock"].sum())}</b>
                current units
            </div>
            """,
            unsafe_allow_html=True,
        )

        analysis_cards = st.columns(6)

        with analysis_cards[0]:
            kpi(
                "Products",
                fmt_int(
                    product_data[
                        "Product"
                    ].nunique()
                ),
                "Filtered products",
                "blue",
            )

        with analysis_cards[1]:
            kpi(
                "Current stock",
                fmt_int(
                    product_data[
                        "Total Stock"
                    ].sum()
                ),
                "Based on Total Stock",
                "teal",
            )

        with analysis_cards[2]:
            kpi(
                "Sizes",
                fmt_int(
                    product_data[
                        "Size"
                    ].nunique()
                ),
                "Normalized sizes",
                "purple",
            )

        with analysis_cards[3]:
            kpi(
                "Colours",
                fmt_int(
                    product_data[
                        "Color"
                    ].nunique()
                ),
                "Filtered colours",
                "dark",
            )

        with analysis_cards[4]:
            kpi(
                "Returns",
                fmt_int(
                    product_data[
                        "Return Stock"
                    ].sum()
                ),
                "Return units",
                "red",
            )

        with analysis_cards[5]:
            kpi(
                "Out of stock",
                fmt_int(
                    product_data[
                        "Total Stock"
                    ].eq(0).sum()
                ),
                "Zero-stock variants",
                "red",
            )

        # ----------------------------------------------------
        # SIZE ANALYSIS AND COLOUR-SIZE MATRIX
        # ----------------------------------------------------

        size_data = size_summary(product_data)

        stock_matrix = color_size_matrix(
            product_data
        )

        left, right = st.columns([1, 1.25])

        with left:
            size_figure = bar_chart(
                size_data,
                "Size",
                "Total Stock",
                (
                    f"{analysis_title} · "
                    "Current Stock by Size"
                ),
            )

            if size_figure is not None:
                st.plotly_chart(
                    size_figure,
                    width="stretch",
                )

            st.dataframe(
                size_data,
                width="stretch",
                hide_index=True,
                column_config={
                    "Total Stock": (
                        st.column_config.NumberColumn(
                            format="%d"
                        )
                    ),
                    "Returns": (
                        st.column_config.NumberColumn(
                            format="%d"
                        )
                    ),
                    "Variants": (
                        st.column_config.NumberColumn(
                            format="%d"
                        )
                    ),
                    "Out of Stock": (
                        st.column_config.NumberColumn(
                            format="%d"
                        )
                    ),
                },
            )

        with right:
            st.markdown(
                f"**Colour × Size Matrix — "
                f"{analysis_title}**"
            )

            st.dataframe(
                matrix_style(stock_matrix),
                width="stretch",
            )

            heatmap_body = (
                stock_matrix
                .drop(
                    index="TOTAL",
                    errors="ignore",
                )
                .drop(
                    columns="Total",
                    errors="ignore",
                )
            )

            if not heatmap_body.empty:
                heatmap = go.Figure(
                    go.Heatmap(
                        z=heatmap_body.values,
                        x=list(
                            heatmap_body.columns
                        ),
                        y=list(
                            heatmap_body.index
                        ),
                        text=heatmap_body.values,
                        texttemplate="%{text}",
                        colorscale=[
                            [0, "#F8FAFC"],
                            [.45, "#9FD8D1"],
                            [1, "#146B63"],
                        ],
                        hovertemplate=(
                            "Colour: %{y}<br>"
                            "Size: %{x}<br>"
                            "Total Stock: %{z}"
                            "<extra></extra>"
                        ),
                    )
                )

                heatmap.update_layout(
                    title=(
                        f"{analysis_title} · "
                        "Colour and Size Stock"
                    ),
                    template="plotly_white",
                    height=max(
                        350,
                        42 * len(
                            heatmap_body
                        ) + 130,
                    ),
                    margin=dict(
                        l=20,
                        r=20,
                        t=55,
                        b=30,
                    ),
                )

                st.plotly_chart(
                    heatmap,
                    width="stretch",
                )

        # ----------------------------------------------------
        # PRODUCT × SIZE MATRIX
        # ----------------------------------------------------

        st.markdown(
            "**Filtered Product × Size Stock**"
        )

        filtered_product_matrix = (
            product_size_matrix(
                product_data
            )
        )

        st.dataframe(
            matrix_style(
                filtered_product_matrix
            ),
            width="stretch",
            height=min(
                650,
                max(
                    250,
                    75
                    + len(
                        filtered_product_matrix
                    ) * 36,
                ),
            ),
        )

        # ----------------------------------------------------
        # FILTERED RETURN MATRIX
        # ----------------------------------------------------

        if (
            product_data[
                "Return Stock"
            ].sum() > 0
        ):
            st.markdown(
                "**Filtered Return Stock — "
                "Colour × Size**"
            )

            filtered_return_matrix = (
                color_size_matrix(
                    product_data,
                    "Return Stock",
                )
            )

            st.dataframe(
                matrix_style(
                    filtered_return_matrix
                ),
                width="stretch",
            )

        # ----------------------------------------------------
        # COMPLETE FILTERED VARIANTS
        # ----------------------------------------------------

        st.markdown(
            "**Complete filtered product variants**"
        )

        variant_columns = [
            "Category",
            "Product",
            "SKU",
            "Size",
            "Color",
            "Price",
            "Total Stock",
            "Return Stock",
            "Return Value",
            "Stock Status",
            "Return Status",
        ]

        variant_data = (
            product_data
            .sort_values(
                [
                    "Product",
                    "Size Rank",
                    "Color",
                ]
            )
            [variant_columns]
        )

        st.dataframe(
            variant_data,
            width="stretch",
            hide_index=True,
            height=480,
            column_config={
                "Price": (
                    st.column_config.NumberColumn(
                        format="৳%d"
                    )
                ),
                "Total Stock": (
                    st.column_config.NumberColumn(
                        format="%d"
                    )
                ),
                "Return Stock": (
                    st.column_config.NumberColumn(
                        format="%d"
                    )
                ),
                "Return Value": (
                    st.column_config.NumberColumn(
                        format="৳%d"
                    )
                ),
            },
        )


# ============================================================
# CATEGORY STOCK
# ============================================================

with tabs[2]:
    st.markdown(
        """
        <div class="section-title">
            Category & Product Stock
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <div class="section-sub">
            Compare every category and product across all
            normalized sizes.
        </div>
        """,
        unsafe_allow_html=True,
    )

    selected_category = st.selectbox(
        "Filter category",
        [
            "All Categories",
            *sorted(
                inventory[
                    "Category"
                ].unique()
            ),
        ],
        key="category_table",
    )

    if selected_category == "All Categories":
        category_data = inventory
    else:
        category_data = inventory[
            inventory["Category"].eq(
                selected_category
            )
        ]

    category_matrix = product_size_matrix(
        category_data
    )

    st.dataframe(
        matrix_style(category_matrix),
        width="stretch",
        height=650,
    )

    product_totals = (
        category_data.groupby(
            ["Category", "Product"],
            as_index=False,
        )
        .agg(
            **{
                "Total Stock": (
                    "Total Stock",
                    "sum",
                ),
                "Returns": (
                    "Return Stock",
                    "sum",
                ),
                "Return_Value": (
                    "Return Value",
                    "sum",
                ),
                "Variants": (
                    "SKU",
                    "nunique",
                ),
                "Sizes": (
                    "Size",
                    "nunique",
                ),
            }
        )
        .sort_values(
            "Total Stock",
            ascending=False,
        )
    )

    st.dataframe(
        product_totals,
        width="stretch",
        hide_index=True,
        column_config={
            "Return_Value": (
                st.column_config.NumberColumn(
                    "Return Value",
                    format="৳%d",
                )
            )
        },
    )

    category_product_figure = bar_chart(
        product_totals
        .head(20)
        .sort_values("Total Stock"),
        "Product",
        "Total Stock",
        "Top Products in Selected Category",
        orientation="h",
    )

    if category_product_figure is not None:
        st.plotly_chart(
            category_product_figure,
            width="stretch",
        )


# ============================================================
# RETURN ANALYSIS
# ============================================================

with tabs[3]:
    st.markdown(
        """
        <div class="section-title">
            Premium Return Analysis
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <div class="section-sub">
            Isolate every return-marked unit by category,
            product, colour, size and SKU.
        </div>
        """,
        unsafe_allow_html=True,
    )

    returns = inventory[
        inventory["Return Stock"].gt(0)
    ].copy()

    return_cards = st.columns(5)

    with return_cards[0]:
        kpi(
            "Return units",
            fmt_int(
                returns[
                    "Return Stock"
                ].sum()
            ),
            "Exact return quantity",
            "red",
        )

    with return_cards[1]:
        kpi(
            "Return value",
            fmt_money(
                returns[
                    "Return Value"
                ].sum()
            ),
            "At listed price",
            "red",
        )

    with return_cards[2]:
        kpi(
            "Affected SKUs",
            fmt_int(
                returns[
                    "SKU"
                ].nunique()
            ),
            "Return variants",
            "amber",
        )

    with return_cards[3]:
        kpi(
            "Products",
            fmt_int(
                returns[
                    "Product"
                ].nunique()
            ),
            "Affected styles",
            "purple",
        )

    with return_cards[4]:
        kpi(
            "Categories",
            fmt_int(
                returns[
                    "Category"
                ].nunique()
            ),
            "Affected groups",
            "dark",
        )

    if returns.empty:
        st.success(
            "No return stock is currently "
            "marked in this workbook."
        )

        st.plotly_chart(
            return_status_chart(inventory),
            width="stretch",
        )

    else:
        rf1, rf2, rf3 = st.columns(
            [1, 1.4, 1]
        )

        with rf1:
            return_categories = st.multiselect(
                "Return category",
                sorted(
                    returns[
                        "Category"
                    ].unique()
                ),
                placeholder="All categories",
            )

        if return_categories:
            return_pool = returns[
                returns["Category"].isin(
                    return_categories
                )
            ]
        else:
            return_pool = returns

        with rf2:
            return_products = st.multiselect(
                "Return product",
                sorted(
                    return_pool[
                        "Product"
                    ].unique()
                ),
                placeholder="All products",
            )

        with rf3:
            return_sizes = st.multiselect(
                "Return size",
                order_sizes(
                    return_pool["Size"]
                ),
                placeholder="All sizes",
            )

        return_view = filter_inventory(
            return_pool,
            products=return_products,
            sizes=return_sizes,
        )

        left, right = st.columns([1.2, 1])

        return_by_product = (
            return_view.groupby(
                ["Category", "Product"],
                as_index=False,
            )
            .agg(
                Returns=(
                    "Return Stock",
                    "sum",
                ),
                Return_Value=(
                    "Return Value",
                    "sum",
                ),
                SKUs=(
                    "SKU",
                    "nunique",
                ),
            )
            .sort_values(
                "Returns",
                ascending=False,
            )
        )

        return_by_product["Label"] = (
            return_by_product["Product"]
            + " · "
            + return_by_product["Category"]
        )

        with left:
            return_figure = bar_chart(
                return_by_product
                .head(20)
                .sort_values("Returns"),
                "Label",
                "Returns",
                "Return Units by Product",
                orientation="h",
            )

            if return_figure is not None:
                return_figure.update_traces(
                    marker_color="#E05252"
                )

                st.plotly_chart(
                    return_figure,
                    width="stretch",
                )

        with right:
            st.plotly_chart(
                return_status_chart(inventory),
                width="stretch",
            )

        st.markdown(
            "**Return Colour × Size Matrix**"
        )

        st.dataframe(
            matrix_style(
                color_size_matrix(
                    return_view,
                    "Return Stock",
                )
            ),
            width="stretch",
        )

        return_columns = [
            "Category",
            "Product",
            "SKU",
            "Size",
            "Color",
            "Price",
            "Total Stock",
            "Return Stock",
            "Return Value",
            "Brand",
        ]

        return_export = (
            return_view
            .sort_values(
                [
                    "Product",
                    "Size Rank",
                ]
            )
            [return_columns]
        )

        st.dataframe(
            return_export,
            width="stretch",
            hide_index=True,
            height=520,
            column_config={
                "Price": (
                    st.column_config.NumberColumn(
                        format="৳%d"
                    )
                ),
                "Total Stock": (
                    st.column_config.NumberColumn(
                        format="%d"
                    )
                ),
                "Return Stock": (
                    st.column_config.NumberColumn(
                        format="%d"
                    )
                ),
                "Return Value": (
                    st.column_config.NumberColumn(
                        format="৳%d"
                    )
                ),
            },
        )

        st.download_button(
            "Download return inventory (CSV)",
            return_export.to_csv(
                index=False
            ).encode("utf-8-sig"),
            "Vibes_Return_Inventory.csv",
            "text/csv",
        )


# ============================================================
# COMPLETE SKU DATA
# ============================================================

with tabs[4]:
    st.markdown(
        """
        <div class="section-title">
            Complete SKU Data
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <div class="section-sub">
            Searchable audit table behind every portal number.
        </div>
        """,
        unsafe_allow_html=True,
    )

    raw_query = st.text_input(
        "Quick search",
        placeholder=(
            "Product, SKU, colour, "
            "size or category"
        ),
        key="raw_search",
    )

    raw_view = filter_inventory(
        inventory,
        search_text=raw_query,
    )

    raw_columns = [
        "Category",
        "Product",
        "Name",
        "SKU",
        "Size",
        "Color",
        "Price",
        "Status",
        "Stock",
        "WI Stock",
        "WI Out Stock",
        "Total Stock",
        "Return Stock",
        "Return Value",
        "Stock Status",
        "Return Status",
        "Brand",
    ]

    st.dataframe(
        raw_view[raw_columns],
        width="stretch",
        hide_index=True,
        height=690,
        column_config={
            "Price": (
                st.column_config.NumberColumn(
                    format="৳%d"
                )
            ),
            "Total Stock": (
                st.column_config.NumberColumn(
                    format="%d"
                )
            ),
            "Return Stock": (
                st.column_config.NumberColumn(
                    format="%d"
                )
            ),
            "Return Value": (
                st.column_config.NumberColumn(
                    format="৳%d"
                )
            ),
        },
    )

    st.download_button(
        "Download complete inventory (CSV)",
        raw_view[raw_columns]
        .to_csv(index=False)
        .encode("utf-8-sig"),
        "Vibes_Complete_Inventory.csv",
        "text/csv",
    )


st.caption(
    "VIBES × Prelegent · Every current-stock KPI, chart, "
    "filter and matrix is based exclusively on Total Stock. "
    "Returns use Return Stock."
)

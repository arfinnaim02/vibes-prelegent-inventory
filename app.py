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
DEFAULT_FILE = APP_DIR / "data" / "Prelegent_Inventory_Analysis.xlsx"

st.set_page_config(
    page_title="VIBES × Prelegent Inventory Portal",
    page_icon="◆",
    layout="wide",
    initial_sidebar_state="expanded",
)

# -----------------------------
# Premium UI
# -----------------------------
st.markdown(
    """
<style>
:root {
  --ink:#101828; --muted:#667085; --line:#E4E7EC; --paper:#FFFFFF;
  --bg:#F5F7FA; --green:#067647; --green-bg:#ECFDF3;
  --red:#B42318; --red-bg:#FEF3F2; --amber:#B54708; --amber-bg:#FFFAEB;
  --blue:#175CD3; --blue-bg:#EFF8FF; --purple:#6941C6; --purple-bg:#F9F5FF;
}
.stApp { background: var(--bg); }
.block-container { padding-top: 1.4rem; padding-bottom: 3rem; max-width: 1600px; }
[data-testid="stSidebar"] { background: #0B0F19; }
[data-testid="stSidebar"] * { color: #F9FAFB; }
[data-testid="stSidebar"] .stSelectbox label,
[data-testid="stSidebar"] .stSlider label,
[data-testid="stSidebar"] .stFileUploader label { color:#F9FAFB !important; }
[data-testid="stSidebar"] .stCaptionContainer { color:#D0D5DD !important; }

.portal-title { font-size: 2rem; font-weight: 850; letter-spacing:-0.04em; color:#101828; margin-bottom:0.15rem; }
.portal-sub { color:#667085; font-size:0.98rem; margin-bottom:1.2rem; }
.eyebrow { color:#475467; font-size:0.73rem; text-transform:uppercase; letter-spacing:.11em; font-weight:750; }
.section-title { font-size:1.25rem; font-weight:800; color:#101828; margin:1.25rem 0 .55rem; }
.section-sub { color:#667085; margin-top:-.35rem; margin-bottom:.8rem; font-size:.9rem; }

.kpi-card { background:#fff; border:1px solid #EAECF0; border-radius:18px; padding:18px 18px 15px; min-height:122px; box-shadow:0 1px 2px rgba(16,24,40,.03); }
.kpi-label { font-size:.78rem; color:#667085; text-transform:uppercase; letter-spacing:.055em; font-weight:700; }
.kpi-value { margin-top:8px; font-size:1.75rem; font-weight:850; color:#101828; letter-spacing:-.035em; }
.kpi-note { margin-top:5px; font-size:.78rem; color:#98A2B3; }
.kpi-green { border-top:4px solid #12B76A; }
.kpi-blue { border-top:4px solid #2E90FA; }
.kpi-red { border-top:4px solid #F04438; }
.kpi-amber { border-top:4px solid #F79009; }
.kpi-purple { border-top:4px solid #7F56D9; }
.kpi-dark { border-top:4px solid #101828; }

.notice { background:#fff; border:1px solid #EAECF0; border-radius:14px; padding:13px 15px; color:#475467; font-size:.88rem; }
.notice strong { color:#101828; }
.success-box { background:#ECFDF3; border:1px solid #ABEFC6; border-radius:14px; padding:13px 15px; color:#067647; font-size:.88rem; }
.danger-box { background:#FEF3F2; border:1px solid #FECDCA; border-radius:14px; padding:13px 15px; color:#B42318; font-size:.88rem; }
.action-pill { display:inline-block; border-radius:999px; padding:5px 9px; font-size:.73rem; font-weight:800; }
.pill-send {background:#EFF8FF;color:#175CD3}.pill-return{background:#FEF3F2;color:#B42318}.pill-move{background:#FFFAEB;color:#B54708}.pill-hold{background:#ECFDF3;color:#067647}

div[data-testid="stDataFrame"] { border:1px solid #EAECF0; border-radius:14px; overflow:hidden; background:white; }
div[data-testid="stPlotlyChart"] { background:white; border:1px solid #EAECF0; border-radius:16px; padding:6px; }
.stTabs [data-baseweb="tab-list"] { gap:8px; }
.stTabs [data-baseweb="tab"] { border-radius:10px; padding:8px 14px; background:#fff; border:1px solid #EAECF0; }
.stTabs [aria-selected="true"] { background:#101828 !important; color:#fff !important; }
hr { border:none; border-top:1px solid #EAECF0; margin:1.1rem 0; }
</style>
""",
    unsafe_allow_html=True,
)

NUMERIC_COLS = [
    "Price", "Discount Price", "Stock", "WI Stock", "WI Out Stock",
    "Salable Stock", "Return Stock", "Effective Available", "Planning Send Gap"
]
REQUIRED_CORE = ["Normalized Product", "Category", "Size", "Color", "Salable Stock", "WI Stock", "Return Stock"]


def fmt_int(v: float | int) -> str:
    return f"{int(round(float(v))):,}"


def fmt_money(v: float | int) -> str:
    return f"৳{float(v):,.0f}"


def kpi(label: str, value: str, note: str, tone: str = "dark") -> None:
    st.markdown(
        f'<div class="kpi-card kpi-{tone}"><div class="kpi-label">{label}</div>'
        f'<div class="kpi-value">{value}</div><div class="kpi-note">{note}</div></div>',
        unsafe_allow_html=True,
    )


def natural_size_key(value: object):
    s = str(value).strip().upper()
    aliases = {"N/A": 0, "NA": 0, "FREE": 1, "FREESIZE": 1, "XS": 10, "S": 20, "M": 30,
               "L": 40, "XL": 50, "2XL": 60, "XXL": 60, "3XL": 70, "XXXL": 70,
               "4XL": 80, "XXXXL": 80, "5XL": 90}
    if s in aliases:
        return (0, aliases[s], s)
    nums = re.findall(r"\d+(?:\.\d+)?", s)
    if nums:
        return (1, float(nums[0]), s)
    return (2, 9999, s)


def order_sizes(values: Iterable[object]) -> list[str]:
    vals = [str(v).strip() for v in values if pd.notna(v) and str(v).strip()]
    return sorted(dict.fromkeys(vals), key=natural_size_key)


@st.cache_data(show_spinner=False)
def load_inventory_from_bytes(file_bytes: bytes, filename: str) -> pd.DataFrame:
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    preferred = "SKU Detail" if "SKU Detail" in xls.sheet_names else None
    if preferred:
        raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=preferred, header=3)
    else:
        raw = None
        for sh in xls.sheet_names:
            candidate = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sh)
            if {"Salable Stock", "WI Stock", "Return Stock"}.issubset(set(candidate.columns.astype(str))):
                raw = candidate
                break
        if raw is None:
            raise ValueError("Could not find a sheet with Salable Stock, WI Stock and Return Stock columns.")
    return prepare_df(raw)


@st.cache_data(show_spinner=False)
def load_inventory_from_path(path_str: str) -> pd.DataFrame:
    path = Path(path_str)
    xls = pd.ExcelFile(path)
    if "SKU Detail" in xls.sheet_names:
        raw = pd.read_excel(path, sheet_name="SKU Detail", header=3)
    else:
        raw = None
        for sh in xls.sheet_names:
            candidate = pd.read_excel(path, sheet_name=sh)
            if {"Salable Stock", "WI Stock", "Return Stock"}.issubset(set(candidate.columns.astype(str))):
                raw = candidate
                break
        if raw is None:
            raise ValueError("Could not find a usable inventory sheet.")
    return prepare_df(raw)


def prepare_df(raw: pd.DataFrame) -> pd.DataFrame:
    df = raw.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all")

    if "Normalized Product" not in df.columns:
        if "Product" in df.columns:
            df["Normalized Product"] = df["Product"]
        elif "Original Name" in df.columns:
            df["Normalized Product"] = df["Original Name"]
        else:
            df["Normalized Product"] = "Unclassified Product"
    if "Category" not in df.columns:
        df["Category"] = "Unclassified"
    if "Original Name" not in df.columns:
        df["Original Name"] = df["Normalized Product"]
    if "Size" not in df.columns:
        df["Size"] = "N/A"
    if "Color" not in df.columns:
        df["Color"] = "N/A"
    if "SKU" not in df.columns:
        df["SKU"] = ""
    if "Price" not in df.columns:
        df["Price"] = 0
    if "WI Out Stock" not in df.columns:
        df["WI Out Stock"] = 0
    if "Replenish Eligible" not in df.columns:
        df["Replenish Eligible"] = "YES"

    for c in NUMERIC_COLS:
        if c not in df.columns:
            df[c] = 0
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    for c in ["Normalized Product", "Category", "Original Name", "Size", "Color", "SKU", "Replenish Eligible"]:
        df[c] = df[c].fillna("N/A").astype(str).str.strip()

    df["Mall Stock"] = df["Salable Stock"]
    df["Warehouse Stock"] = df["WI Stock"]
    df["Current Physical"] = df["Salable Stock"] + df["WI Stock"]
    df["Return Value"] = df["Return Stock"] * df["Price"]
    df["Mall Retail Value"] = df["Salable Stock"] * df["Price"]
    df["Physical Retail Value"] = df["Current Physical"] * df["Price"]
    df["Return Block"] = (df["Return Stock"] > 0) | (~df["Replenish Eligible"].str.upper().eq("YES"))
    return df


def add_scenario_columns(df: pd.DataFrame, target_floor: int, mall_floor: int) -> pd.DataFrame:
    x = df.copy()
    # First use stock already at Prelegent warehouse to fill mall shortages.
    x["Mall Gap"] = (mall_floor - x["Salable Stock"]).clip(lower=0)
    x["Move WI → Mall"] = x[["Mall Gap", "WI Stock"]].min(axis=1).where(~x["Return Block"], 0)
    # Then consider additional VIBES shipment against total stock already physically at Prelegent.
    x["Safe Send"] = (target_floor - x["Current Physical"]).clip(lower=0).where(~x["Return Block"], 0)
    x["Safe Send"] = x["Safe Send"].astype(int)
    x["Move WI → Mall"] = x["Move WI → Mall"].astype(int)
    x["Return Qty"] = x["Return Stock"].astype(int)

    conditions = [
        x["Return Block"],
        x["Move WI → Mall"] > 0,
        x["Safe Send"] > 0,
    ]
    choices = ["RETURN / DO NOT SEND", "MOVE WAREHOUSE → MALL", "SEND FROM VIBES"]
    x["Action"] = "HOLD / MAINTAIN"
    for cond, choice in reversed(list(zip(conditions, choices))):
        x.loc[cond, "Action"] = choice
    return x


def pivot_matrix(df: pd.DataFrame, metric: str) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()
    ordered = order_sizes(df["Size"].unique())
    p = pd.pivot_table(df, index="Color", columns="Size", values=metric, aggfunc="sum", fill_value=0)
    p = p.reindex(columns=[c for c in ordered if c in p.columns], fill_value=0)
    p["Total"] = p.sum(axis=1)
    total = p.sum(axis=0)
    p.loc["TOTAL"] = total
    return p.astype(int)


def matrix_styler(p: pd.DataFrame, tone: str = "blue"):
    palette = {
        "blue": ("#F5FAFF", "#175CD3"),
        "green": ("#F3FFF7", "#067647"),
        "red": ("#FFF5F4", "#B42318"),
        "amber": ("#FFFAEB", "#B54708"),
        "purple": ("#F9F5FF", "#6941C6"),
    }
    bg, fg = palette[tone]
    return (
        p.style
        .format("{:,.0f}")
        .background_gradient(cmap="Blues" if tone == "blue" else "Greens" if tone == "green" else "Reds" if tone == "red" else "Purples", axis=None)
        .set_properties(**{"text-align":"center", "font-weight":"600"})
        .set_table_styles([
            {"selector":"th", "props":[("background-color",fg),("color","white"),("font-weight","800"),("text-align","center")]},
            {"selector":"td", "props":[("border-color","#EAECF0")]},
        ])
    )


def plot_bar(data: pd.DataFrame, x: str, y: str, title: str, color: str | None = None, orientation="v"):
    if data.empty:
        return None
    if orientation == "h":
        fig = px.bar(data, x=y, y=x, color=color, orientation="h", text_auto=True)
    else:
        fig = px.bar(data, x=x, y=y, color=color, text_auto=True)
    fig.update_layout(
        title=title, template="plotly_white", margin=dict(l=20,r=20,t=55,b=25),
        legend_title_text="", height=390, font=dict(family="Arial", color="#344054"),
        title_font=dict(size=17, color="#101828"), plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF"
    )
    fig.update_traces(marker_line_width=0, textposition="outside", cliponaxis=False)
    return fig


def plot_size_chart(df: pd.DataFrame, metric: str, title: str):
    if df.empty:
        return None
    g = df.groupby("Size", as_index=False)[metric].sum()
    order = order_sizes(g["Size"])
    g["_order"] = g["Size"].map({s:i for i,s in enumerate(order)})
    g = g.sort_values("_order")
    fig = px.bar(g, x="Size", y=metric, text_auto=True)
    fig.update_layout(title=title, template="plotly_white", height=360, margin=dict(l=20,r=20,t=55,b=25), showlegend=False)
    fig.update_traces(marker_color="#101828", textposition="outside", cliponaxis=False)
    return fig


def heatmap_matrix(p: pd.DataFrame, title: str):
    if p.empty:
        return None
    body = p.drop(index="TOTAL", errors="ignore").drop(columns="Total", errors="ignore")
    if body.empty:
        return None
    fig = go.Figure(data=go.Heatmap(
        z=body.values,
        x=[str(c) for c in body.columns],
        y=[str(i) for i in body.index],
        text=body.values,
        texttemplate="%{text}",
        colorscale=[[0,"#F8FAFC"],[0.5,"#D0D5DD"],[1,"#101828"]],
        colorbar=dict(title="Units"),
        hovertemplate="Colour: %{y}<br>Size: %{x}<br>Units: %{z}<extra></extra>"
    ))
    fig.update_layout(title=title, height=max(330, 48*len(body.index)+120), margin=dict(l=20,r=20,t=55,b=30), template="plotly_white")
    return fig


# -----------------------------
# Sidebar / data
# -----------------------------
with st.sidebar:
    st.markdown("## VIBES × PRELEGENT")
    st.caption("Premium Inventory Decision Portal")
    st.markdown("---")
    upload = st.file_uploader("Upload latest inventory workbook", type=["xlsx"])
    target_floor = st.slider("Target total stock / variant", min_value=1, max_value=10, value=2, step=1,
                             help="Minimum desired units across Mall + Prelegent Warehouse before VIBES sends more.")
    mall_floor = st.slider("Target mall floor / variant", min_value=0, max_value=5, value=1, step=1,
                           help="If mall is below this level and WI Stock exists, portal suggests moving warehouse stock to mall first.")
    st.markdown("---")
    st.caption("Locked logic")
    st.caption("• WI Out is movement, not current stock")
    st.caption("• Return-stock variants are blocked from send")
    st.caption("• Warehouse stock is used before new VIBES shipment")

try:
    if upload is not None:
        base_df = load_inventory_from_bytes(upload.getvalue(), upload.name)
        data_label = upload.name
    else:
        base_df = load_inventory_from_path(str(DEFAULT_FILE))
        data_label = DEFAULT_FILE.name
except Exception as exc:
    st.error(f"Could not load inventory: {exc}")
    st.stop()

df = add_scenario_columns(base_df, target_floor, mall_floor)

# Header
st.markdown('<div class="eyebrow">VIBES / CONSIGNMENT INVENTORY</div>', unsafe_allow_html=True)
st.markdown('<div class="portal-title">Prelegent Inventory Command Center</div>', unsafe_allow_html=True)
st.markdown(
    f'<div class="portal-sub">Board-level stock visibility + product/colour/size drill-down · Data: <b>{data_label}</b></div>',
    unsafe_allow_html=True,
)

pages = st.tabs(["Board View", "Product Explorer", "Size Analytics", "Return Control", "Safe Send Planner", "SKU Data"])

# -----------------------------
# 1. Board View
# -----------------------------
with pages[0]:
    mall = df["Salable Stock"].sum()
    wh = df["WI Stock"].sum()
    physical = df["Current Physical"].sum()
    wiout = df["WI Out Stock"].sum()
    ret = df["Return Stock"].sum()
    send = df["Safe Send"].sum()
    move = df["Move WI → Mall"].sum()
    blocked = int((df["Return Block"]).sum())

    c = st.columns(6)
    with c[0]: kpi("Mall stock", fmt_int(mall), "Current salable units", "blue")
    with c[1]: kpi("Warehouse", fmt_int(wh), "Current WI Stock", "purple")
    with c[2]: kpi("Physical stock", fmt_int(physical), "Mall + Warehouse", "dark")
    with c[3]: kpi("Return pool", fmt_int(ret), "Separated from active stock", "red")
    with c[4]: kpi("Safe send", fmt_int(send), f"Target floor = {target_floor}", "green")
    with c[5]: kpi("WI → Mall", fmt_int(move), f"Mall floor = {mall_floor}", "amber")

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(
        f'<div class="notice"><strong>Control check:</strong> WI Out Stock = <b>{fmt_int(wiout)}</b> units and stays separate from current physical inventory. '
        f'<b>{blocked:,}</b> variant rows are currently blocked from replenishment because they are return-marked or not eligible.</div>',
        unsafe_allow_html=True,
    )

    cat = df.groupby("Category", as_index=False).agg(
        Mall=("Salable Stock","sum"), Warehouse=("WI Stock","sum"), Return=("Return Stock","sum"),
        Safe_Send=("Safe Send","sum"), Move_to_Mall=("Move WI → Mall","sum"),
        Retail_Value=("Mall Retail Value","sum")
    )
    cat["Physical"] = cat["Mall"] + cat["Warehouse"]
    cat = cat.sort_values("Physical", ascending=False)

    left, right = st.columns([1.25, 1])
    with left:
        melt = cat.melt(id_vars="Category", value_vars=["Mall","Warehouse","Return"], var_name="Stock Type", value_name="Units")
        fig = px.bar(melt, x="Category", y="Units", color="Stock Type", barmode="stack", text_auto=True,
                     color_discrete_map={"Mall":"#175CD3","Warehouse":"#7F56D9","Return":"#D92D20"})
        fig.update_layout(title="Inventory Position by Category", template="plotly_white", height=430, margin=dict(l=20,r=20,t=55,b=80), legend_title_text="")
        fig.update_xaxes(tickangle=-35)
        st.plotly_chart(fig, use_container_width=True)
    with right:
        act = cat[["Category","Safe_Send","Move_to_Mall","Return"]].melt(id_vars="Category", var_name="Decision", value_name="Units")
        act = act[act["Units"] > 0]
        fig2 = px.bar(act, x="Units", y="Category", color="Decision", orientation="h", text_auto=True,
                      color_discrete_map={"Safe_Send":"#12B76A","Move_to_Mall":"#F79009","Return":"#F04438"})
        fig2.update_layout(title="Board Action Quantities", template="plotly_white", height=430, margin=dict(l=20,r=20,t=55,b=25), legend_title_text="")
        st.plotly_chart(fig2, use_container_width=True)

    st.markdown('<div class="section-title">Board Decision Table</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">Return is exact from source. Send is scenario-based and excludes return-marked variants.</div>', unsafe_allow_html=True)
    decision = cat.rename(columns={"Safe_Send":"Safe Send", "Move_to_Mall":"Move WI → Mall", "Retail_Value":"Mall Retail Value"})
    decision["Net VIBES Movement"] = decision["Safe Send"] - decision["Return"]
    st.dataframe(
        decision[["Category","Mall","Warehouse","Physical","Return","Move WI → Mall","Safe Send","Net VIBES Movement","Mall Retail Value"]],
        use_container_width=True, hide_index=True,
        column_config={"Mall Retail Value": st.column_config.NumberColumn(format="৳%d")}
    )

    top_return = cat.sort_values("Return", ascending=False).iloc[0]
    top_send = cat.sort_values("Safe_Send", ascending=False).iloc[0]
    top_move = cat.sort_values("Move_to_Mall", ascending=False).iloc[0]
    st.markdown('<div class="section-title">Simple Board Story</div>', unsafe_allow_html=True)
    a,b,c3 = st.columns(3)
    with a:
        st.markdown(f'<div class="danger-box"><b>1 · Return first</b><br>{top_return.Category}: {fmt_int(top_return.Return)} units marked to come back. These variants are automatically blocked from new-send recommendations.</div>', unsafe_allow_html=True)
    with b:
        st.markdown(f'<div class="success-box"><b>2 · Replenish selectively</b><br>{top_send.Category}: up to {fmt_int(top_send.Safe_Send)} safe-send units under the current target-floor scenario.</div>', unsafe_allow_html=True)
    with c3:
        st.markdown(f'<div class="notice"><b>3 · Use their warehouse first</b><br>{top_move.Category}: {fmt_int(top_move.Move_to_Mall)} units can potentially move from WI Stock to mall before VIBES sends more.</div>', unsafe_allow_html=True)

# -----------------------------
# 2. Product Explorer
# -----------------------------
with pages[1]:
    st.markdown('<div class="section-title">Product / Variant Explorer</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">Filter a product, then click a variant row to inspect the exact SKU, size, colour and action.</div>', unsafe_allow_html=True)

    fc1, fc2 = st.columns([1,2])
    categories = ["All Categories"] + sorted(df["Category"].dropna().unique().tolist())
    with fc1:
        selected_cat = st.selectbox("Category", categories, key="prod_cat")
    sub = df if selected_cat == "All Categories" else df[df["Category"] == selected_cat]
    products = sorted(sub["Normalized Product"].dropna().unique().tolist())
    with fc2:
        selected_product = st.selectbox("Product / Style", products, key="prod_style")

    prod = sub[sub["Normalized Product"] == selected_product].copy()
    mall = prod["Salable Stock"].sum(); wh = prod["WI Stock"].sum(); ret = prod["Return Stock"].sum(); send = prod["Safe Send"].sum(); move = prod["Move WI → Mall"].sum()
    kc = st.columns(5)
    with kc[0]: kpi("Mall", fmt_int(mall), "Salable", "blue")
    with kc[1]: kpi("Warehouse", fmt_int(wh), "WI Stock", "purple")
    with kc[2]: kpi("Return", fmt_int(ret), "Do not resend", "red")
    with kc[3]: kpi("Safe Send", fmt_int(send), "After exclusions", "green")
    with kc[4]: kpi("Move WI → Mall", fmt_int(move), "Use before sending", "amber")

    matrix_tabs = st.tabs(["Current Physical", "Mall", "Warehouse", "Return", "Safe Send"])
    matrix_specs = [
        ("Current Physical","blue"), ("Salable Stock","green"), ("WI Stock","purple"), ("Return Stock","red"), ("Safe Send","green")
    ]
    for tab, (metric,tone) in zip(matrix_tabs, matrix_specs):
        with tab:
            p = pivot_matrix(prod, metric)
            if p.empty:
                st.info("No data for this product.")
            else:
                l,r = st.columns([1.2,1])
                with l:
                    st.markdown(f"**Colour × Size Matrix — {metric}**")
                    st.dataframe(matrix_styler(p, tone), use_container_width=True)
                with r:
                    fig = heatmap_matrix(p, f"{selected_product} · {metric}")
                    if fig: st.plotly_chart(fig, use_container_width=True)

    st.markdown('<div class="section-title">Variant Inspector</div>', unsafe_allow_html=True)
    variant_cols = ["SKU","Color","Size","Salable Stock","WI Stock","Current Physical","Return Stock","Move WI → Mall","Safe Send","Action"]
    variants = prod[variant_cols].copy().sort_values(["Color","Size"])
    event = st.dataframe(
        variants,
        use_container_width=True,
        hide_index=True,
        on_select="rerun",
        selection_mode="single-row",
        key="variant_grid",
        column_config={
            "Action": st.column_config.TextColumn(width="large"),
            "Safe Send": st.column_config.NumberColumn(help="Automatically 0 if variant is in Return Stock."),
        },
    )
    selected_rows = event.selection.rows if hasattr(event, "selection") else []
    if selected_rows:
        row = variants.iloc[selected_rows[0]]
        st.markdown("---")
        d1,d2,d3,d4,d5,d6 = st.columns(6)
        with d1: kpi("SKU", str(row["SKU"]), "Selected variant", "dark")
        with d2: kpi("Colour", str(row["Color"]), "", "blue")
        with d3: kpi("Size", str(row["Size"]), "", "purple")
        with d4: kpi("Current", fmt_int(row["Current Physical"]), "Mall + WI", "dark")
        with d5: kpi("Return", fmt_int(row["Return Stock"]), "Blocked if > 0", "red")
        with d6: kpi("Safe Send", fmt_int(row["Safe Send"]), str(row["Action"]), "green" if row["Safe Send"] else "amber")
        if row["Return Stock"] > 0:
            st.markdown('<div class="danger-box"><b>RETURN LOCK:</b> This exact SKU / colour / size is in the return pool, so the portal will not recommend sending it again.</div>', unsafe_allow_html=True)

# -----------------------------
# 3. Size Analytics
# -----------------------------
with pages[2]:
    st.markdown('<div class="section-title">Size Analytics</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">See where stock is concentrated and which sizes are missing across any category or product.</div>', unsafe_allow_html=True)
    sc1,sc2,sc3 = st.columns([1,1.4,1])
    with sc1:
        size_cat = st.selectbox("Category", ["All Categories"] + sorted(df["Category"].unique().tolist()), key="size_cat")
    sdf = df if size_cat == "All Categories" else df[df["Category"] == size_cat]
    with sc2:
        size_products = ["All Products"] + sorted(sdf["Normalized Product"].unique().tolist())
        size_prod = st.selectbox("Product", size_products, key="size_prod")
    if size_prod != "All Products": sdf = sdf[sdf["Normalized Product"] == size_prod]
    with sc3:
        size_metric = st.selectbox("Metric", ["Current Physical","Salable Stock","WI Stock","Return Stock","Safe Send"], key="size_metric")

    l,r = st.columns([1.1,1])
    with l:
        fig = plot_size_chart(sdf, size_metric, f"{size_metric} by Size")
        if fig: st.plotly_chart(fig, use_container_width=True)
    with r:
        p = pivot_matrix(sdf, size_metric)
        fig = heatmap_matrix(p, f"Colour × Size · {size_metric}")
        if fig: st.plotly_chart(fig, use_container_width=True)

    size_summary = sdf.groupby("Size", as_index=False).agg(
        Mall=("Salable Stock","sum"), Warehouse=("WI Stock","sum"), Current=("Current Physical","sum"),
        Return=("Return Stock","sum"), Safe_Send=("Safe Send","sum")
    )
    order = order_sizes(size_summary["Size"])
    size_summary["_o"] = size_summary["Size"].map({s:i for i,s in enumerate(order)})
    size_summary = size_summary.sort_values("_o").drop(columns="_o")
    st.dataframe(size_summary, use_container_width=True, hide_index=True)

# -----------------------------
# 4. Return Control
# -----------------------------
with pages[3]:
    returns = df[df["Return Stock"] > 0].copy()
    st.markdown('<div class="section-title">Return Control Center</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">Returned inventory is isolated here and automatically blocked from replenishment.</div>', unsafe_allow_html=True)
    c = st.columns(4)
    with c[0]: kpi("Return units", fmt_int(returns["Return Stock"].sum()), "Exact source quantity", "red")
    with c[1]: kpi("Return value", fmt_money(returns["Return Value"].sum()), "At listed price", "red")
    with c[2]: kpi("Return products", fmt_int(returns["Normalized Product"].nunique()), "Styles affected", "purple")
    with c[3]: kpi("Blocked variants", fmt_int(len(returns)), "Safe Send forced to zero", "dark")

    if returns.empty:
        st.success("No return stock is currently marked.")
    else:
        rc1, rc2 = st.columns([1,1])
        bycat = returns.groupby("Category", as_index=False)["Return Stock"].sum().sort_values("Return Stock", ascending=False)
        byprod = returns.groupby(["Category","Normalized Product"], as_index=False)["Return Stock"].sum().sort_values("Return Stock", ascending=False)
        with rc1:
            fig = plot_bar(bycat, "Category", "Return Stock", "Return Units by Category", orientation="h")
            if fig: st.plotly_chart(fig, use_container_width=True)
        with rc2:
            top = byprod.head(15).copy()
            top["Label"] = top["Normalized Product"] + " · " + top["Category"]
            fig = plot_bar(top, "Label", "Return Stock", "Top Return Products", orientation="h")
            if fig: st.plotly_chart(fig, use_container_width=True)

        rcat = st.selectbox("Inspect return category", sorted(returns["Category"].unique().tolist()), key="ret_cat")
        rsub = returns[returns["Category"] == rcat]
        rprod = st.selectbox("Inspect return product", sorted(rsub["Normalized Product"].unique().tolist()), key="ret_prod")
        rp = rsub[rsub["Normalized Product"] == rprod]
        p = pivot_matrix(rp, "Return Stock")
        l,r = st.columns([1.2,1])
        with l: st.dataframe(matrix_styler(p, "red"), use_container_width=True)
        with r:
            fig = heatmap_matrix(p, f"Return Matrix · {rprod}")
            if fig: st.plotly_chart(fig, use_container_width=True)

        ret_export = returns[["Category","Normalized Product","SKU","Color","Size","Return Stock","Return Value","Action"]]
        st.download_button("Download Return List (CSV)", ret_export.to_csv(index=False).encode("utf-8-sig"), "Prelegent_Return_List.csv", "text/csv")

# -----------------------------
# 5. Safe Send Planner
# -----------------------------
with pages[4]:
    st.markdown('<div class="section-title">Safe Send Planner</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">A conservative send plan: returned variants are excluded, and existing Prelegent warehouse stock is considered first.</div>', unsafe_allow_html=True)

    plan = df[(df["Safe Send"] > 0) | (df["Move WI → Mall"] > 0) | (df["Return Stock"] > 0)].copy()
    pc = st.columns(4)
    with pc[0]: kpi("Send from VIBES", fmt_int(df["Safe Send"].sum()), f"Target total = {target_floor}", "green")
    with pc[1]: kpi("Move WI → Mall", fmt_int(df["Move WI → Mall"].sum()), f"Target mall = {mall_floor}", "amber")
    with pc[2]: kpi("Return to VIBES", fmt_int(df["Return Stock"].sum()), "Marked source stock", "red")
    with pc[3]: kpi("Net external units", fmt_int(df["Safe Send"].sum()-df["Return Stock"].sum()), "Send minus return", "dark")

    st.markdown('<div class="success-box"><b>Protection rule:</b> if Return Stock &gt; 0 for a variant, Safe Send is always 0 for that variant — even if its mall stock is empty.</div>', unsafe_allow_html=True)

    bycat = df.groupby("Category", as_index=False).agg(
        Mall=("Salable Stock","sum"), Warehouse=("WI Stock","sum"), Move_WI_to_Mall=("Move WI → Mall","sum"),
        Safe_Send=("Safe Send","sum"), Return=("Return Stock","sum")
    )
    bycat["Net External"] = bycat["Safe_Send"] - bycat["Return"]
    st.dataframe(bycat.sort_values(["Safe_Send","Return"], ascending=False), use_container_width=True, hide_index=True)

    st.markdown("**SKU-level actionable plan**")
    action_cols = ["Category","Normalized Product","SKU","Color","Size","Salable Stock","WI Stock","Current Physical","Return Stock","Move WI → Mall","Safe Send","Action"]
    action_grid = plan[action_cols].sort_values(["Action","Category","Normalized Product","Color","Size"])
    st.dataframe(action_grid, use_container_width=True, hide_index=True, height=520)

    safe_only = df[df["Safe Send"] > 0][action_cols]
    st.download_button("Download Safe Send List (CSV)", safe_only.to_csv(index=False).encode("utf-8-sig"), "Prelegent_Safe_Send_List.csv", "text/csv")

# -----------------------------
# 6. Raw SKU Data
# -----------------------------
with pages[5]:
    st.markdown('<div class="section-title">SKU Data</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">Searchable audit layer behind every dashboard number.</div>', unsafe_allow_html=True)
    q = st.text_input("Search SKU / product / colour / size", placeholder="e.g. RFPB, Black, XL, Polo")
    raw_view = df.copy()
    if q.strip():
        mask = pd.Series(False, index=raw_view.index)
        for c in ["SKU","Normalized Product","Category","Color","Size"]:
            mask = mask | raw_view[c].astype(str).str.contains(q, case=False, na=False, regex=False)
        raw_view = raw_view[mask]
    show_cols = ["Category","Normalized Product","Original Name","SKU","Color","Size","Price","Salable Stock","WI Stock","WI Out Stock","Current Physical","Return Stock","Return Block","Move WI → Mall","Safe Send","Action"]
    st.dataframe(raw_view[show_cols], use_container_width=True, hide_index=True, height=650)

st.caption("VIBES × Prelegent · Decision support portal. Send quantities are planning scenarios and should be validated with sell-through history when available.")

"""
Universal Finance AI — Streamlit App
Matches friend's output: Sunburst chart, Cash Flow Trend, Categorized Ledger
Supports: GoldmanSachs (CSV), JPMC (Excel), PNCBank/WellsFargo (PDF)
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import pdfplumber
import re
import glob
import os
from io import BytesIO

st.set_page_config(
    page_title="Universal Finance AI",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ── Minimal clean styling ─────────────────────────────────────────────────────
st.markdown("""
<style>
    .main { padding: 2rem; }
    .stApp { background: #ffffff; }
    h1 { font-size: 2rem !important; font-weight: 800 !important; }
    .metric-card {
        background: #f8f9fa;
        border-radius: 12px;
        padding: 20px;
        border: 1px solid #e9ecef;
    }
    .metric-value { font-size: 2rem; font-weight: 700; color: #1a1a2e; }
    .metric-label { font-size: 0.85rem; color: #6c757d; margin-bottom: 4px; }
    footer { visibility: hidden; }
    #MainMenu { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ── Category Rules ────────────────────────────────────────────────────────────
RULES = {
    "Groceries":    ["grocer","fruit","bread","vegetable","supermar","dairy","household essential","food shop","bought food","kirana","fresh produce","weekly shop"],
    "Bills":        ["utility","electric","water bill","gas bill","mortgage","rent","loan","insurance","subscription","cable","internet","phone bill","settled util","apartment","car loan"],
    "Food & Dining":["restaurant","cafe","coffee","snack","lunch","dinner","breakfast","fast food","pizza","burger","dining","eatery","takeout","meal"],
    "Transport":    ["fuel","petrol","diesel","transport card","cab","uber","ola","taxi","train ticket","bus","flight","airline","car fuel","commute","parking","toll"],
    "Shopping":     ["clothing","retail store","online shop","amazon","flipkart","mall","apparel","fashion","shoes","accessories","department store","purchased cloth"],
    "Healthcare":   ["medical","clinic","doctor","hospital","pharmacy","medicine","health","dental","consultation","check-up","lab test","prescription"],
    "Education":    ["textbook","university","college","course","tuition","school","training","workshop","seminar","online course","enrollment","books"],
    "Travel":       ["travel insurance","hotel","resort","vacation","holiday","airbnb","booking","trip","tour","amusement park","theme park","sightseeing"],
    "Entertainment":["movie","cinema","netflix","spotify","streaming","gaming","concert","event","show","ticket","museum","sport"],
    "Investments":  ["investment","stock","mutual fund","sip","demat","trading","equity","bond","fixed deposit","ppf","crypto","cryptocurrency","shares","freelance","received","income","salary","bonus","dividend"],
    "Income":       ["salary","salary credit","received","freelance","bonus","income","dividend","reimbursement","credit from","payment received","wages","stipend"],
    "ATM & Cash":   ["atm","cash withdrawal","cash deposit"],
    "Transfers":    ["transfer","neft","imps","upi","rtgs","sent","wire","remittance","credit card bill","installment","settled","topped up","public transport card"],
}

KW_INDEX = sorted(
    [(kw, cat) for cat, kws in RULES.items() for kw in kws],
    key=lambda x: len(x[0]), reverse=True
)

def categorize(desc: str) -> str:
    t = str(desc).lower()
    for kw, cat in KW_INDEX:
        if kw in t:
            return cat
    return "Others"

# ── Parsers ───────────────────────────────────────────────────────────────────

def parse_date(v):
    for fmt in ["%m/%d/%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y",
                "%m-%d-%Y", "%d-%b-%Y", "%d/%b/%Y", "%b-%d-%Y",
                "%m/%d/%y", "%d/%m/%y"]:
        try:
            return pd.to_datetime(str(v).strip(), format=fmt)
        except: pass
    try:
        return pd.to_datetime(str(v).strip(), dayfirst=False)
    except:
        return pd.NaT

def parse_amount(v):
    try:
        return abs(float(re.sub(r"[^0-9.\-]", "", str(v)) or 0))
    except:
        return 0.0

def load_goldman(path_or_bytes, name="file"):
    """AccID, Details, Bill $, Timestamp"""
    try:
        if isinstance(path_or_bytes, (str, os.PathLike)):
            df = pd.read_csv(path_or_bytes, dtype=str)
        else:
            df = pd.read_csv(path_or_bytes, dtype=str)
        df.columns = df.columns.str.strip()
        rows = []
        for _, r in df.iterrows():
            date = parse_date(r.get("Timestamp",""))
            desc = str(r.get("Details","")).strip()
            amt  = parse_amount(r.get("Bill $", 0))
            if pd.isna(date) or not desc or amt == 0: continue
            rows.append({"date": date, "description": desc, "amount": amt,
                         "category": categorize(desc), "source": name})
        return pd.DataFrame(rows)
    except Exception as e:
        st.warning(f"Could not load {name}: {e}")
        return pd.DataFrame()

def load_jpmc(path_or_bytes, name="file"):
    """Date, TranscID, FromAccount, To_Account/VAN, Amount, Transaction Description"""
    try:
        if isinstance(path_or_bytes, (str, os.PathLike)):
            df = pd.read_excel(path_or_bytes, dtype=str)
        else:
            df = pd.read_excel(path_or_bytes, dtype=str)
        df.columns = df.columns.str.strip()
        rows = []
        desc_col = next((c for c in df.columns if "description" in c.lower() or "detail" in c.lower()), None)
        date_col = next((c for c in df.columns if "date" in c.lower()), None)
        amt_col  = next((c for c in df.columns if "amount" in c.lower()), None)
        if not all([desc_col, date_col, amt_col]): return pd.DataFrame()
        for _, r in df.iterrows():
            date = parse_date(r[date_col])
            desc = str(r[desc_col]).strip()
            amt  = parse_amount(r[amt_col])
            if pd.isna(date) or not desc or amt == 0: continue
            rows.append({"date": date, "description": desc, "amount": amt,
                         "category": categorize(desc), "source": name})
        return pd.DataFrame(rows)
    except Exception as e:
        st.warning(f"Could not load {name}: {e}")
        return pd.DataFrame()

def load_pdf(path_or_bytes, name="file"):
    """Date, Amount ($), Transaction Description"""
    try:
        rows = []
        pdf_obj = pdfplumber.open(path_or_bytes) if not isinstance(path_or_bytes, (str, os.PathLike)) else pdfplumber.open(path_or_bytes)
        with pdf_obj as pdf:
            for page in pdf.pages:
                for table in (page.extract_tables() or []):
                    if len(table) < 2: continue
                    headers = [str(h).strip().lower() if h else "" for h in table[0]]
                    date_i = next((i for i,h in enumerate(headers) if "date" in h), None)
                    amt_i  = next((i for i,h in enumerate(headers) if "amount" in h), None)
                    desc_i = next((i for i,h in enumerate(headers) if "description" in h or "detail" in h), None)
                    if None in [date_i, amt_i, desc_i]: continue
                    for row in table[1:]:
                        if not row or len(row) <= max(date_i, amt_i, desc_i): continue
                        date = parse_date(row[date_i])
                        amt  = parse_amount(row[amt_i])
                        desc = str(row[desc_i]).strip()
                        if pd.isna(date) or not desc or amt == 0: continue
                        rows.append({"date": date, "description": desc, "amount": amt,
                                     "category": categorize(desc), "source": name})
        return pd.DataFrame(rows)
    except Exception as e:
        st.warning(f"Could not load PDF {name}: {e}")
        return pd.DataFrame()

def load_file(path_or_bytes, name):
    n = name.lower()
    if n.endswith(".csv"):
        return load_goldman(path_or_bytes, name)
    elif n.endswith((".xlsx", ".xls")):
        return load_jpmc(path_or_bytes, name)
    elif n.endswith(".pdf"):
        return load_pdf(path_or_bytes, name)
    return pd.DataFrame()

def load_all_datasets():
    ds = os.path.join(os.path.dirname(__file__), "datasets")
    frames = []
    files  = sorted(glob.glob(os.path.join(ds, "*")))
    for f in files:
        name = os.path.basename(f)
        df   = load_file(f, name)
        if not df.empty:
            frames.append(df)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

# ── KPI Cards ─────────────────────────────────────────────────────────────────
def show_kpis(df):
    deb = df[~df["category"].isin(["Income","Investments"])]
    total_exp  = deb["amount"].sum()
    total_txns = len(df)
    top_cat    = deb["category"].value_counts().idxmax() if len(deb) > 0 else "N/A"
    avg_txn    = deb["amount"].mean() if len(deb) > 0 else 0

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.metric("Total Expenses", f"₹{total_exp:,.2f}")
    with c2:
        st.metric("Total Transactions", f"{total_txns:,}")
    with c3:
        st.metric("Top Spend Category", top_cat)
    with c4:
        st.metric("Avg Transaction", f"₹{avg_txn:,.2f}")

# ── Sunburst / Spending Hierarchy ─────────────────────────────────────────────
def show_sunburst(df):
    cat_totals = df.groupby(["category","description"])["amount"].sum().reset_index()
    # Limit descriptions to top 5 per category to avoid clutter
       top = top[top["amount"] > 0].copy()
           .apply(lambda x: x.nlargest(5, "amount"))
           .reset_index(drop=True))

    fig = px.sunburst(
        top,
        path=["category", "description"],
        values="amount",
        color="category",
       color_discrete_sequence=px.colors.qualitative.Set3,
    )
    fig.update_traces(textinfo="label", insidetextorientation="radial")
    fig.update_layout(
        margin=dict(t=10, b=10, l=10, r=10),
        height=480,
    )
    return fig

# ── Cash Flow Trend ───────────────────────────────────────────────────────────
def show_cashflow(df):
    df2 = df.copy()
    df2["date_dt"] = pd.to_datetime(df2["date"])
    df2 = df2.sort_values("date_dt")

    fig = px.bar(
        df2, x="date_dt", y="amount",
        color_discrete_sequence=["#4472C4"],
        labels={"date_dt": "date_dt", "amount": "amount"},
    )
    fig.update_layout(
        height=380,
        xaxis_title="date_dt",
        yaxis_title="amount",
        showlegend=False,
        plot_bgcolor="white",
        paper_bgcolor="white",
        xaxis=dict(gridcolor="#f0f0f0"),
        yaxis=dict(gridcolor="#f0f0f0"),
        margin=dict(t=20, b=40),
    )
    fig.update_traces(marker_line_width=0)
    return fig

# ── Main App ──────────────────────────────────────────────────────────────────

def main():
    # Header
    st.markdown("## 🏦 Universal Finance AI")

    # Upload section
    st.markdown("📁 **Upload Bank Statement**")
    uploaded = st.file_uploader(
        "",
        type=["csv", "xlsx", "xls", "pdf"],
        accept_multiple_files=True,
        label_visibility="collapsed"
    )

    # Load data
    if uploaded:
        frames = []
        for f in uploaded:
            df = load_file(f, f.name)
            if not df.empty:
                frames.append(df)
        if frames:
            data = pd.concat(frames, ignore_index=True)
        else:
            st.error("Could not parse uploaded files.")
            return
    else:
        # Auto-load all datasets
        with st.spinner("Loading all datasets..."):
            data = load_all_datasets()
        if data.empty:
            st.warning("No datasets found. Upload a file above.")
            return

    # Ensure date column
    data["date"] = pd.to_datetime(data["date"], errors="coerce")
    data = data.dropna(subset=["date"])
    data = data.sort_values("date").reset_index(drop=True)

    # ── KPIs
    show_kpis(data)

    st.markdown("---")

    # ── Charts row
    col1, col2 = st.columns([1, 1.5])

    with col1:
        st.markdown("### Spending Hierarchy")
        fig_sun = show_sunburst(data)
        st.plotly_chart(fig_sun, use_container_width=True)

    with col2:
        st.markdown("### Cash Flow Trend")
        fig_cf = show_cashflow(data)
        st.plotly_chart(fig_cf, use_container_width=True)

    st.markdown("---")

    # ── Categorized Ledger
    st.markdown("### Final Categorized Ledger")

    # Filters
    f1, f2, f3 = st.columns([2, 1, 1])
    with f1:
        search = st.text_input("🔍 Search description", "")
    with f2:
        cats = ["All"] + sorted(data["category"].unique().tolist())
        sel_cat = st.selectbox("Category", cats)
    with f3:
        sources = ["All"] + sorted(data["source"].unique().tolist())
        sel_src = st.selectbox("Source", sources)

    filtered = data.copy()
    if search:
        filtered = filtered[filtered["description"].str.contains(search, case=False, na=False)]
    if sel_cat != "All":
        filtered = filtered[filtered["category"] == sel_cat]
    if sel_src != "All":
        filtered = filtered[filtered["source"] == sel_src]

    # Display table
    display = filtered[["date", "description", "amount", "category", "source"]].copy()
    display["date"] = display["date"].dt.strftime("%Y-%m-%d")
    display["amount"] = display["amount"].round(2)
    display = display.reset_index(drop=True)

    st.dataframe(
        display,
        use_container_width=True,
        height=420,
        column_config={
            "date":        st.column_config.TextColumn("date"),
            "description": st.column_config.TextColumn("description", width="large"),
            "amount":      st.column_config.NumberColumn("amount", format="%.2f"),
            "category":    st.column_config.TextColumn("category"),
            "source":      st.column_config.TextColumn("source"),
        }
    )

    st.caption(f"Showing {len(display):,} of {len(data):,} transactions")

    # ── Download
    csv = display.to_csv(index=False).encode()
    st.download_button(
        "⬇ Download Categorized CSV",
        csv,
        "categorized_transactions.csv",
        "text/csv"
    )

if __name__ == "__main__":
    main()

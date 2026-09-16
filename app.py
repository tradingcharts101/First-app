import streamlit as st
import pandas as pd
import requests

st.set_page_config(page_title="TVS Market Navigator", layout="wide")

st.title("📊 TVS Market Navigator - Full Watchlist & Reserve Engine")
st.caption("Automated allocation engine for Core 30, Queues A/B/C, and Reserve List.")

@st.cache_data
def load_database():
    df = pd.read_excel(
        "TVS_Market_Classification_Database.xlsx",
        sheet_name="Master Index"
    )
    df.columns = [c.strip() for c in df.columns]
    return df

try:
    db_df = load_database()
    st.sidebar.success("Database Loaded Successfully!")
except Exception as e:
    st.error(f"Error loading classification database: {e}")
    st.stop()

st.sidebar.header("Weekly RRG Sector Analysis")

default_leading = (
    "REALTY, CONSUMER DURABLES, MEDIA, ENTERTAINMENT & PUBLICATION, "
    "CONSUMER SERVICES, AUTOMOBILES & AUTO COMPS, FOREST MATERIALS, "
    "CONSTRUCTION MATERIALS, SERVICES"
)
default_improving = "INFORMATION TECHNOLOGY"
default_weakening = (
    "TEXTILES, TELECOMMUNICATION, HEALTHCARE, CHEMICALS, DIVERSIFIED, "
    "CAPITAL GOODS, FINANCIAL SERVICES, METALS & MINING"
)
default_lagging = (
    "CONSTRUCTION, OIL, GAS & CONSUMABLE FUELS, FAST MOVING CONSUMER GOODS, "
    "UTILITIES, POWER"
)

leading_input = st.sidebar.text_area("LEADING", default_leading, height=70)
improving_input = st.sidebar.text_area("IMPROVING", default_improving, height=50)
weakening_input = st.sidebar.text_area("WEAKENING", default_weakening, height=70)
lagging_input = st.sidebar.text_area("LAGGING", default_lagging, height=70)

def parse_sectors(text):
    return [s.strip().upper() for s in text.split(",") if s.strip()]

def get_quadrant(sector_name, quadrants):
    sec_upper = str(sector_name).upper()
    for quad, sec_list in quadrants.items():
        if any(s in sec_upper or sec_upper in s for s in sec_list):
            return quad
    return "UNKNOWN"

if st.button("🚀 Run Full TVS Analysis & Generate Watchlists"):
    quadrants = {
        "Leading": parse_sectors(leading_input),
        "Improving": parse_sectors(improving_input),
        "Weakening": parse_sectors(weakening_input),
        "Lagging": parse_sectors(lagging_input)
    }

    df = db_df.copy()
    df["RRG_Quadrant"] = df["Sector"].apply(
        lambda x: get_quadrant(x, quadrants)
    )
    df = df.sort_values(
        by=["Industry", "Market Capitalisation"],
        ascending=[True, False]
    )

    # Core 30 Selection
    core_candidates = df.groupby("Industry").first().reset_index()

    trend_picks = core_candidates[
        core_candidates["RRG_Quadrant"].isin(["Leading", "Improving"])
    ].head(15)

    turn_picks = core_candidates[
        core_candidates["RRG_Quadrant"].isin(["Weakening", "Lagging"])
    ].head(15)

    used_symbols = set(trend_picks["Stock Symbol"]).union(
        set(turn_picks["Stock Symbol"])
    )
    remaining_df = df[~df["Stock Symbol"].isin(used_symbols)]

    # Queue A - Sector Breadth
    queue_a = remaining_df[
        remaining_df["RRG_Quadrant"].isin(["Leading", "Improving"])
    ].head(15)
    used_symbols.update(queue_a["Stock Symbol"])
    remaining_df = df[~df["Stock Symbol"].isin(used_symbols)]

    # Queue B - Cap Contraction
    queue_b = remaining_df[
        remaining_df["RRG_Quadrant"].isin(["Weakening", "Lagging"])
    ].head(15)
    used_symbols.update(queue_b["Stock Symbol"])
    remaining_df = df[~df["Stock Symbol"].isin(used_symbols)]

    # Queue C - Special / Broad Watch
    queue_c = remaining_df.head(15)
    used_symbols.update(queue_c["Stock Symbol"])
    remaining_df = df[~df["Stock Symbol"].isin(used_symbols)]

    # Reserve List
    reserve_list = remaining_df.head(20)

    st.markdown("---")
    st.header("📋 TradingView Copy-Paste Watchlists (75 Tickers + Reserve)")

    st.subheader("1. Core 30 Watchlist (30 Stocks)")
    c1, c2 = st.columns(2)

    with c1:
        st.write("**TREND (15)**")
        t_str = ", ".join(
            [f"NSE:{s}" for s in trend_picks["Stock Symbol"]]
        ) + ","
        st.text_area("Paste String - TREND", t_str, height=90)
        st.dataframe(
            trend_picks[
                ["Stock Symbol", "Industry", "Market Capitalisation"]
            ]
        )

    with c2:
        st.write("**TURN (15)**")
        tu_str = ", ".join(
            [f"NSE:{s}" for s in turn_picks["Stock Symbol"]]
        ) + ","
        st.text_area("Paste String - TURN", tu_str, height=90)
        st.dataframe(
            turn_picks[
                ["Stock Symbol", "Industry", "Market Capitalisation"]
            ]
        )

    st.markdown("---")
    st.subheader("2. Swap Queues (45 Stocks)")
    q1, q2, q3 = st.columns(3)

    with q1:
        st.write("**Queue A - Breadth (15)**")
        qa_str = ", ".join(
            [f"NSE:{s}" for s in queue_a["Stock Symbol"]]
        ) + ","
        st.text_area("Paste String - Queue A", qa_str, height=90)
        st.dataframe(
            queue_a[["Stock Symbol", "Industry", "RRG_Quadrant"]]
        )

    with q2:
        st.write("**Queue B - Cap Contraction (15)**")
        qb_str = ", ".join(
            [f"NSE:{s}" for s in queue_b["Stock Symbol"]]
        ) + ","
        st.text_area("Paste String - Queue B", qb_str, height=90)
        st.dataframe(
            queue_b[["Stock Symbol", "Industry", "RRG_Quadrant"]]
        )

    with q3:
        st.write("**Queue C - Watchlist (15)**")
        qc_str = ", ".join(
            [f"NSE:{s}" for s in queue_c["Stock Symbol"]]
        ) + ","
        st.text_area("Paste String - Queue C", qc_str, height=90)
        st.dataframe(
            queue_c[["Stock Symbol", "Industry", "RRG_Quadrant"]]
        )

    st.markdown("---")
    st.subheader("3. TVS Reserve List")

    res_str = ", ".join(
        [f"NSE:{s}" for s in reserve_list["Stock Symbol"]]
    ) + ","
    st.text_area("Paste String - TVS Reserve List", res_str, height=90)
    st.dataframe(
        reserve_list[
            [
                "Stock Symbol",
                "Company Name",
                "Sector",
                "Industry",
                "Market Capitalisation"
            ]
        ]
    )

    st.success(
        "Complete Watchlist (75 Stocks + Reserve List) Generated Successfully!"
    )

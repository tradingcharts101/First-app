import streamlit as st
import pandas as pd
import requests
import re

st.set_page_config(page_title="TVS Market Navigator", layout="wide")

st.title("📊 TVS Market Navigator - One-Click Screener Engine")
st.caption("Automated allocation engine built on TVS Market Classification Framework.")

# ---------------------------------------------------------
# 1. LOAD CLASSIFICATION DATABASE
# ---------------------------------------------------------
@st.cache_data
def load_database():
    df = pd.read_excel(
        "TVS_Market_Classification_Database.xlsx",
        sheet_name="Master Index"
    )
    # Standardize column names
    df.columns = [c.strip() for c in df.columns]
    return df


try:
    db_df = load_database()
    st.sidebar.success("Database Loaded Successfully!")
except Exception as e:
    st.error(f"Error loading classification database: {e}")
    st.stop()

# ---------------------------------------------------------
# 2. RRG SECTOR QUADRANTS INPUT (SIDEBAR)
# ---------------------------------------------------------
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

leading_input = st.sidebar.text_area(
    "LEADING", default_leading, height=80
)
improving_input = st.sidebar.text_area(
    "IMPROVING", default_improving, height=60
)
weakening_input = st.sidebar.text_area(
    "WEAKENING", default_weakening, height=80
)
lagging_input = st.sidebar.text_area(
    "LAGGING", default_lagging, height=80
)

# ---------------------------------------------------------
# 3. HELPER FUNCTIONS
# ---------------------------------------------------------
def parse_sectors(text):
    return [s.strip().upper() for s in text.split(",") if s.strip()]


def get_quadrant(sector_name, quadrants):
    sec_upper = str(sector_name).upper()

    for quad, sec_list in quadrants.items():
        if any(s in sec_upper or sec_upper in s for s in sec_list):
            return quad

    return "UNKNOWN"


def fetch_chartink_scan(scan_url):
    """Scrapes Chartink scan results dynamically."""
    try:
        session = requests.Session()
        headers = {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36"
            )
        }

        res = session.get(scan_url, headers=headers)

        # Chartink backend processing endpoint call simulation
        # In actual usage, supply your custom scan parameters/cookies
        return []

    except Exception:
        return []


# ---------------------------------------------------------
# 4. MAIN SCREENER RUNNER
# ---------------------------------------------------------
if st.button("🚀 Run One-Click Screen & Generate Watchlist"):

    quadrants = {
        "Leading": parse_sectors(leading_input),
        "Improving": parse_sectors(improving_input),
        "Weakening": parse_sectors(weakening_input),
        "Lagging": parse_sectors(lagging_input),
    }

    st.subheader("1. System Setup & Data Processing")

    # -----------------------------------------------------
    # STEP 1: MATCHING & QUADRANT TAGGING
    # -----------------------------------------------------
    db_df["RRG_Quadrant"] = db_df["Sector"].apply(
        lambda x: get_quadrant(x, quadrants)
    )

    # Example filtered candidate selection from Database
    candidates = db_df.copy()

    st.write(
        f"Total Database Coverage: **{len(db_df)}** stocks "
        f"across **{db_df['Industry'].nunique()}** industries."
    )

    # -----------------------------------------------------
    # STEP 2: CORE 30 ALLOCATION (TREND vs TURN)
    # -----------------------------------------------------
    st.subheader("2. Core 30 Allocation")

    # Sort candidates by Market Cap within each Industry
    candidates = candidates.sort_values(
        by=["Industry", "Market Capitalisation"],
        ascending=[True, False],
    )

    # Assign 1 Top Large/Mid-Cap pick per Industry
    industry_picks = candidates.groupby("Industry").first().reset_index()

    # Split into 15 TREND and 15 TURN slots based on Quadrant Priority
    trend_picks = industry_picks[
        industry_picks["RRG_Quadrant"].isin(["Leading", "Improving"])
    ].head(15)

    turn_picks = industry_picks[
        industry_picks["RRG_Quadrant"].isin(["Weakening", "Lagging"])
    ].head(15)

    # -----------------------------------------------------
    # STEP 3: TRADINGVIEW FORMATTED OUTPUTS
    # -----------------------------------------------------
    st.markdown("---")
    st.header("📋 TradingView Copy-Paste Strings")

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("TREND (15)")

        trend_tickers = [
            f"NSE:{t}" for t in trend_picks["Stock Symbol"].tolist()
        ]
        trend_paste_str = ", ".join(trend_tickers) + ","

        st.text_area(
            "Paste String (TREND)",
            trend_paste_str,
            height=120
        )

        st.dataframe(
            trend_picks[
                [
                    "Stock Symbol",
                    "Industry",
                    "Sector",
                    "RRG_Quadrant",
                    "Market Capitalisation",
                ]
            ]
        )

    with col2:
        st.subheader("TURN (15)")

        turn_tickers = [
            f"NSE:{t}" for t in turn_picks["Stock Symbol"].tolist()
        ]
        turn_paste_str = ", ".join(turn_tickers) + ","

        st.text_area(
            "Paste String (TURN)",
            turn_paste_str,
            height=120
        )

        st.dataframe(
            turn_picks[
                [
                    "Stock Symbol",
                    "Industry",
                    "Sector",
                    "RRG_Quadrant",
                    "Market Capitalisation",
                ]
            ]
        )

    st.markdown("---")
    st.caption(
        "This is a study watchlist for chart reading practice, "
        "not a recommendation to buy or sell."
    )

"""
TVS Market Navigator — Weekly Watchlist Organizer
===================================================

Implements the "Watchlist Organizer" pipeline exactly as specified in the
project's Claude Project Instructions:

  1.  Look up Sector / Industry / Market-Cap Band for every ChartInk hit
      from the Master Classification Database.
  2.  Tag TREND (TVS1+TVS3) vs TURN (TVS2+TVS4) family, and RRG quadrant
      (via each stock's Sector).
  3.  Participation rate per industry = hits / industry universe size.
  3A. Quadrant capacity check + Core quota allocation (12/10/6/2, capped
      and reallocated in priority Leading > Improving > Weakening > Lagging).
  4.  Core 30 = 15 TREND + 15 TURN, one slot per industry, built by walking
      the ranked (by participation rate) candidate list once.
  4B. Queue-capacity recount for unclaimed industries.
  5.  Queue A — Breadth (15): best unclaimed industries, quadrant quotas 6/5/3/1.
  6.  Queue B — Cap Contrast (15): second stock (different cap band) from the
      15 Core industries with the most qualifying members.
  7.  Queue C — Structure Contrast (15): stock from the *other* family for
      Core industries that qualified in both families.
  8.  De-duplication across Core > A > B > C.

  Reporting: Sector Breadth Table, Industry Structure Profile (>=8 hits),
  Industry Participation Ledger, Reserve List (max 8/industry).

This app is CLERICAL ONLY — it never ranks by chart quality, never predicts
prices, and carries no live market data. Every number comes from the two
files you upload; every classification comes from the Master Database.

Run with:  streamlit run app.py
"""

from __future__ import annotations

import io
from dataclasses import dataclass, field
from datetime import date
from typing import Optional

import pandas as pd
import streamlit as st

# --------------------------------------------------------------------------
# Constants
# --------------------------------------------------------------------------

TVS_SHEETS = ["TVS1", "TVS2", "TVS3", "TVS4"]
SCAN_NAMES = {
    "TVS1": "Momentum Continuation",
    "TVS2": "Base Recovery",
    "TVS3": "Trend Pullback",
    "TVS4": "Reversals",
}
# TVS1 > TVS2 > TVS3 > TVS4 priority when the same symbol hits multiple scans.
SCAN_PRIORITY = {"TVS1": 1, "TVS2": 2, "TVS3": 3, "TVS4": 4}
TREND_TAGS = {"TVS1", "TVS3"}
TURN_TAGS = {"TVS2", "TVS4"}

QUADRANTS = ["Leading", "Improving", "Weakening", "Lagging"]
CORE_TARGET_SHARE = {"Leading": 12, "Improving": 10, "Weakening": 6, "Lagging": 2}  # of 30
QUEUE_A_TARGET_SHARE = {"Leading": 6, "Improving": 5, "Weakening": 3, "Lagging": 1}  # of 15
QUADRANT_PRIORITY = ["Leading", "Improving", "Weakening", "Lagging"]

CAP_BANDS = ["Large", "Mid", "Small"]  # priority order used when picking a representative stock

UNLISTED = "UNLISTED"


# --------------------------------------------------------------------------
# Small data holder for everything the pipeline produces, so the UI layer
# can just render `result.<thing>` without recomputing anything.
# --------------------------------------------------------------------------

@dataclass
class PipelineResult:
    hits: pd.DataFrame = field(default_factory=pd.DataFrame)
    unlisted: pd.DataFrame = field(default_factory=pd.DataFrame)
    participation: pd.DataFrame = field(default_factory=pd.DataFrame)
    sector_breadth: pd.DataFrame = field(default_factory=pd.DataFrame)
    industry_structure: pd.DataFrame = field(default_factory=pd.DataFrame)
    ledger: pd.DataFrame = field(default_factory=pd.DataFrame)
    core: pd.DataFrame = field(default_factory=pd.DataFrame)
    queue_a: pd.DataFrame = field(default_factory=pd.DataFrame)
    queue_b: pd.DataFrame = field(default_factory=pd.DataFrame)
    queue_c: pd.DataFrame = field(default_factory=pd.DataFrame)
    reserve: pd.DataFrame = field(default_factory=pd.DataFrame)
    notes: list[str] = field(default_factory=list)  # every reallocation / shortfall message


# --------------------------------------------------------------------------
# STEP 0 — Loading & normalising the two input files
# --------------------------------------------------------------------------

def _find_col(columns: list[str], candidates: list[str]) -> Optional[str]:
    """Case/space-insensitive column matcher."""
    norm = {c.strip().lower().replace(" ", ""): c for c in columns}
    for cand in candidates:
        key = cand.strip().lower().replace(" ", "")
        if key in norm:
            return norm[key]
    return None


def load_master_database(file) -> pd.DataFrame:
    """Reads the Master Classification Database and assigns cap bands.

    Cap bands are re-ranked EVERY run, on this file's Market Cap column:
    rank 1-100 = Large, 101-250 = Mid, 251+ = Small — matching the pipeline
    rule ("assigned dynamically each run by ranking all companies in the
    master database").
    """
    if hasattr(file, "seek"):
        file.seek(0)
    df = pd.read_excel(file, sheet_name="Master Index") if _has_sheet(file, "Master Index") \
        else pd.read_excel(file)

    sym_col = _find_col(df.columns, ["Symbol", "Ticker", "NSE Symbol", "NSE Code", "Stock Symbol"])
    sector_col = _find_col(df.columns, ["Sector"])
    industry_col = _find_col(df.columns, ["Industry"])
    cap_col = _find_col(df.columns, ["Market Cap", "MarketCap", "Mcap", "Market Cap (Cr)",
                                      "Market Capitalisation", "Market Capitalization"])

    missing = [n for n, c in [("Symbol", sym_col), ("Sector", sector_col),
                               ("Industry", industry_col), ("Market Cap", cap_col)] if c is None]
    if missing:
        raise ValueError(
            f"Master database is missing column(s): {', '.join(missing)}. "
            f"Found columns: {list(df.columns)}"
        )

    out = df[[sym_col, sector_col, industry_col, cap_col]].copy()
    out.columns = ["Symbol", "Sector", "Industry", "MarketCap"]
    out["Symbol"] = out["Symbol"].astype(str).str.strip().str.upper()
    out["Sector"] = out["Sector"].astype(str).str.strip()
    out["Industry"] = out["Industry"].astype(str).str.strip()
    out["MarketCap"] = pd.to_numeric(out["MarketCap"], errors="coerce")
    out = out.dropna(subset=["Symbol", "MarketCap"]).drop_duplicates(subset="Symbol")

    out = out.sort_values("MarketCap", ascending=False).reset_index(drop=True)
    rank = out.index + 1
    out["CapBand"] = pd.cut(
        rank, bins=[0, 100, 250, len(out) + 1], labels=["Large", "Mid", "Small"]
    ).astype(str)
    return out


def _has_sheet(file, name: str) -> bool:
    try:
        if hasattr(file, "seek"):
            file.seek(0)
        xl = pd.ExcelFile(file)
        if hasattr(file, "seek"):
            file.seek(0)
        return name in xl.sheet_names
    except Exception:
        return False


def load_screener(file) -> dict[str, pd.DataFrame]:
    """Reads the four TVS1-TVS4 sheets, keeping only Symbol / Close / Volume."""
    if hasattr(file, "seek"):
        file.seek(0)
    xl = pd.ExcelFile(file)
    sheets: dict[str, pd.DataFrame] = {}
    for tag in TVS_SHEETS:
        actual = next((s for s in xl.sheet_names if s.strip().upper() == tag), None)
        if actual is None:
            raise ValueError(f"Screener workbook is missing a '{tag}' sheet. "
                              f"Found sheets: {xl.sheet_names}")
        df = xl.parse(actual)
        sym_col = _find_col(df.columns, ["Symbol", "Ticker"])
        close_col = _find_col(df.columns, ["Close", "Close Price", "LTP"])
        vol_col = _find_col(df.columns, ["Volume", "Volume 1 week ago", "Volume ("])
        if sym_col is None:
            raise ValueError(f"Sheet '{tag}' has no Symbol column. Found: {list(df.columns)}")
        keep = df[[sym_col] + [c for c in [close_col, vol_col] if c]].copy()
        keep.columns = ["Symbol"] + (["Close"] if close_col else []) + (["Volume"] if vol_col else [])
        keep["Symbol"] = keep["Symbol"].astype(str).str.strip().str.upper()
        keep = keep.dropna(subset=["Symbol"]).drop_duplicates(subset="Symbol")
        sheets[tag] = keep
    return sheets


# --------------------------------------------------------------------------
# STEP 1 & 2 — Lookup + dedupe + family/quadrant tagging
# --------------------------------------------------------------------------

def build_hits_table(
    screener: dict[str, pd.DataFrame], master: pd.DataFrame, sector_quadrant: dict[str, str]
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """One row per unique symbol across all four scans, tagged with the FIRST
    sheet it appeared in (TVS1 > TVS2 > TVS3 > TVS4), then enriched with
    Sector / Industry / CapBand and Family / Quadrant.
    Returns (hits_df, unlisted_df).
    """
    rows = []
    for tag, df in screener.items():
        for _, r in df.iterrows():
            rows.append({"Symbol": r["Symbol"], "ScanTag": tag, "Priority": SCAN_PRIORITY[tag]})
    long_df = pd.DataFrame(rows)
    if long_df.empty:
        return pd.DataFrame(), pd.DataFrame()

    long_df = long_df.sort_values("Priority").drop_duplicates(subset="Symbol", keep="first")

    merged = long_df.merge(master, on="Symbol", how="left")
    unlisted = merged[merged["Sector"].isna()][["Symbol", "ScanTag"]].reset_index(drop=True)

    hits = merged.dropna(subset=["Sector"]).copy()
    hits["Family"] = hits["ScanTag"].map(lambda t: "TREND" if t in TREND_TAGS else "TURN")
    hits["Quadrant"] = hits["Sector"].map(sector_quadrant).fillna("Unmapped")
    hits = hits.reset_index(drop=True)
    return hits, unlisted


# --------------------------------------------------------------------------
# STEP 3 — Participation rate & Sector breadth
# --------------------------------------------------------------------------

def compute_participation(hits: pd.DataFrame, master: pd.DataFrame) -> pd.DataFrame:
    universe = master.groupby("Industry")["Symbol"].nunique().rename("Universe")
    industry_sector = master.groupby("Industry")["Sector"].agg(lambda s: s.mode().iat[0])

    if hits.empty:
        return pd.DataFrame(columns=["Industry", "Sector", "Quadrant", "Hits", "Universe", "Rate"])

    hit_counts = hits.groupby("Industry")["Symbol"].nunique().rename("Hits")
    trend_counts = hits[hits.Family == "TREND"].groupby("Industry")["Symbol"].nunique().rename("TrendHits")
    turn_counts = hits[hits.Family == "TURN"].groupby("Industry")["Symbol"].nunique().rename("TurnHits")

    part = pd.concat([hit_counts, trend_counts, turn_counts], axis=1).fillna(0)
    part["TrendHits"] = part["TrendHits"].astype(int)
    part["TurnHits"] = part["TurnHits"].astype(int)
    part = part.join(universe, how="left")
    part["Rate"] = (part["Hits"] / part["Universe"] * 100).round(2)
    part["Sector"] = part.index.map(industry_sector)
    part["Quadrant"] = part.index.map(
        hits.drop_duplicates("Industry").set_index("Industry")["Quadrant"]
    )
    part = part.reset_index().rename(columns={"index": "Industry"})
    return part.sort_values("Rate", ascending=False).reset_index(drop=True)


def compute_sector_breadth(hits: pd.DataFrame, master: pd.DataFrame) -> pd.DataFrame:
    if hits.empty:
        return pd.DataFrame(columns=["Sector", "RRG Quadrant", "Qualifying", "Universe", "Breadth %"])
    universe = master.groupby("Sector")["Symbol"].nunique().rename("Universe")
    qualifying = hits.groupby("Sector")["Symbol"].nunique().rename("Qualifying")
    quadrant = hits.drop_duplicates("Sector").set_index("Sector")["Quadrant"]
    out = pd.concat([qualifying, universe], axis=1).dropna(subset=["Qualifying"])
    out["RRG Quadrant"] = out.index.map(quadrant)
    out["Breadth %"] = (out["Qualifying"] / out["Universe"] * 100).round(2)
    out = out.reset_index().rename(columns={"index": "Sector"})
    return out[["Sector", "RRG Quadrant", "Qualifying", "Universe", "Breadth %"]] \
        .sort_values("Breadth %", ascending=False).reset_index(drop=True)


def compute_industry_structure(hits: pd.DataFrame, participation: pd.DataFrame) -> pd.DataFrame:
    if hits.empty:
        return pd.DataFrame(columns=["Industry", "Hits", "Universe", "Rate %",
                                      "Near 52W High %", "Below 200 EMA %"])
    tag_counts = hits.groupby(["Industry", "ScanTag"])["Symbol"].nunique().unstack(fill_value=0)
    for t in TVS_SHEETS:
        if t not in tag_counts.columns:
            tag_counts[t] = 0
    rows = []
    for industry, hits_row in participation.set_index("Industry").iterrows():
        if hits_row["Hits"] < 8:
            continue
        t = tag_counts.loc[industry] if industry in tag_counts.index else pd.Series(0, index=TVS_SHEETS)
        near_high = round((t["TVS1"] + t["TVS2"]) / hits_row["Hits"] * 100, 2)
        below_ema = round(t["TVS4"] / hits_row["Hits"] * 100, 2)
        rows.append({
            "Industry": industry, "Hits": int(hits_row["Hits"]), "Universe": int(hits_row["Universe"]),
            "Rate %": hits_row["Rate"], "Near 52W High %": near_high, "Below 200 EMA %": below_ema,
        })
    out = pd.DataFrame(rows)
    return out.sort_values("Rate %", ascending=False).reset_index(drop=True) if not out.empty else out


# --------------------------------------------------------------------------
# STEP 3A / 4B — Capacity-aware quota allocation (shared by Core & Queue A)
# --------------------------------------------------------------------------

def allocate_quotas(available_by_quadrant: dict[str, int], target: dict[str, int], notes: list[str],
                     label: str) -> dict[str, int]:
    """Caps each quadrant's target at what's actually available, then pushes
    any deficit down the priority order Leading > Improving > Weakening > Lagging,
    respecting each quadrant's own cap. Records every reallocation in `notes`.
    """
    quota = {q: min(target[q], available_by_quadrant.get(q, 0)) for q in QUADRANT_PRIORITY}
    deficit = sum(target.values()) - sum(quota.values())
    for q in QUADRANT_PRIORITY:
        if deficit <= 0:
            break
        headroom = available_by_quadrant.get(q, 0) - quota[q]
        if headroom <= 0:
            continue
        add = min(headroom, deficit)
        quota[q] += add
        deficit -= add
        short_from = [f"{k} short {target[k] - min(target[k], available_by_quadrant.get(k, 0))}"
                      for k in QUADRANT_PRIORITY if target[k] > available_by_quadrant.get(k, 0)]
        notes.append(f"{label}: moved {add} slot(s) to {q} ({'; '.join(short_from) if short_from else 'rebalance'}).")
    if deficit > 0:
        notes.append(f"{label}: {deficit} slot(s) could not be filled — not enough qualifying industries in total.")
    return quota


# --------------------------------------------------------------------------
# STEP 4 — Core 30
# --------------------------------------------------------------------------

def pick_representative(hits: pd.DataFrame, industry: str, exclude_symbols: set[str] = frozenset(),
                         exclude_band: Optional[str] = None, family: Optional[str] = None) -> Optional[pd.Series]:
    """Largest market-cap stock in `industry`, preferring Large > Mid > Small band."""
    pool = hits[(hits.Industry == industry) & (~hits.Symbol.isin(exclude_symbols))]
    if family is not None:
        pool = pool[pool.Family == family]
    if exclude_band is not None:
        pool = pool[pool.CapBand != exclude_band]
    if pool.empty:
        return None
    for band in CAP_BANDS:
        band_pool = pool[pool.CapBand == band]
        if not band_pool.empty:
            return band_pool.sort_values("MarketCap", ascending=False).iloc[0]
    return pool.sort_values("MarketCap", ascending=False).iloc[0]


def build_core30(participation: pd.DataFrame, hits: pd.DataFrame, notes: list[str]) -> pd.DataFrame:
    if participation.empty:
        notes.append("Core 30: no qualifying industries — nothing to allocate.")
        return pd.DataFrame(columns=["Ticker", "Industry", "Sector", "Quadrant", "Scan Tag", "Cap Band", "Section"])

    available = participation.groupby("Quadrant")["Industry"].nunique().to_dict()
    quota = allocate_quotas(available, CORE_TARGET_SHARE, notes, "Core 30")

    candidates = []
    for q in QUADRANT_PRIORITY:
        sub = participation[participation.Quadrant == q].sort_values("Rate", ascending=False)
        candidates.append(sub.head(quota[q]))
    candidate_df = pd.concat(candidates).sort_values("Rate", ascending=False).reset_index(drop=True)

    trend_list, turn_list, used_symbols = [], [], set()
    for _, row in candidate_df.iterrows():
        industry = row["Industry"]
        trend_first = row["TrendHits"] >= row["TurnHits"]
        prefer_trend = trend_first if row["TrendHits"] != row["TurnHits"] else len(trend_list) <= len(turn_list)

        if prefer_trend and len(trend_list) < 15:
            section, family = "TREND", None
        elif (not prefer_trend) and len(turn_list) < 15:
            section, family = "TURN", None
        elif len(trend_list) < 15:
            section, family = "TREND", None
        elif len(turn_list) < 15:
            section, family = "TURN", None
        else:
            continue  # both sections already full

        rep = pick_representative(hits, industry, exclude_symbols=used_symbols)
        if rep is None:
            continue
        used_symbols.add(rep["Symbol"])
        entry = {
            "Ticker": rep["Symbol"], "Industry": industry, "Sector": rep["Sector"],
            "Quadrant": row["Quadrant"], "Scan Tag": rep["ScanTag"], "Cap Band": rep["CapBand"],
            "Section": section,
        }
        (trend_list if section == "TREND" else turn_list).append(entry)

    if len(trend_list) < 15 or len(turn_list) < 15:
        notes.append(f"Core 30: only reached TREND={len(trend_list)}/15, TURN={len(turn_list)}/15 "
                      f"— fewer than 30 qualifying industries were available.")

    return pd.DataFrame(trend_list + turn_list)


# --------------------------------------------------------------------------
# STEP 5 — Queue A (Breadth)
# --------------------------------------------------------------------------

def build_queue_a(participation: pd.DataFrame, hits: pd.DataFrame, core: pd.DataFrame,
                   notes: list[str]) -> pd.DataFrame:
    claimed = set(core["Industry"]) if not core.empty else set()
    unclaimed = participation[~participation.Industry.isin(claimed)].sort_values("Rate", ascending=False)
    if unclaimed.empty:
        return pd.DataFrame(columns=["Ticker", "Industry", "Sector", "Quadrant", "Scan Tag", "Cap Band"])

    available = unclaimed.groupby("Quadrant")["Industry"].nunique().to_dict()
    quota = allocate_quotas(available, QUEUE_A_TARGET_SHARE, notes, "Queue A")

    picked_industries: list[str] = []
    for q in QUADRANT_PRIORITY:
        sub = unclaimed[unclaimed.Quadrant == q].sort_values("Rate", ascending=False)
        picked_industries += list(sub.head(quota[q])["Industry"])

    if len(picked_industries) < 15:
        remaining = unclaimed[~unclaimed.Industry.isin(picked_industries)].sort_values("Rate", ascending=False)
        fill = list(remaining.head(15 - len(picked_industries))["Industry"])
        if fill:
            notes.append(f"Queue A: fewer than 15 industries matched quadrant quotas — "
                          f"filled {len(fill)} more from the largest uncovered industries regardless of quadrant.")
        picked_industries += fill

    used = set(core["Ticker"]) if not core.empty else set()
    rows = []
    for industry in picked_industries[:15]:
        rep = pick_representative(hits, industry, exclude_symbols=used)
        if rep is None:
            continue
        used.add(rep["Symbol"])
        q = unclaimed.set_index("Industry").loc[industry, "Quadrant"]
        rows.append({"Ticker": rep["Symbol"], "Industry": industry, "Sector": rep["Sector"],
                      "Quadrant": q, "Scan Tag": rep["ScanTag"], "Cap Band": rep["CapBand"]})
    if len(rows) < 15:
        notes.append(f"Queue A: only {len(rows)}/15 slots filled — not enough unclaimed industries in total.")
    return pd.DataFrame(rows)


# --------------------------------------------------------------------------
# STEP 6 — Queue B (Cap Contrast)
# --------------------------------------------------------------------------

def build_queue_b(participation: pd.DataFrame, hits: pd.DataFrame, core: pd.DataFrame,
                   notes: list[str]) -> pd.DataFrame:
    if core.empty:
        return pd.DataFrame(columns=["Ticker", "Industry", "Sector", "Quadrant", "Scan Tag", "Cap Band",
                                      "Contrasts With"])
    core_part = participation[participation.Industry.isin(core["Industry"])].sort_values("Hits", ascending=False)
    top15 = core_part.head(15)

    used = set(core["Ticker"])
    rows = []
    for _, prow in top15.iterrows():
        industry = prow["Industry"]
        core_pick = core[core.Industry == industry].iloc[0]
        if prow["Hits"] < 2:
            notes.append(f"Queue B: '{industry}' has only {int(prow['Hits'])} qualifying member — skipped "
                          f"(needs 2+).")
            continue
        rep = pick_representative(hits, industry, exclude_symbols=used, exclude_band=core_pick["Cap Band"])
        if rep is None:
            notes.append(f"Queue B: '{industry}' has no second stock in a different cap band — skipped.")
            continue
        used.add(rep["Symbol"])
        rows.append({"Ticker": rep["Symbol"], "Industry": industry, "Sector": rep["Sector"],
                      "Quadrant": core_pick["Quadrant"], "Scan Tag": rep["ScanTag"], "Cap Band": rep["CapBand"],
                      "Contrasts With": core_pick["Ticker"]})
    if len(rows) < 15:
        notes.append(f"Queue B: only {len(rows)}/15 slots filled — fewer than 15 Core industries had a "
                      f"usable second cap-band member.")
    return pd.DataFrame(rows)


# --------------------------------------------------------------------------
# STEP 7 — Queue C (Structure Contrast)
# --------------------------------------------------------------------------

def build_queue_c(participation: pd.DataFrame, hits: pd.DataFrame, core: pd.DataFrame, queue_b: pd.DataFrame,
                   notes: list[str]) -> pd.DataFrame:
    if core.empty:
        return pd.DataFrame(columns=["Ticker", "Industry", "Sector", "Quadrant", "Scan Tag", "Cap Band",
                                      "Contrasts With"])
    core_part = participation[participation.Industry.isin(core["Industry"])]
    both = core_part[(core_part.TrendHits > 0) & (core_part.TurnHits > 0)].sort_values("Rate", ascending=False)

    used = set(core["Ticker"]) | (set(queue_b["Ticker"]) if not queue_b.empty else set())
    rows = []
    for _, prow in both.head(15).iterrows():
        industry = prow["Industry"]
        core_pick = core[core.Industry == industry].iloc[0]
        other_family = "TURN" if core_pick["Section"] == "TREND" else "TREND"
        used_bands = {core_pick["Cap Band"]}
        if not queue_b.empty:
            b_row = queue_b[queue_b.Industry == industry]
            if not b_row.empty:
                used_bands.add(b_row.iloc[0]["Cap Band"])

        pool = hits[(hits.Industry == industry) & (hits.Family == other_family) & (~hits.Symbol.isin(used))]
        pool = pool[~pool.CapBand.isin(used_bands)]
        rep = None
        for band in CAP_BANDS:
            band_pool = pool[pool.CapBand == band]
            if not band_pool.empty:
                rep = band_pool.sort_values("MarketCap", ascending=False).iloc[0]
                break
        if rep is None and not pool.empty:
            rep = pool.sort_values("MarketCap", ascending=False).iloc[0]
        if rep is None:
            notes.append(f"Queue C: '{industry}' has no usable {other_family} stock in an unused cap band — "
                          f"skipped.")
            continue
        used.add(rep["Symbol"])
        rows.append({"Ticker": rep["Symbol"], "Industry": industry, "Sector": rep["Sector"],
                      "Quadrant": core_pick["Quadrant"], "Scan Tag": rep["ScanTag"], "Cap Band": rep["CapBand"],
                      "Contrasts With": core_pick["Ticker"]})
    if len(rows) < 15:
        notes.append(f"Queue C: only {len(rows)}/15 slots filled — fewer than 15 Core industries qualified in "
                      f"both TREND and TURN with a usable contrast stock.")
    return pd.DataFrame(rows)


# --------------------------------------------------------------------------
# Ledger & Reserve list
# --------------------------------------------------------------------------

def build_ledger(participation: pd.DataFrame, core: pd.DataFrame, queue_a: pd.DataFrame) -> pd.DataFrame:
    if participation.empty:
        return pd.DataFrame(columns=["Industry", "Hits", "Universe", "Rate %", "Won Slot In"])
    won = {}
    for industry in core.get("Industry", []):
        won[industry] = "Core 30"
    for industry in queue_a.get("Industry", []):
        won.setdefault(industry, "Queue A")
    out = participation[["Industry", "Hits", "Universe", "Rate"]].copy()
    out = out.rename(columns={"Rate": "Rate %"})
    out["Won Slot In"] = out["Industry"].map(won).fillna("—")
    return out.sort_values("Rate %", ascending=False).reset_index(drop=True)


def build_reserve_list(hits: pd.DataFrame, core: pd.DataFrame) -> pd.DataFrame:
    rows = []
    if core.empty:
        return pd.DataFrame(columns=["Industry", "Selected", "Selected Band", "Reserve"])
    for industry in core["Industry"].unique():
        selected = core[core.Industry == industry].iloc[0]
        pool = hits[(hits.Industry == industry) & (hits.Symbol != selected["Ticker"])] \
            .sort_values("MarketCap", ascending=False)
        reserves = list(pool["Symbol"].head(8))
        extra = max(0, len(pool) - 8)
        reserve_str = ", ".join(reserves) + (f" (+{extra} more)" if extra else "")
        rows.append({
            "Industry": industry, "Selected": selected["Ticker"], "Selected Band": selected["Cap Band"],
            "Reserve": reserve_str if reserve_str else "—",
        })
    return pd.DataFrame(rows).sort_values("Industry").reset_index(drop=True)


# --------------------------------------------------------------------------
# Orchestration
# --------------------------------------------------------------------------

def run_pipeline(master: pd.DataFrame, screener: dict[str, pd.DataFrame],
                  sector_quadrant: dict[str, str]) -> PipelineResult:
    result = PipelineResult()
    hits, unlisted = build_hits_table(screener, master, sector_quadrant)
    result.hits, result.unlisted = hits, unlisted

    if hits.empty:
        result.notes.append("No scan symbols matched the master database — check the uploads.")
        return result

    result.participation = compute_participation(hits, master)
    result.sector_breadth = compute_sector_breadth(hits, master)
    result.industry_structure = compute_industry_structure(hits, result.participation)

    result.core = build_core30(result.participation, hits, result.notes)
    result.queue_a = build_queue_a(result.participation, hits, result.core, result.notes)
    result.queue_b = build_queue_b(result.participation, hits, result.core, result.notes)
    result.queue_c = build_queue_c(result.participation, hits, result.core, result.queue_b, result.notes)

    # STEP 8 — de-duplicate Core > A > B > C (each build_* step above already
    # excludes symbols used by earlier queues, so this is a final safety net).
    seen: set[str] = set()
    for name in ["core", "queue_a", "queue_b", "queue_c"]:
        df = getattr(result, name)
        if df.empty:
            continue
        dupes = df["Ticker"].isin(seen)
        if dupes.any():
            result.notes.append(f"{name}: removed {dupes.sum()} duplicate ticker(s) already used upstream.")
        setattr(result, name, df[~dupes].reset_index(drop=True))
        seen |= set(getattr(result, name)["Ticker"])

    result.ledger = build_ledger(result.participation, result.core, result.queue_a)
    result.reserve = build_reserve_list(hits, result.core)
    return result


# --------------------------------------------------------------------------
# Export helpers
# --------------------------------------------------------------------------

def paste_string(df: pd.DataFrame) -> str:
    if df.empty or "Ticker" not in df.columns:
        return ""
    return "".join(f"NSE:{t}," for t in df["Ticker"])


def numbered_list(df: pd.DataFrame) -> str:
    if df.empty or "Ticker" not in df.columns:
        return ""
    return "\n".join(f"{i+1}. {t}" for i, t in enumerate(df["Ticker"]))


def build_main_workbook(result: PipelineResult) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        result.core.to_excel(writer, sheet_name="Core 30", index=False)
        result.queue_a.to_excel(writer, sheet_name="Queue A - Breadth", index=False)
        result.queue_b.to_excel(writer, sheet_name="Queue B - Cap Contrast", index=False)
        result.queue_c.to_excel(writer, sheet_name="Queue C - Structure", index=False)
        result.sector_breadth.to_excel(writer, sheet_name="Sector Breadth", index=False)
        result.industry_structure.to_excel(writer, sheet_name="Industry Structure", index=False)
        result.ledger.to_excel(writer, sheet_name="Participation Ledger", index=False)
        result.reserve.to_excel(writer, sheet_name="Reserve List", index=False)
        if not result.unlisted.empty:
            result.unlisted.to_excel(writer, sheet_name="Unlisted", index=False)
    return buf.getvalue()


def build_reserve_workbook(result: PipelineResult) -> bytes:
    tracker_rows = []
    for _, row in result.core.iterrows():
        tracker_rows.append({"Industry": row["Industry"], "Type": "Selected",
                              "NSE:TICKER": f"NSE:{row['Ticker']}", "Status": ""})
    for _, row in result.reserve.iterrows():
        for tkr in [t.strip() for t in row["Reserve"].split(",") if t.strip() and "more)" not in t]:
            tracker_rows.append({"Industry": row["Industry"], "Type": "Reserve",
                                  "NSE:TICKER": f"NSE:{tkr}", "Status": ""})
    tracker = pd.DataFrame(tracker_rows)

    import_rows = []
    for _, row in result.reserve.iterrows():
        industry_hits = result.hits[result.hits.Industry == row["Industry"]]
        bulk = "".join(f"NSE:{s}," for s in industry_hits["Symbol"].unique())
        import_rows.append({"Industry": row["Industry"], "TradingView Import String": bulk})
    import_df = pd.DataFrame(import_rows)

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        tracker.to_excel(writer, sheet_name="Tracker", index=False)
        import_df.to_excel(writer, sheet_name="TradingView Import", index=False)
    return buf.getvalue()


# --------------------------------------------------------------------------
# Streamlit UI
# --------------------------------------------------------------------------

def main():
    st.set_page_config(page_title="TVS Market Navigator", layout="wide")
    st.title("TVS Market Navigator — Watchlist Organizer")
    st.caption(
        "Clerical only: lookup, ranking by given numbers, arithmetic, formatting. "
        "No live market data, no chart-quality ranking, no price predictions."
    )

    with st.sidebar:
        st.header("1. Upload files")
        master_file = st.file_uploader("Master Classification Database (.xlsx)", type=["xlsx"])
        screener_file = st.file_uploader("Weekly Screener workbook — TVS1-TVS4 sheets (.xlsx)", type=["xlsx"])

        st.header("2. RRG sector classification")
        master_preview: Optional[pd.DataFrame] = None
        master_error: Optional[str] = None
        if master_file is not None:
            try:
                master_preview = load_master_database(master_file)
            except Exception as e:
                master_error = str(e)

        sector_quadrant: dict[str, str] = {}
        if master_error:
            st.error(f"Could not read the Master Database: {master_error}")
        elif master_preview is not None:
            sectors = sorted(master_preview["Sector"].unique())
            st.caption(f"Assign each of the {len(sectors)} sectors found in your database to its current "
                       f"RRG quadrant.")
            for sector in sectors:
                sector_quadrant[sector] = st.selectbox(sector, QUADRANTS, index=1, key=f"q_{sector}")
        else:
            st.caption("Upload a valid Master Database above — the sector list will appear here, read "
                       "straight from your file's Sector column.")

        run = st.button("Run pipeline", type="primary", use_container_width=True)

    if not run:
        st.info("Upload both files, set the RRG quadrants in the sidebar, then click **Run pipeline**.")
        return

    if master_file is None or screener_file is None:
        st.error("Please upload both the Master Database and the Screener workbook.")
        return
    if master_error:
        st.error(f"Could not read the Master Database: {master_error}")
        return

    try:
        master = master_preview if master_preview is not None else load_master_database(master_file)
        screener = load_screener(screener_file)
    except Exception as e:
        st.error(f"Could not read the uploaded files: {e}")
        return

    result = run_pipeline(master, screener, sector_quadrant)

    if result.notes:
        with st.expander("⚠️ Allocation notes (reallocations / shortfalls)", expanded=True):
            for n in result.notes:
                st.write("• " + n)

    tabs = st.tabs([
        "Core 30", "Queue A — Breadth", "Queue B — Cap Contrast", "Queue C — Structure",
        "Sector Breadth", "Industry Structure", "Participation Ledger", "Reserve List",
    ])

    def render_section(tab, df: pd.DataFrame, with_strings: bool = False):
        with tab:
            if df.empty:
                st.write("No data.")
                return
            st.dataframe(df, use_container_width=True, hide_index=True)
            if with_strings:
                col1, col2 = st.columns(2)
                col1.text_area("Paste string (TradingView)", paste_string(df), height=100)
                col2.text_area("Numbered list", numbered_list(df), height=100)

    render_section(tabs[0], result.core, with_strings=True)
    render_section(tabs[1], result.queue_a, with_strings=True)
    render_section(tabs[2], result.queue_b, with_strings=True)
    render_section(tabs[3], result.queue_c, with_strings=True)
    render_section(tabs[4], result.sector_breadth)
    render_section(tabs[5], result.industry_structure)
    render_section(tabs[6], result.ledger)
    render_section(tabs[7], result.reserve)

    if not result.unlisted.empty:
        with st.expander(f"Unlisted symbols ({len(result.unlisted)}) — not found in the master database"):
            st.dataframe(result.unlisted, use_container_width=True, hide_index=True)

    st.divider()
    today = date.today().isoformat()
    c1, c2 = st.columns(2)
    c1.download_button(
        "Download full results workbook", data=build_main_workbook(result),
        file_name=f"TVS_Core30_Queues_{today}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    c2.download_button(
        "Download Reserve List workbook", data=build_reserve_workbook(result),
        file_name=f"TVS_Reserve_List_{today}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.caption("This is a study watchlist for chart reading practice, not a recommendation to buy or sell.")


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""
=============================================================================
 2026 FANTASY DRAFT COMPANION  --  live draft-room assistant (Streamlit)
=============================================================================
A fast, local, single-file draft companion for one specific league:

    10-team, 2026 Season Redraft, Sleeper, Snake draft, pick 8 of 16 rounds.
    Roster: 1 QB / 2 RB / 2 WR / 1 TE / 2 FLEX (W/R/T) / 1 SUPERFLEX (W/R/T/Q)
            / 7 BN / 2 IR

WHAT THIS IS
------------
This app CONSUMES the finished board in players_2026.csv. It does not
recompute scoring, VBD, tiers, or ADP -- GlobalValue / GlobalRank / Tier
(int 1..12) / ADP / ValueDelta are taken exactly as build_draft_kit_2026.py
produced them. (This is the deliberate v7 decision: no live-recalculating
engine substituted in for the static board mid-draft.)

WHAT IT ADDS over the Excel workbook
-----------------------------------
- Real drafted-state management: inline "Draft" / "Gone" buttons on every
  board row (Gone = an opponent took him). Both remove him from the board;
  only "Draft" fills your roster slots and fires the guardrails. Press the
  same button again to un-draft; press the other to swap owner cleanly.
- Crash safety: every action autosaves to draft_state.json next to this
  script. A browser refresh, a Streamlit rerun, or a tab crash mid-round
  recovers exactly where you were. Undo-last-pick and a confirm-gated
  Reset are in the sidebar.
- Full filterable colour-coded board (search / position / bye-week /
  contract-year / hide-no-ADP + sort) with a focus drawer (tick a row) for
  the full cautions / ADP-delta / SoS / bye breakdown before you commit.
- 2026 NFL bye weeks, ingested from the FantasyPros bye-week cheatsheet by
  cross-referencing player -> team. Shown on the board and in the drawer,
  filterable and sortable, with a bye-stacking guardrail.
- A sticky sidebar: snake pick clock, live roster by slot (greedy-by-
  GlobalValue fill, same logic as simulate_draft.py), roster guardrails
  (3rd-TE bloat, Superflex-QB pacing, bye stacking), roster bye spread,
  and live positional / tier scarcity.
- 16-round roster export as CSV and as a text summary (bye column included).

NOT here (deliberately): a playoff-weeks (14-17) strength-of-schedule
metric. The pipeline only carries the real, cross-verified Weeks 1-4
schedule; a playoff-weeks index needs the Week 14-17 matchups sourced and
tested, which is deferred rather than faked.

HOW TO RUN
----------
    pip install streamlit pandas
    streamlit run draft_app.py

The pure helper functions (snake math, slot assignment, guardrails) are
importable without Streamlit; see test_draft_app.py.
"""

from __future__ import annotations

import collections
import csv
import json
import re
import time
from datetime import datetime
from pathlib import Path

import pandas as pd

# =============================================================================
# LEAGUE CONSTANTS  (hardcoded to this one league, same as the kit builder)
# =============================================================================

N_TEAMS = 10
N_ROUNDS = 16
MY_DRAFT_SLOT = 8          # 1-indexed pick slot, pick 8 of 10
BENCH_SLOTS = 7            # 9 starters + 7 bench = 16 (IR not drafted here)

STARTER_SLOTS = ["QB", "RB1", "RB2", "WR1", "WR2", "TE", "FLEX1", "FLEX2", "SUPERFLEX"]
FLEX_ELIGIBLE = {"RB", "WR", "TE"}
SF_ELIGIBLE = {"QB", "RB", "WR", "TE"}

STARTER_NEED = {"QB": 1, "RB": 2, "WR": 2, "TE": 1}   # dedicated (non-flex) starter slots

# "Superflex-caliber" QB / "elite TE-premium" TE -- same PosTier cutoffs the
# kit builder's SF/TEP tag column and Draft Pivot text use.
SF_QB_POSTIER_CUTOFF = 4
ELITE_TE_POSTIER_CUTOFF = 2

CSV_PATH = Path(__file__).with_name("players_2026.csv")
BYE_CHEATSHEET_PATH = Path(__file__).with_name(
    "FantasyPros_Fantasy_Football_Bye_Week_Cheatsheet.csv")
FULL_SOS_PATH = Path(__file__).with_name(
    "FantasyPros_Fantasy_Football_2026_Strength_Of_Schedule.csv")
STATE_PATH = Path(__file__).with_name("draft_state.json")
STATE_VERSION = 1

# Full-season SoS label vocabulary -- deliberately the SAME words the Weeks-1-4
# `SoS_Label` uses, so the two horizons read on one scale (easiest -> hardest).
# FantasyPros stars: 1 = toughest schedule .. 5 = easiest.
FULL_SOS_LABEL = {5: "Very Soft", 4: "Soft", 3: "Neutral", 2: "Tough", 1: "Gauntlet"}

NFL_NAME_TO_CODE = {
    "Arizona Cardinals": "ARI", "Atlanta Falcons": "ATL", "Baltimore Ravens": "BAL",
    "Buffalo Bills": "BUF", "Carolina Panthers": "CAR", "Chicago Bears": "CHI",
    "Cincinnati Bengals": "CIN", "Cleveland Browns": "CLE", "Dallas Cowboys": "DAL",
    "Denver Broncos": "DEN", "Detroit Lions": "DET", "Green Bay Packers": "GB",
    "Houston Texans": "HOU", "Indianapolis Colts": "IND", "Jacksonville Jaguars": "JAC",
    "Kansas City Chiefs": "KC", "Las Vegas Raiders": "LV", "Los Angeles Chargers": "LAC",
    "Los Angeles Rams": "LAR", "Miami Dolphins": "MIA", "Minnesota Vikings": "MIN",
    "New England Patriots": "NE", "New Orleans Saints": "NO", "New York Giants": "NYG",
    "New York Jets": "NYJ", "Philadelphia Eagles": "PHI", "Pittsburgh Steelers": "PIT",
    "San Francisco 49ers": "SF", "Seattle Seahawks": "SEA", "Tampa Bay Buccaneers": "TB",
    "Tennessee Titans": "TEN", "Washington Commanders": "WAS",
}

# Drafting 3+ starters who share a bye week means a real hole in that week's
# lineup; 4+ is a near-guaranteed loss that week.
BYE_STACK_WARN = 3
BYE_STACK_ERROR = 4

OWNER_ME = "ME"
OWNER_OTHER = "OTHER"


# =============================================================================
# PURE HELPERS  (no Streamlit -- unit-tested in test_draft_app.py)
# =============================================================================

def snake_pick_slots(draft_slot: int = MY_DRAFT_SLOT,
                     n_teams: int = N_TEAMS,
                     n_rounds: int = N_ROUNDS) -> list[int]:
    """Overall pick numbers belonging to `draft_slot` across a snake draft.

    Slot 8 of 10 -> [8, 13, 28, 33, 48, 53, ...].
    """
    picks = []
    for rnd in range(1, n_rounds + 1):
        if rnd % 2 == 1:
            overall = (rnd - 1) * n_teams + draft_slot
        else:
            overall = (rnd - 1) * n_teams + (n_teams - draft_slot + 1)
        picks.append(overall)
    return picks


def overall_to_round_pick(overall: int, n_teams: int = N_TEAMS) -> tuple[int, int]:
    """1-indexed overall pick -> (round, pick_within_round)."""
    rnd = (overall - 1) // n_teams + 1
    pick_in_round = (overall - 1) % n_teams + 1
    return rnd, pick_in_round


def next_my_pick(current_overall: int, my_slots: list[int]) -> tuple[int | None, int]:
    """Return (next scheduled overall pick >= current_overall, picks_until).

    If current_overall is itself one of my slots, picks_until is 0
    ("on the clock"). If the draft is past my last pick, returns (None, 0).
    """
    upcoming = [s for s in my_slots if s >= current_overall]
    if not upcoming:
        return None, 0
    nxt = min(upcoming)
    return nxt, nxt - current_overall


def upcoming_my_picks(current_overall: int,
                      my_slots: list[int]) -> tuple[int | None, int | None]:
    """(your next scheduled overall pick, the one after that). Either can be
    None near the end of the draft."""
    up = sorted(s for s in my_slots if s >= current_overall)
    return (up[0] if up else None, up[1] if len(up) > 1 else None)


# The ADP column is a 12-team Superflex tool; our room is 10-team, so a
# player goes a bit earlier here than his raw ADP number. Scale pick
# position down by this factor when judging "will he make it back to me?".
ADP_TEAM_SCALE = N_TEAMS / 12.0


def adp_pick_risk(adp, next_pick: int | None, pick_after: int | None,
                  scale: float = ADP_TEAM_SCALE) -> str:
    """Can you safely let this player slide past your next pick?

    Returns:
        ""       -- yes: his (team-scaled) ADP is at/after your pick-after-next
        "fringe" -- 50/50: ADP lands between your next two picks
        "gone"   -- no: ADP is at/before your very next pick
    """
    if adp is None or pd.isna(adp) or not next_pick:
        return ""
    eff = float(adp) * scale
    if eff <= next_pick:
        return "gone"
    if pick_after and eff < pick_after:
        return "fringe"
    return ""


def _global_value(row: dict) -> float:
    gv = row.get("GlobalValue")
    try:
        gv = float(gv)
    except (TypeError, ValueError):
        return float("-inf")
    return gv if pd.notna(gv) else float("-inf")


def assign_roster(my_rows: list[dict]) -> tuple[dict, list[dict]]:
    """Greedy slot fill, best (highest GlobalValue) player first.

    Identical strategy to simulate_draft.py's assign_slots: fill the
    dedicated slots, then the two FLEX, then SUPERFLEX, remainder to bench.
    Returns ({slot_label: row|None}, [bench rows]).
    """
    players = sorted(my_rows, key=_global_value, reverse=True)
    remaining = list(players)
    slots: dict = {}

    def take(pos_set: set[str], label: str) -> None:
        for i, p in enumerate(remaining):
            if p.get("Pos") in pos_set:
                slots[label] = remaining.pop(i)
                return
        slots[label] = None

    take({"QB"}, "QB")
    take({"RB"}, "RB1")
    take({"RB"}, "RB2")
    take({"WR"}, "WR1")
    take({"WR"}, "WR2")
    take({"TE"}, "TE")
    take(FLEX_ELIGIBLE, "FLEX1")
    take(FLEX_ELIGIBLE, "FLEX2")
    take(SF_ELIGIBLE, "SUPERFLEX")
    return slots, remaining


def position_counts(my_rows: list[dict]) -> dict[str, int]:
    counts = {"QB": 0, "RB": 0, "WR": 0, "TE": 0}
    for r in my_rows:
        pos = r.get("Pos")
        if pos in counts:
            counts[pos] += 1
    return counts


def _clean(v) -> str:
    """Empty string for None / NaN / literal 'nan'; stripped str otherwise.

    (Guards the `x or ""` trap: a float NaN is truthy, so `nan or ""` is nan.)
    """
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    s = str(v).strip()
    return "" if s.lower() == "nan" else s


def _short_note(row: dict, limit: int = 140) -> str:
    note = _clean(row.get("Notes"))
    if len(note) > limit:
        note = note[:limit].rsplit(" ", 1)[0] + "..."
    return note


def candidate_cautions(row: dict, my_counts: dict[str, int]) -> list[str]:
    """Pre-draft flags for a player you're about to take (informational)."""
    out: list[str] = []
    pos = row.get("Pos")

    if pos == "TE" and my_counts.get("TE", 0) >= 2:
        out.append(f"WARNING  This would be TE #{my_counts['TE'] + 1} -- TE-premium bloat trap; "
                   f"a 3rd TE only rides your bench.")
    if pos == "QB" and my_counts.get("QB", 0) >= 2:
        out.append(f"WARNING  QB #{my_counts['QB'] + 1} -- in this Superflex format a 3rd QB is "
                   f"bench / bye-week cover only.")

    vd = row.get("ValueDelta")
    adp = row.get("ADP")
    grk = row.get("GlobalRank")
    if pd.notna(vd):
        vd = float(vd)
        if vd <= -10:
            out.append(f"REACH  ADP {adp:.0f} vs global rank {int(grk)} (delta {vd:+.0f}) -- "
                       f"the market says you can wait.")
        elif vd >= 10:
            out.append(f"VALUE  sliding {vd:+.0f} past the model's rank {int(grk)} (ADP {adp:.0f}).")

    flag = _clean(row.get("Flag")).upper()
    if flag in {"BUST", "GAMBLE"}:
        note = _short_note(row)
        out.append(f"{flag}  {note}" if note else f"{flag} flag on this player.")

    dur = _clean(row.get("DurabilityNote"))
    if dur:
        out.append(f"DURABILITY  {dur}")

    return out


def roster_guardrails(my_rows: list[dict]) -> list[tuple[str, str]]:
    """(level, message) pairs. level in {'error','warning','info','success'}.

    Pacing keys off how many picks YOU have made (== your current round),
    not the global pick number.
    """
    out: list[tuple[str, str]] = []
    counts = position_counts(my_rows)
    my_round = len(my_rows)                # after your Nth pick you're "in round N"
    qb, rb, wr, te = counts["QB"], counts["RB"], counts["WR"], counts["TE"]

    # --- TE-premium bloat -------------------------------------------------
    if te >= 3:
        out.append(("error", f"TE bloat: {te} TEs rostered. TE premium tempts overdraft -- "
                             f"everything past your 2nd TE is bench only."))

    # --- Superflex QB pacing -------------------------------------------------
    if qb >= 3:
        out.append(("info", f"{qb} QBs rostered -- a 3rd QB is bench / bye-week cover in this format."))
    if my_round >= 9 and qb < 2:
        out.append(("warning", f"Only {qb} QB through {my_round} of your picks -- you need 2 startable "
                               f"arms (QB + SUPERFLEX). Superflex rooms fill both slots."))
    elif my_round >= 5 and qb == 0:
        out.append(("warning", f"No QB through {my_round} of your picks -- don't get boxed out of the "
                               f"Superflex-caliber QB pool."))

    # --- Late-draft starter holes -----------------------------------------
    if my_round >= 12:
        for pos, need in STARTER_NEED.items():
            have = counts[pos]
            if have < need:
                out.append(("warning", f"Still short at {pos}: {have}/{need} startable with "
                                       f"{max(0, N_ROUNDS - my_round)} picks left."))

    # --- Bye-week stacking ----------------------------------------------
    for wk, names in sorted(roster_bye_counts(my_rows).items()):
        if len(names) >= BYE_STACK_ERROR:
            out.append(("error", f"Bye stack: {len(names)} players idle in Week {wk} "
                                 f"({', '.join(names)}) -- that week is a write-off."))
        elif len(names) >= BYE_STACK_WARN:
            out.append(("warning", f"Bye stack: {len(names)} on the Week {wk} bye "
                                   f"({', '.join(names)}) -- likely a lineup hole that week."))

    # --- Done -----------------------------------------------------------
    if my_round >= N_ROUNDS:
        out.append(("success", f"Roster complete -- {my_round} picks in."))
    elif not out:
        out.append(("success", f"Roster on track (QB {qb} / RB {rb} / WR {wr} / TE {te})."))

    return out


def roster_bye_counts(my_rows: list[dict]) -> dict[int, list[str]]:
    """{bye_week: [player names]} for the drafted players that have a bye."""
    byes: dict[int, list[str]] = {}
    for r in my_rows:
        b = r.get("Bye")
        if b is not None and pd.notna(b):
            byes.setdefault(int(b), []).append(str(r.get("Name", "?")))
    return byes


def bye_stack_caution(player_bye, my_rows: list[dict]) -> str | None:
    """Pre-draft flag: does this player's bye already hold >= 2 of your guys?"""
    if player_bye is None or pd.isna(player_bye):
        return None
    wk = int(player_bye)
    same = [str(r.get("Name", "?")) for r in my_rows
            if r.get("Bye") is not None and pd.notna(r.get("Bye")) and int(r["Bye"]) == wk]
    if len(same) >= 2:
        return (f"BYE STACK  this makes {len(same) + 1} on the Week {wk} bye "
                f"({', '.join(same)}) -- you'll be starting scrubs that week.")
    return None


def scarcity(available: pd.DataFrame) -> dict:
    """Live counts of what's left on the board."""
    by_pos = {p: int((available["Pos"] == p).sum()) for p in ["QB", "RB", "WR", "TE"]}
    by_tier = (available.groupby("Tier").size().sort_index().to_dict())
    sf_qb_left = int(((available["Pos"] == "QB") &
                     (available["PosTier"] <= SF_QB_POSTIER_CUTOFF)).sum())
    elite_te_left = int(((available["Pos"] == "TE") &
                         (available["PosTier"] <= ELITE_TE_POSTIER_CUTOFF)).sum())
    return {
        "by_pos": by_pos,
        "by_tier": {int(k): int(v) for k, v in by_tier.items()},
        "sf_qb_left": sf_qb_left,
        "elite_te_left": elite_te_left,
    }


def tier_cliffs(available: pd.DataFrame, tier_col: str = "Tier",
                low: int = 2) -> list[tuple[int, int]]:
    """Tiers about to dry up: `low` or fewer players still available AND at
    least one player still sitting in a deeper tier (so running out is a real
    value drop-off, not just the end of the pool). (tier, remaining) pairs,
    best (lowest-numbered) tier first.
    """
    if tier_col not in getattr(available, "columns", []) or available.empty:
        return []
    counts = available.groupby(tier_col).size()
    counts = counts[counts.index.notna()]
    if counts.empty:
        return []
    max_tier = int(max(int(x) for x in counts.index))
    out: list[tuple[int, int]] = []
    for t in sorted(int(x) for x in counts.index):
        n = int(counts.loc[t])
        if 1 <= n <= low and t < max_tier:
            out.append((t, n))
    return out


_FLAG_FLASH = {
    "STUD": ("Model Stud", "green"),
    "FLOOR": ("Safe Floor", "green"),
    "SLEEPER": ("Late Sleeper", "blue"),
    "GAMBLE": ("Boom / Bust", "orange"),
    "BUST": ("Priced Up — fade", "red"),
    "HANDCUFF": ("Handcuff", "gray"),
    "DEPTH": ("Bench Depth", "gray"),
}
_INJURY_KEYWORDS = (
    ("hamstring", "Hamstring Risk"), ("groin", "Groin Risk"),
    ("ankle", "Ankle Risk"), ("knee", "Knee Risk"), ("acl", "ACL History"),
    ("psoas", "Psoas Risk"), ("concussion", "Concussion History"),
    ("shoulder", "Shoulder Risk"), ("achilles", "Achilles History"),
)
_VOLUME_KEYWORDS = (
    "bell-cow", "bell cow", "high-volume", "target share", "true wr1",
    "workhorse", "every-down", "three-down", "lead back", "bell cow",
)


def flash_tags(row: dict) -> list[tuple[str, str]]:
    """Punchy plain-language (label, colour) chips for the focus drawer -- risk
    and upside processable in ~2 seconds. Derived only from data already on the
    row (Flag / ContractYear / DurabilityNote / Notes / overlay columns)."""
    tags: list[tuple[str, str]] = []

    flag = _clean(row.get("Flag")).upper()
    if flag in _FLAG_FLASH:
        tags.append(_FLAG_FLASH[flag])
    if str(row.get("ContractYear", "")).strip().upper() == "Y":
        tags.append(("Contract Year", "violet"))

    dur = _clean(row.get("DurabilityNote")).lower()
    for kw, label in _INJURY_KEYWORDS:
        if kw in dur:
            tags.append((label, "red"))
            break
    if bool(row.get("Monitor")):
        tags.append(("Injury Watch", "red"))
    if bool(row.get("HighOffense")):
        tags.append(("Elite Offense", "green"))
    hcf = _clean(row.get("HandcuffFor"))
    if hcf:
        tags.append((f"Backs up {hcf}", "gray"))

    vd = row.get("ValueDelta")
    if pd.notna(vd):
        vd = float(vd)
        if vd >= 15:
            tags.append(("Model Value", "green"))
        elif vd <= -15:
            tags.append(("Market Reach", "orange"))

    note = (_clean(row.get("Notes")) + " " + _clean(row.get("NewsNote"))).lower()
    if any(k in note for k in _VOLUME_KEYWORDS):
        tags.append(("High-Volume Role", "green"))

    seen: set[str] = set()
    return [t for t in tags if not (t[0] in seen or seen.add(t[0]))]


def resolve_action(current: str | None, clicked: str) -> tuple[str, str | None]:
    """Inline toggle/swap state machine for the board's Draft/Gone buttons.

    `current` is the player's existing owner (None / "ME" / "OTHER");
    `clicked` is which button was pressed ("ME" or "OTHER"). Returns one of:
        ("add", clicked)     -- player was undrafted, claim him
        ("remove", None)     -- same button re-pressed, un-draft (toggle off)
        ("swap", clicked)    -- other owner pressed, reassign cleanly
    """
    if current == clicked:
        return ("remove", None)
    if current is None:
        return ("add", clicked)
    return ("swap", clicked)


# =============================================================================
# DRAFT-DAY NEWS OVERLAY  (draft_day_news.py, layered onto the CSV at load)
# =============================================================================

def apply_news_overlay(df: pd.DataFrame, news: dict,
                       high_offense: set[str]) -> pd.DataFrame:
    """Add overlay columns from draft_day_news.py -- non-destructive, the base
    situational research in `Notes` is never overwritten.

        NewsNote     -- fresh hand-entered note ("" if none)
        Monitor      -- bool, injury-monitored
        HandcuffFor  -- starter this player directly backs up ("" if none)
        HighOffense  -- bool, team is in HIGH_OFFENSE_TEAMS
    """
    df = df.copy()
    news = news or {}
    high = set(high_offense or set())

    def _get(name, key, default=""):
        return (news.get(name) or {}).get(key, default)

    df["NewsNote"] = df["Name"].map(lambda n: str(_get(n, "note", "") or "")).fillna("")
    df["Monitor"] = df["Name"].map(lambda n: bool(_get(n, "monitor", False)))
    df["HandcuffFor"] = df["Name"].map(
        lambda n: str(_get(n, "handcuff_for", "") or "")).fillna("")
    df["HighOffense"] = df["Team"].isin(high)
    return df


def effective_note(row: dict) -> str:
    """Draft-day note first (🆕-prefixed), the researched note kept behind it."""
    fresh = _clean(row.get("NewsNote"))
    base = _clean(row.get("Notes"))
    if not fresh:
        return base
    if not base:
        return f"🆕 {fresh}"
    return f"🆕 {fresh}  ·  {base}"


# =============================================================================
# STRENGTH OF SCHEDULE  (full-season, alongside the Weeks-1-4 metric)
# =============================================================================

def parse_full_sos(path: Path) -> dict[tuple[str, str], int]:
    """{(TEAM_CODE, POS): star} from the FantasyPros full-season SoS CSV.

    Stars are 1 (toughest schedule) .. 5 (easiest) and are split by position
    (QB / RB / WR / TE). Returns {} if the file is missing or unreadable.
    """
    if not path.exists():
        return {}
    try:
        raw = pd.read_csv(path, skiprows=1)      # line 1 is the "Star ratings:" note
    except Exception:
        return {}
    raw.columns = [str(c).strip() for c in raw.columns]
    if "TEAM" not in raw.columns:
        return {}
    out: dict[tuple[str, str], int] = {}
    for _, r in raw.iterrows():
        code = NFL_NAME_TO_CODE.get(str(r["TEAM"]).strip())
        if not code:
            continue
        for pos in ("QB", "RB", "WR", "TE"):
            v = pd.to_numeric(r.get(pos), errors="coerce")
            if pd.notna(v) and 1 <= int(v) <= 5:
                out[(code, pos)] = int(v)
    return out


def full_sos_label(star) -> str:
    """Star (1..5, 5 = easiest) -> the shared SoS label. '—' if unknown."""
    if star is None or pd.isna(star):
        return "—"
    return FULL_SOS_LABEL.get(int(star), "—")


# =============================================================================
# BYE-WEEK INGESTION
# =============================================================================

def parse_team_byes(cheatsheet_path: Path,
                    name_to_team: dict[str, str]) -> dict[str, int]:
    """Build a {TEAM: bye_week} map from the FantasyPros bye-week cheatsheet.

    The cheatsheet is sectioned by ``"Week N Bye"`` headers followed by rows
    of player names (with an ECR number after each). Bye weeks are a team
    property, so we read every listed player, look up the team we already
    know for them from players_2026.csv, and take the majority week per team
    (robust to the odd offseason move). Returns {} if the file is absent.
    """
    if not cheatsheet_path.exists():
        return {}
    votes: dict[str, collections.Counter] = collections.defaultdict(collections.Counter)
    week: int | None = None
    try:
        with open(cheatsheet_path, newline="", encoding="utf-8") as fh:
            for row in csv.reader(fh):
                if not row:
                    continue
                header = re.match(r"\s*Week\s+(\d+)\s+Bye", row[0])
                if header:
                    week = int(header.group(1))
                    continue
                if week is None:
                    continue
                for cell in row:
                    name = cell.strip()
                    if not name or name.upper() == "ECR":
                        continue
                    if name.replace(".", "", 1).isdigit():   # an ECR number
                        continue
                    team = name_to_team.get(name)
                    if team and team != "FA":
                        votes[team][week] += 1
    except OSError:
        return {}
    return {team: cnt.most_common(1)[0][0] for team, cnt in votes.items() if cnt}


# =============================================================================
# STATE PERSISTENCE
# =============================================================================

def load_state() -> dict:
    if STATE_PATH.exists():
        try:
            data = json.loads(STATE_PATH.read_text(encoding="utf-8"))
            if isinstance(data, dict) and isinstance(data.get("picks"), list):
                return data
        except (json.JSONDecodeError, OSError):
            pass
    return {"version": STATE_VERSION, "picks": []}


def save_state(picks: list[dict]) -> None:
    payload = {"version": STATE_VERSION, "picks": picks,
               "saved_at": datetime.now().isoformat(timespec="seconds")}
    tmp = STATE_PATH.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    tmp.replace(STATE_PATH)


# =============================================================================
# STREAMLIT APP
# =============================================================================

def _vd_bg(v):
    """Green (slide / discount) -> red (reach) background for ValueDelta."""
    if pd.isna(v):
        return ""
    v = float(v)
    if v >= 12:
        return "background-color: #1b5e20; color: #ffffff"
    if v >= 4:
        return "background-color: #388e3c; color: #ffffff"
    if v > -4:
        return ""
    if v > -12:
        return "background-color: #c62828; color: #ffffff"
    return "background-color: #8e0000; color: #ffffff"


# Keyed to the real Flag vocabulary in players_2026.csv:
# STUD / FLOOR / SLEEPER / GAMBLE / HANDCUFF / BUST / DEPTH (no "CAUTION" --
# that word only ever appears inside Notes text, never in the Flag column).
_FLAG_BG = {
    "STUD": "background-color: #1b5e20; color: #ffffff",
    "FLOOR": "background-color: #1565c0; color: #ffffff",
    "SLEEPER": "background-color: #00695c; color: #ffffff",
    "GAMBLE": "background-color: #ef6c00; color: #ffffff",
    "HANDCUFF": "background-color: #455a64; color: #ffffff",   # slate = insurance
    "BUST": "background-color: #8e0000; color: #ffffff",
    "DEPTH": "color: #999999",
}


def _flag_bg(v):
    """Colour-code the situational Flag tag (STUD / FLOOR / HANDCUFF / BUST ...)."""
    return _FLAG_BG.get(str(v).strip().upper(), "")


def run_app() -> None:
    import streamlit as st

    st.set_page_config(page_title="2026 Draft Companion", page_icon="🏈", layout="wide")

    if not CSV_PATH.exists():
        st.error(f"Cannot find {CSV_PATH.name} next to this script. "
                 f"Run `python build_draft_kit_2026.py` first to generate it.")
        st.stop()

    @st.cache_data(show_spinner=False)
    def load_players() -> pd.DataFrame:
        df = pd.read_csv(CSV_PATH)
        df["Name"] = df["Name"].astype(str)
        # NOTE the column names: `Tier` is the integer 1..12, `TierLevel` is
        # the text label "Tier 1".."Tier 12" (persistent display string).
        for col in ["ADP", "ValueDelta", "GlobalValue", "GlobalRank",
                    "Tier", "PosTier", "ECR_Pos", "SoS_Rank"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce")
        # 2026 NFL bye weeks: team property, ingested from the FantasyPros
        # bye-week cheatsheet by cross-referencing player -> team. Nullable
        # Int64 so unknown byes sort to the bottom like the no-ADP players.
        team_bye = parse_team_byes(BYE_CHEATSHEET_PATH,
                                   dict(zip(df["Name"], df["Team"])))
        df["Bye"] = df["Team"].map(team_bye).astype("Int64")

        # Full-season SoS (FantasyPros, position-split). Kept on the SAME
        # label scale as the Weeks-1-4 `SoS_Label`. FullSoS_Rank is 1 (easiest)
        # .. 5 (hardest) so it sorts the same direction as SoS_Rank.
        sos_map = parse_full_sos(FULL_SOS_PATH)
        star = df.apply(lambda r: sos_map.get((r["Team"], r["Pos"])), axis=1)
        df["FullSoS_Star"] = pd.to_numeric(star, errors="coerce").astype("Int64")
        df["FullSoS_Rank"] = (6 - df["FullSoS_Star"]).astype("Int64")
        df["FullSoS_Label"] = [
            "N/A (FA)" if t == "FA" else full_sos_label(s)
            for t, s in zip(df["Team"], df["FullSoS_Star"])
        ]
        return df

    def load_news() -> tuple[dict, set]:
        """draft_day_news.py, re-read every rerun so mid-draft edits land."""
        try:
            import importlib
            import draft_day_news
            importlib.reload(draft_day_news)
            return (getattr(draft_day_news, "NEWS", {}) or {},
                    set(getattr(draft_day_news, "HIGH_OFFENSE_TEAMS", set()) or set()))
        except Exception:
            return {}, set()

    _base = load_players()
    _news, _high_off = load_news()
    # Overlay is applied OUTSIDE the cache (328 rows, trivial) so edits to
    # draft_day_news.py take effect on the next Rerun with no cache clear.
    players = apply_news_overlay(_base, _news, _high_off)
    bye_known = players["Bye"].notna().any()
    news_unmatched = sorted(n for n in _news if n not in set(players["Name"]))
    by_name = {r["Name"]: r for r in players.to_dict("records")}

    # ---- session <- disk (once) ----------------------------------------
    if "picks" not in st.session_state:
        st.session_state.picks = load_state()["picks"]

    picks: list[dict] = st.session_state.picks
    drafted_names = {p["name"] for p in picks}
    my_names = [p["name"] for p in picks if p["owner"] == OWNER_ME]
    my_rows = [by_name[n] for n in my_names if n in by_name]
    my_counts = position_counts(my_rows)

    def record_pick(name: str, owner: str) -> None:
        # Check LIVE state, not the run-start `drafted_names` snapshot -- an
        # inline swap does remove_pick + record_pick in one callback, and a
        # stale snapshot would wrongly block the re-add.
        if any(p["name"] == name for p in st.session_state.picks):
            return
        st.session_state.picks.append({
            "name": name,
            "pos": by_name.get(name, {}).get("Pos"),
            "team": by_name.get(name, {}).get("Team"),
            "owner": owner,
            "overall": len(st.session_state.picks) + 1,
            "ts": time.time(),
        })
        save_state(st.session_state.picks)

    def undo_last() -> None:
        if st.session_state.picks:
            st.session_state.picks.pop()
            save_state(st.session_state.picks)

    def remove_pick(name: str) -> None:
        st.session_state.picks = [p for p in st.session_state.picks if p["name"] != name]
        for i, p in enumerate(st.session_state.picks, start=1):
            p["overall"] = i
        save_state(st.session_state.picks)

    def reset_draft() -> None:
        st.session_state.picks = []
        save_state(st.session_state.picks)

    available = players[~players["Name"].isin(drafted_names)].copy()

    # =====================================================================
    # SIDEBAR
    # =====================================================================
    my_slots = snake_pick_slots()
    picks_made = len(picks)
    current_overall = picks_made + 1
    cur_round, cur_in_round = overall_to_round_pick(current_overall)
    nxt, away = next_my_pick(current_overall, my_slots)

    with st.sidebar:
        if news_unmatched:
            st.warning("draft_day_news.py names not found in the board: "
                       + ", ".join(news_unmatched))
        st.markdown("### Pick clock")
        c1, c2 = st.columns(2)
        c1.metric("Overall pick", f"#{current_overall}")
        c2.metric("Round", f"{cur_round}.{cur_in_round:02d}")
        if nxt is None:
            st.caption("Your draft is done -- all 16 picks are in.")
        elif away == 0:
            st.success("YOU ARE ON THE CLOCK")
        else:
            st.info(f"Your next pick: #{nxt}  ({away} pick{'s' if away != 1 else ''} away)")
        after_next = [s for s in my_slots if s > (nxt or current_overall)][:4]
        if after_next:
            st.caption("Then: " + " -> ".join(f"#{s}" for s in after_next))

        st.divider()
        st.markdown("### My roster")
        slots, bench = assign_roster(my_rows)
        for label in STARTER_SLOTS:
            r = slots.get(label)
            if r:
                st.markdown(f"**{label}** &nbsp; {r['Name']} · {r['Pos']}-{r['Team']}")
            else:
                st.markdown(f"**{label}** &nbsp; <span style='color:#888'>— empty —</span>",
                            unsafe_allow_html=True)
        st.markdown(f"**Bench** &nbsp; {len(bench)}/{BENCH_SLOTS} used")
        for r in bench[:BENCH_SLOTS]:
            st.caption(f"BN · {r['Name']} ({r['Pos']}-{r['Team']})")
        if len(bench) > BENCH_SLOTS:
            st.warning(f"{len(bench) - BENCH_SLOTS} pick(s) over the 16-man limit.")
        st.caption(f"Mix: QB {my_counts['QB']} · RB {my_counts['RB']} · "
                   f"WR {my_counts['WR']} · TE {my_counts['TE']}  "
                   f"(total {len(my_rows)}/{N_ROUNDS})")
        _byes = roster_bye_counts(my_rows)
        if _byes:
            _bstr = " · ".join(
                f"Wk{wk}: {len(names)}" + ("  ⚠" if len(names) >= BYE_STACK_WARN else "")
                for wk, names in sorted(_byes.items()))
            st.caption(f"Byes: {_bstr}")

        st.divider()
        st.markdown("### Guardrails")
        for level, msg in roster_guardrails(my_rows):
            getattr(st, level)(msg)

        st.divider()
        st.markdown("### Scarcity left on board")
        sc = scarcity(available)
        p = sc["by_pos"]
        st.caption(f"QB {p['QB']} · RB {p['RB']} · WR {p['WR']} · TE {p['TE']} available")
        _np2, _pa2 = upcoming_my_picks(current_overall, my_slots)
        _target = _pa2 if away == 0 else _np2
        if _target is not None:
            _gone = int((available["ADP"].dropna() * ADP_TEAM_SCALE <= _target).sum())
            st.caption(f"🔥 ~{_gone} available now project gone by your "
                       f"pick #{_target}")
        st.caption(f"Superflex-caliber QB left (PosTier<={SF_QB_POSTIER_CUTOFF}): "
                   f"**{sc['sf_qb_left']}**")
        st.caption(f"Elite TE-premium TE left (PosTier<={ELITE_TE_POSTIER_CUTOFF}): "
                   f"**{sc['elite_te_left']}**")
        if available["HighOffense"].any():
            _ho = int((available["HighOffense"] & available["ADP"].notna()
                       & (available["ADP"] >= 100)).sum())
            st.caption(f"⚡ high-offense ceiling targets left (ADP ≥ 100): **{_ho}**")

        # --- tier-cliff warnings (don't get caught sleeping on a run) -----
        _gcliffs = tier_cliffs(available, "Tier", low=2)
        for t, n in _gcliffs[:3]:
            st.warning(f"⛰️ **Global Tier {t}** — only {n} left, then the "
                       f"next value drop-off.")
        _pos_lines = []
        for pos in ("QB", "RB", "WR", "TE"):
            for pt, n in tier_cliffs(available[available["Pos"] == pos],
                                     "PosTier", low=1)[:1]:
                _pos_lines.append(f"{pos} tier {pt}: {n} left")
        if _pos_lines:
            st.caption("⛰️ positional last-in-tier — " + " · ".join(_pos_lines))

        with st.expander("Remaining by NFL tier"):
            tier_df = pd.DataFrame(
                [{"Tier": k, "Players left": v} for k, v in sc["by_tier"].items()]
            )
            st.dataframe(tier_df, hide_index=True, use_container_width=True)

        st.divider()
        st.markdown("### Draft controls")
        st.button("Undo last pick", on_click=undo_last, use_container_width=True,
                  disabled=not picks)
        with st.expander("Reset entire draft"):
            confirm = st.checkbox("Yes, wipe all picks and start over")
            st.button("RESET DRAFT", on_click=reset_draft, type="primary",
                      use_container_width=True, disabled=not confirm)
        st.caption(f"Autosaving to `{STATE_PATH.name}`")

    # =====================================================================
    # MAIN
    # =====================================================================
    st.title("2026 Fantasy Draft Companion")
    st.progress(min(picks_made / (N_TEAMS * N_ROUNDS), 1.0),
                text=f"{picks_made} / {N_TEAMS * N_ROUNDS} picks made")

    tab_board, tab_roster, tab_history = st.tabs(
        ["Available board", "My roster", "Draft history"])

    # ------------------------------------------------------------------ BOARD
    with tab_board:
        # ---- inline action state machine -------------------------------
        def owner_of(name: str):
            for p in st.session_state.picks:
                if p["name"] == name:
                    return p["owner"]
            return None

        def _queue_cautions(name: str) -> None:
            mine_rows = [by_name[q["name"]] for q in st.session_state.picks
                         if q["owner"] == OWNER_ME and q["name"] != name
                         and q["name"] in by_name]
            cnt = position_counts(mine_rows)
            msgs = candidate_cautions(by_name.get(name, {}), cnt)[:3]
            stack = bye_stack_caution(by_name.get(name, {}).get("Bye"), mine_rows)
            if stack:
                msgs.append(stack)
            for c in msgs:
                st.toast(c, icon="⚠️")

        def _apply(name: str, target: str) -> None:
            act, who = resolve_action(owner_of(name), target)
            if act == "remove":
                remove_pick(name)
            elif act == "add":
                record_pick(name, who)
            else:  # swap owners cleanly
                remove_pick(name)
                record_pick(name, who)
            if act in ("add", "swap") and who == OWNER_ME:
                _queue_cautions(name)

        def _click(state_key: str, target: str) -> None:
            c = st.session_state.get(state_key)
            names = st.session_state.get("_board_names", [])
            if c is not None and 0 <= c["row"] < len(names):
                _apply(names[c["row"]], target)

        def _do_draft() -> None:
            _click("draft_click", OWNER_ME)

        def _do_gone() -> None:
            _click("gone_click", OWNER_OTHER)

        def _do_view() -> None:
            c = st.session_state.get("view_click")
            names = st.session_state.get("_board_names", [])
            if c is not None and 0 <= c["row"] < len(names):
                st.session_state["_focus_name"] = names[c["row"]]

        st.markdown("#### Board")
        pv, sh = st.columns([1, 1])
        pool_view = pv.radio(
            "Pool view", ["Available only", "Show all (ghosted)"],
            horizontal=True,
            help="'Show all' keeps drafted players visible, dimmed, with live "
                 "toggle buttons so a mis-click is fixed right on the row.")
        ghosted = pool_view.startswith("Show all")
        sos_horizon = sh.radio(
            "SoS horizon", ["Early (Wk 1-4)", "Full season"], horizontal=True,
            help="Early = each team's real Weeks 1-4 opponents (from the kit). "
                 "Full season = FantasyPros' position-split full-year rating. "
                 "Same easiest→hardest label scale for both.")
        full_sos = sos_horizon.startswith("Full")
        sos_label_col = "FullSoS_Label" if full_sos else "SoS_Label"
        sos_rank_col = "FullSoS_Rank" if full_sos else "SoS_Rank"
        sos_head = f"SoS · {'Full' if full_sos else 'Wk1-4'}"

        f1, f2, f3, f4 = st.columns([2.2, 1.5, 1.5, 1.8])
        search = f1.text_input("Search name / team", "").strip().lower()
        pos_filter = f2.multiselect("Position", ["QB", "RB", "WR", "TE"], default=[])
        bye_options = sorted(int(b) for b in players["Bye"].dropna().unique())
        bye_filter = f3.multiselect("Bye week", bye_options, default=[],
                                    disabled=not bye_known)
        sos_options = [x for x in ["Very Soft", "Soft", "Neutral", "Tough",
                                   "Gauntlet", "N/A (FA)"]
                       if x in set(players[sos_label_col])]
        sos_filter = f4.multiselect(f"{sos_head} tier", sos_options, default=[])

        g1, g2, g3, g4 = st.columns(4)
        cy_only = g1.checkbox("Contract year only")
        hide_no_adp = g2.checkbox("Hide no-ADP")
        high_off_only = g3.checkbox("⚡ High-offense only",
                                    disabled=not players["HighOffense"].any())
        monitor_only = g4.checkbox("🚑 Injury-monitored only",
                                   disabled=not players["Monitor"].any())

        s1, s2, s3 = st.columns([2, 1, 3])
        sort_by = s1.selectbox(
            "Sort by",
            ["GlobalValue", "GlobalRank", "ADP", "ValueDelta", "ECR_Pos", "Tier",
             "Bye", "SoS", "Name"],
            index=0)
        sort_dir = s2.radio("Order", ["Desc", "Asc"], horizontal=True,
                            index=0 if sort_by == "GlobalValue" else 1)
        max_tier = int(players["Tier"].max())
        tier_max = s3.slider("Max NFL tier to show", 1, max_tier, max_tier)

        view = players.copy()
        if not ghosted:
            view = view[~view["Name"].isin(drafted_names)]
        if search:
            view = view[view["Name"].str.lower().str.contains(search, na=False) |
                        view["Team"].str.lower().str.contains(search, na=False)]
        if pos_filter:
            view = view[view["Pos"].isin(pos_filter)]
        if bye_filter:
            view = view[view["Bye"].isin(bye_filter)]
        if sos_filter:
            view = view[view[sos_label_col].isin(sos_filter)]
        if cy_only:
            view = view[view["ContractYear"] == "Y"]
        if hide_no_adp:
            view = view[view["ADP"].notna()]
        if high_off_only:
            view = view[view["HighOffense"]]
        if monitor_only:
            view = view[view["Monitor"]]
        view = view[view["Tier"] <= tier_max]

        ascending = (sort_dir == "Asc")
        # "SoS" sorts by the selected horizon's rank (1 = easiest schedule).
        sort_key = sos_rank_col if sort_by == "SoS" else sort_by
        # NaN always sinks to the bottom, whichever direction we sort.
        view = view.sort_values(sort_key, ascending=ascending, na_position="last",
                                kind="mergesort").reset_index(drop=True)

        display_names = view["Name"].tolist()
        st.session_state["_board_names"] = display_names
        owners = [owner_of(n) for n in display_names]

        # Snake-turn awareness: your next two picks, and per-row "can he
        # slide back to me?" risk ("" / "fringe" / "gone").
        next_pick, pick_after = upcoming_my_picks(current_overall, my_slots)
        adp_risks = [adp_pick_risk(a, next_pick, pick_after)
                     for a in view["ADP"].tolist()]
        _RISK_ICON = {"": "", "fringe": "⏳", "gone": "🔥"}

        def _labels(o):
            if o == OWNER_ME:
                return ("✔ Mine", "Gone")     # click Mine=undo, Gone=swap
            if o == OWNER_OTHER:
                return ("Draft", "✔ Gone")    # click Gone=undo, Draft=swap
            return ("Draft", "Gone")

        def _risk(r) -> str:
            s = ""
            vd = r["ValueDelta"]
            if pd.notna(vd) and float(vd) <= -10:
                s += "⚠"
            if _clean(r.get("DurabilityNote")):
                s += "🩹"
            return s

        def _tags(r) -> str:
            """Curated-context glyphs: 🔒 handcuff · 🚑 injury-monitored · ⚡ high-offense."""
            s = ""
            if str(r.get("Flag", "")).strip().upper() == "HANDCUFF" or _clean(r.get("HandcuffFor")):
                s += "🔒"
            if bool(r.get("Monitor")):
                s += "🚑"
            if bool(r.get("HighOffense")):
                s += "⚡"
            return s

        show = pd.DataFrame({
            "View": ["🔍"] * len(view),
            "Draft": [_labels(o)[0] for o in owners],
            "Gone": [_labels(o)[1] for o in owners],
            "G.Rk": view["GlobalRank"],
            "Name": view["Name"].fillna("").astype(str),
            "Pos": view["Pos"].fillna("").astype(str),
            "Team": view["Team"].fillna("").astype(str),
            "Bye": view["Bye"],
            "Pos ECR": view["ECR_Pos"],
            "NFL Tier": view["Tier"],
            "ADP": view["ADP"],
            "Value Δ": view["ValueDelta"],
            "Risk": [(_risk(r) + _RISK_ICON[rk])
                     for (_, r), rk in zip(view.iterrows(), adp_risks)],
            "Tags": [_tags(r) for _, r in view.iterrows()],
            "CY": view["ContractYear"].map({"Y": "CY", "N": ""}).fillna(""),
            "Flag": view["Flag"].fillna("").astype(str),
            "SoS": view[sos_label_col].fillna("").astype(str),
            "Notes": [effective_note(r) for _, r in view.iterrows()],
        })
        if ghosted:
            show.insert(7, "Status", [
                "— MINE —" if o == OWNER_ME else ("— GONE —" if o == OWNER_OTHER else "")
                for o in owners])

        def _num(v, plus=False):
            if pd.isna(v):
                return "—"
            return f"{v:+.1f}" if plus else f"{v:.1f}"

        def _dim(row):
            drafted = ghosted and bool(row.get("Status", ""))
            css = "color:#8a8a8a; background-color:transparent" if drafted else ""
            return [css] * len(row)

        _adp_col = list(show.columns).index("ADP")
        _ADP_CSS = {"gone": "background-color:#9a3412; color:#ffffff",
                    "fringe": "background-color:#854d0e; color:#ffffff"}

        def _adp_bg_row(row):
            styles = [""] * len(row)
            i = row.name
            rk = adp_risks[i] if isinstance(i, int) and 0 <= i < len(adp_risks) else ""
            if rk in _ADP_CSS:
                styles[_adp_col] = _ADP_CSS[rk]
            return styles

        styler = (show.style
                  .map(_vd_bg, subset=["Value Δ"])
                  .map(_flag_bg, subset=["Flag"])
                  .apply(_adp_bg_row, axis=1)
                  .format({"ADP": lambda v: _num(v),
                           "Value Δ": lambda v: _num(v, plus=True),
                           "Bye": lambda v: "—" if pd.isna(v) else f"{int(v)}"}))
        if ghosted:
            styler = styler.apply(_dim, axis=1)

        cfg = {
            "View": st.column_config.ButtonColumn(
                "", key="view_click", on_click=_do_view, width="small",
                help="Open this player's deep-dive drawer below the table"),
            "Draft": st.column_config.ButtonColumn(
                "", key="draft_click", on_click=_do_draft, width=96),
            "Gone": st.column_config.ButtonColumn(
                "", key="gone_click", on_click=_do_gone, width=96),
            "G.Rk": st.column_config.Column(
                "G.Rk", width="small",
                help="Global Rank — the model's overall value-based rank across "
                     "ALL positions (1 = best pick available regardless of "
                     "position). Sorted on Global Value by default."),
            "Pos": st.column_config.Column("Pos", width="small",
                                           help="Position: QB / RB / WR / TE"),
            "Team": st.column_config.Column("Team", width="small",
                                            help="NFL team ('FA' = free agent)"),
            "Bye": st.column_config.Column(
                "Bye", width="small", help="2026 NFL bye week"),
            "Pos ECR": st.column_config.Column(
                "Pos ECR", width="small",
                help="Positional Expert Consensus Rank — the market's consensus "
                     "rank among players AT THIS POSITION (FantasyPros, ~100+ "
                     "experts). Lower = more highly regarded."),
            "NFL Tier": st.column_config.Column(
                "NFL Tier", width="small",
                help="Global value tier (1–12). Tiers are natural scoring "
                     "cliffs cut from gaps in Global Value — players in one tier "
                     "are roughly interchangeable; the drop BETWEEN tiers is "
                     "where value is lost."),
            "ADP": st.column_config.Column(
                "ADP", width="small",
                help="Average Draft Position (12-team Superflex, FantasyPros). "
                     "Amber cell = at real risk of being sniped before your next "
                     "turn (scaled to our 10-team room)."),
            "Value Δ": st.column_config.Column(
                "Value Δ", width="small",
                help="Value Delta = ADP − Global Rank.  POSITIVE (green): the "
                     "market lets him slide PAST the model's rank — a draft-day "
                     "discount.  NEGATIVE (red): the market drafts him AHEAD of "
                     "the model — a reach."),
            "CY": st.column_config.Column(
                "CY", width="small",
                help="Contract Year — 'CY' if the player is in the final year of "
                     "his deal (common usage/motivation bump; also a trade / "
                     "hold-out risk)."),
            "Risk": st.column_config.TextColumn(
                "⚑", width="small",
                help="Computed risk:  ⚠ reach vs ADP (Δ ≤ -10)  ·  🩹 durability "
                     "note  ·  🔥 likely gone before your next pick  ·  ⏳ on the "
                     "bubble between your next two picks"),
            "Tags": st.column_config.TextColumn(
                "Tags", width="small",
                help="Curated context:  🔒 handcuff  ·  🚑 injury-monitored "
                     "(draft-day news)  ·  ⚡ elite offensive environment"),
            "Flag": st.column_config.Column(
                "Flag", width="small",
                help="Model conviction / risk tag — STUD (elite, buy) · FLOOR "
                     "(safe floor) · SLEEPER (late upside) · GAMBLE (boom/bust) "
                     "· BUST (priced above the model — fade) · HANDCUFF (RB "
                     "insurance) · DEPTH (bench filler)."),
            "SoS": st.column_config.TextColumn(
                sos_head, width="small",
                help=("Weeks 1-4 opponents (from the kit)" if not full_sos
                      else "FantasyPros full-season, position-split")
                     + " — Very Soft (easiest) → Gauntlet (hardest). "
                       "Switch with the SoS horizon toggle."),
            "Notes": st.column_config.TextColumn(
                "Notes", width="large",
                help="🆕 = a draft_day_news.py override (fresh intel); the "
                     "researched note is kept behind it."),
        }
        st.dataframe(
            styler, hide_index=True, use_container_width=True, height=560,
            column_config=cfg, key="board_table")
        st.caption(f"{len(show)} shown · {len(drafted_names)} drafted · "
                   f"{len(players) - len(drafted_names)} still available  ·  "
                   f"tap 🔍 for a player's deep-dive")

        with st.expander("ℹ️  What the columns & tags mean"):
            st.markdown(
                "| Column | Meaning |\n|---|---|\n"
                "| **G.Rk** | Global Rank — model's overall value rank across all "
                "positions (1 = best pick, any position) |\n"
                "| **Pos ECR** | Positional Expert Consensus Rank — market's "
                "consensus rank *within* the position (lower = better) |\n"
                "| **NFL Tier** | Global value tier 1–12; the gap *between* tiers "
                "is a scoring cliff |\n"
                "| **ADP** | Average Draft Position (12-team Superflex) |\n"
                "| **Value Δ** | ADP − Global Rank. **+green** = slides past the "
                "model → discount. **−red** = drafted ahead of the model → reach |\n"
                "| **CY** | Contract Year — final year of the player's deal |\n"
                "| **Flag** | STUD / FLOOR / SLEEPER / GAMBLE / BUST / HANDCUFF / "
                "DEPTH — model conviction & risk |\n"
                "| **SoS** | schedule difficulty for the selected horizon — Very "
                "Soft (easiest) → Gauntlet (hardest) |\n\n"
                "**⚑ risk glyphs:** ⚠ reach · 🩹 durability note · 🔥 likely gone "
                "before your next pick · ⏳ on the bubble  \n"
                "**Tag glyphs:** 🔒 handcuff · 🚑 injury-monitored · ⚡ elite "
                "offense  \n"
                "**Sidebar:** ⛰️ = a value tier about to dry up before your next turn.")

        # ---- focus drawer: driven by the 🔍 View button ----------------
        focus_name = st.session_state.get("_focus_name")
        if focus_name and focus_name in by_name:
            fr = by_name[focus_name]
            with st.container(border=True):
                hc1, hc2 = st.columns([8, 1])
                hc1.markdown(
                    f"### {fr['Name']} — {fr['Pos']} · {fr['Team']}  "
                    f"<span style='color:#888'>NFL Tier {int(fr['Tier'])} · "
                    f"Global #{int(fr['GlobalRank'])} · Pos ECR {int(fr['ECR_Pos'])}"
                    f"</span>", unsafe_allow_html=True)
                hc2.button("✕", key="_focus_clear",
                           on_click=lambda: st.session_state.pop("_focus_name", None))

                _chips = flash_tags(fr)
                if _chips:
                    st.markdown(" ".join(f":{c}-badge[{t}]" for t, c in _chips))

                if _clean(fr.get("NewsNote")):
                    st.warning("🆕 " + _clean(fr["NewsNote"]))
                _hcf = _clean(fr.get("HandcuffFor"))
                if _hcf:
                    st.info(f"🔒 **Handcuff for {_hcf}** — direct contingency.")
                elif str(fr.get("Flag", "")).strip().upper() == "HANDCUFF":
                    st.info("🔒 **Handcuff** — see the note for who he backs up.")
                if bool(fr.get("HighOffense")):
                    st.info(f"⚡ **Elite offensive environment** ({fr['Team']}) — "
                            f"ceiling target, esp. rounds 11–16.")
                if bool(fr.get("Monitor")):
                    st.info("🚑 **Injury-monitored** — confirm his status the "
                            "morning of the draft.")

                d1, d2, d3, d4, d5 = st.columns(5)
                adp, vd = fr["ADP"], fr["ValueDelta"]
                d1.metric("ADP", f"{adp:.1f}" if pd.notna(adp) else "—")
                d2.metric("Value Δ (ADP − rank)",
                          f"{vd:+.1f}" if pd.notna(vd) else "—",
                          help="positive = market is letting him slide past the model")
                d3.metric("SoS · Wk 1-4", _clean(fr["SoS_Label"]) or "—",
                          help="each team's real Weeks 1-4 opponents (from the kit)")
                d4.metric("SoS · Full season", _clean(fr.get("FullSoS_Label")) or "—",
                          help="FantasyPros full-season, position-split")
                fbye = fr.get("Bye")
                d5.metric("Bye week", f"Wk {int(fbye)}" if pd.notna(fbye) else "—")

                frisk = adp_pick_risk(adp, next_pick, pick_after)
                if frisk == "gone" and pd.notna(adp):
                    st.warning(
                        f"🔥 **Likely gone before your next pick (#{next_pick}).** "
                        f"ADP {adp:.0f} ≈ pick {adp * ADP_TEAM_SCALE:.0f} in a "
                        f"10-team room — take him now or expect to lose him.")
                elif frisk == "fringe" and pd.notna(adp):
                    st.info(
                        f"⏳ **On the bubble.** ADP {adp:.0f} ≈ pick "
                        f"{adp * ADP_TEAM_SCALE:.0f} lands between your picks "
                        f"#{next_pick} and #{pick_after} — roughly a coin flip "
                        f"he makes it back.")

                bye_hits = [n for wk, names in roster_bye_counts(my_rows).items()
                            if pd.notna(fbye) and wk == int(fbye) for n in names]
                if bye_hits:
                    st.write(f"• **Bye overlap:** shares Week {int(fbye)} with "
                             f"{len(bye_hits)} of your roster ({', '.join(bye_hits)}).")

                cau = candidate_cautions(fr, my_counts)
                for c in (cau or ["No flags — clean pick at this roster spot."]):
                    st.write("• " + c)
                note = _clean(fr.get("Notes"))
                if note:
                    st.caption(note)
                dur = _clean(fr.get("DurabilityNote"))
                if dur:
                    st.caption("🩹 " + dur)

    # ------------------------------------------------------------------ ROSTER
    with tab_roster:
        st.markdown("#### Starting lineup (auto-optimized by GlobalValue)")
        slots, bench = assign_roster(my_rows)
        lineup_rows = []
        for label in STARTER_SLOTS:
            r = slots.get(label)
            rb = r.get("Bye") if r else None
            lineup_rows.append({
                "Slot": label,
                "Player": r["Name"] if r else "—",
                "Pos": r["Pos"] if r else "",
                "Team": r["Team"] if r else "",
                "Bye": int(rb) if rb is not None and pd.notna(rb) else "",
                "NFL Tier": int(r["Tier"]) if r and pd.notna(r["Tier"]) else "",
                "ADP": f"{r['ADP']:.1f}" if r and pd.notna(r["ADP"]) else "—",
                "Global Rank": int(r["GlobalRank"]) if r else "",
            })
        st.dataframe(pd.DataFrame(lineup_rows), hide_index=True, use_container_width=True)

        st.markdown(f"#### Bench ({len(bench)}/{BENCH_SLOTS})")
        if bench:
            bench_df = pd.DataFrame([{
                "Player": r["Name"], "Pos": r["Pos"], "Team": r["Team"],
                "Bye": int(r["Bye"]) if r.get("Bye") is not None and pd.notna(r.get("Bye")) else "",
                "NFL Tier": int(r["Tier"]) if pd.notna(r["Tier"]) else "",
                "ADP": f"{r['ADP']:.1f}" if pd.notna(r["ADP"]) else "—",
            } for r in bench])
            st.dataframe(bench_df, hide_index=True, use_container_width=True)
        else:
            st.caption("No bench players yet.")

        st.divider()
        exp_df = build_export_df(my_rows, picks)
        st.download_button(
            "Download roster CSV",
            data=exp_df.to_csv(index=False).encode("utf-8"),
            file_name=f"my_draft_roster_{datetime.now():%Y%m%d_%H%M}.csv",
            mime="text/csv",
        )
        st.download_button(
            "Download roster text summary",
            data=build_text_summary(my_rows, picks),
            file_name=f"my_draft_roster_{datetime.now():%Y%m%d_%H%M}.txt",
            mime="text/plain",
        )

    # ------------------------------------------------------------------ HISTORY
    with tab_history:
        st.markdown("#### Every pick logged (mine and opponents')")
        if picks:
            hist_rows = []
            for pk in picks:
                rnd, pir = overall_to_round_pick(pk["overall"])
                r = by_name.get(pk["name"], {})
                hist_rows.append({
                    "Pick": pk["overall"],
                    "Rd": f"{rnd}.{pir:02d}",
                    "Owner": "MINE" if pk["owner"] == OWNER_ME else "other",
                    "Player": pk["name"],
                    "Pos": pk.get("pos") or r.get("Pos", ""),
                    "Team": pk.get("team") or r.get("Team", ""),
                    "Bye": int(r["Bye"]) if r and pd.notna(r.get("Bye")) else "",
                    "NFL Tier": int(r["Tier"]) if r and pd.notna(r.get("Tier")) else "",
                    "ADP": f"{r['ADP']:.1f}" if r and pd.notna(r.get("ADP")) else "—",
                })
            st.dataframe(pd.DataFrame(hist_rows), hide_index=True, use_container_width=True)

            st.markdown("##### Fix a mistake")
            m1, m2 = st.columns([3, 1])
            bad = m1.selectbox("Remove a specific pick (wrong entry)",
                               options=[p["name"] for p in picks], index=None,
                               placeholder="pick a player to un-draft...")
            m2.button("Remove", disabled=not bad,
                      on_click=lambda: remove_pick(bad) if bad else None)
        else:
            st.caption("No picks yet. Use the Quick draft bar on the Available board tab.")


def build_export_df(my_rows: list[dict], picks: list[dict]) -> pd.DataFrame:
    """16-row roster export: slot assignment + when you drafted each guy."""
    slots, bench = assign_roster(my_rows)
    drafted_overall = {p["name"]: p["overall"] for p in picks if p["owner"] == OWNER_ME}
    rows = []

    def add(slot_label: str, r: dict | None) -> None:
        if r is None:
            rows.append({"Slot": slot_label, "Player": "", "Pos": "", "Team": "",
                         "Bye": "", "NFL_Tier": "", "ADP": "", "ValueDelta": "",
                         "GlobalRank": "", "DraftedOverall": "", "Round": ""})
            return
        ov = drafted_overall.get(r["Name"])
        rnd = overall_to_round_pick(ov)[0] if ov else ""
        bye = r.get("Bye")
        rows.append({
            "Slot": slot_label,
            "Player": r["Name"],
            "Pos": r["Pos"],
            "Team": r["Team"],
            "Bye": int(bye) if bye is not None and pd.notna(bye) else "",
            "NFL_Tier": int(r["Tier"]) if pd.notna(r["Tier"]) else "",
            "ADP": f"{r['ADP']:.1f}" if pd.notna(r["ADP"]) else "",
            "ValueDelta": f"{r['ValueDelta']:+.1f}" if pd.notna(r["ValueDelta"]) else "",
            "GlobalRank": int(r["GlobalRank"]),
            "DraftedOverall": ov or "",
            "Round": rnd,
        })

    for label in STARTER_SLOTS:
        add(label, slots.get(label))
    for i in range(BENCH_SLOTS):
        add(f"BN{i + 1}", bench[i] if i < len(bench) else None)
    return pd.DataFrame(rows)


def build_text_summary(my_rows: list[dict], picks: list[dict]) -> str:
    df = build_export_df(my_rows, picks)
    counts = position_counts(my_rows)
    byes = roster_bye_counts(my_rows)
    bye_line = " / ".join(f"Wk{wk}: {len(n)}" for wk, n in sorted(byes.items())) or "none"
    lines = [
        "2026 FANTASY DRAFT -- MY ROSTER",
        f"Generated {datetime.now():%Y-%m-%d %H:%M}",
        f"League: {N_TEAMS}-team, {N_ROUNDS} rounds, pick {MY_DRAFT_SLOT}, Superflex/TE-premium",
        f"Position mix: QB {counts['QB']} / RB {counts['RB']} / WR {counts['WR']} / TE {counts['TE']}",
        f"Bye spread: {bye_line}",
        "",
        f"{'SLOT':<10} {'PLAYER':<26} {'POS':<4} {'TM':<4} {'BYE':>4} {'TIER':<5} {'ADP':>6} {'RD':>3}",
        "-" * 70,
    ]
    for _, r in df.iterrows():
        lines.append(f"{r['Slot']:<10} {str(r['Player']):<26} {str(r['Pos']):<4} "
                     f"{str(r['Team']):<4} {str(r['Bye']):>4} {str(r['NFL_Tier']):<5} "
                     f"{str(r['ADP']):>6} {str(r['Round']):>3}")
    return "\n".join(lines) + "\n"


if __name__ == "__main__":
    run_app()

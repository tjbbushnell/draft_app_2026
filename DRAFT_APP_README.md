# 2026 Draft Companion (`draft_app.py`)

Live draft-room assistant. Reads `players_2026.csv` (the finished board from
`build_draft_kit_2026.py`), `FantasyPros_Fantasy_Football_Bye_Week_Cheatsheet.csv`
for 2026 bye weeks, and **`draft_day_news.py`** for hand-entered draft-day intel
(injury updates, handcuff linkages, high-offense teams). It manages drafted state,
roster slots, guardrails, and export. It does **not** recompute scoring/VBD/tiers/ADP
— those columns are consumed exactly as the kit builder produced them, and
`players_2026.csv` is never modified.

### `draft_day_news.py` — the manual overlay

Edit this file right up to (and during) the draft, then hit **Rerun** in the app —
it's re-read every run, no cache clear needed. It holds:

- `NEWS` — `{player name: {"note": "...", "monitor": True, "handcuff_for": "Starter"}}`.
  The fresh note shows `🆕`-prefixed on the board and as its own line in the drawer;
  the researched note is kept behind it. `monitor` → `🚑` badge. `handcuff_for` →
  `🔒` badge + a linkage line (use it for handcuffs whose base `Flag` is something
  else, e.g. Jonathon Brooks stays flagged `BUST` **and** gets `🔒` for Hubbard).
- `HIGH_OFFENSE_TEAMS` — team codes you rate as elite scoring environments. **There
  is no implied-total data in the project**; this is your call, seeded from a
  heuristic (most top-80 skill players). Drives the `⚡` tag and the
  "⚡ high-offense only" filter.

A typo'd name in `NEWS` shows a sidebar warning (and fails `test_draft_app.py`).

**Strength of schedule — two horizons.** The `SoS horizon` toggle (above the
table) switches the `SoS` column, the `SoS` filter, the `SoS` sort option, and
the label in the focus drawer between:

- **Early (Wk 1–4)** — each team's real Weeks 1–4 opponents, from the kit
  (`players_2026.csv` `SoS_Label`).
- **Full season** — FantasyPros' position-split (QB/RB/WR/TE) full-year rating
  from `FantasyPros_Fantasy_Football_2026_Strength_Of_Schedule.csv`, mapped by
  team name → code and put on the **same** `Very Soft → Gauntlet` label scale.
  FA players show `N/A (FA)`.

The focus drawer shows **both** horizons side by side. Still *not* built: a
playoff-weeks (14–17) SoS — the pipeline has no Week 14–17 matchups to derive it
from, so it stays deferred rather than faked.

## Launch

```bash
pip install streamlit pandas
streamlit run draft_app.py
```

On this machine the interpreter is the `py` launcher and `streamlit` is already
installed, so:

```bash
py -m streamlit run draft_app.py
```

It opens at http://localhost:8501. Keep the terminal window open during the draft.

## Using it during the draft

- **Inline actions**: every row has three buttons — **`🔍`** (open the deep-dive
  drawer under the table), **Draft** (claim for your team — fills a roster slot,
  fires guardrails, pops a caution toast), **Gone** (opponent took him — off the
  board, no roster effect). No selection checkbox; `🔍` is the only non-toggle
  control. (Streamlit can't wire a plain text cell to a callback, so the drawer
  opens from the `🔍` button, not a name-click.)
- **Toggle / swap** (misclick recovery): press the **same** button again to
  un-draft (row shows `✔ Mine` / `✔ Gone` once claimed). Press the **other**
  button to reassign cleanly — e.g. a player you marked *Gone* who actually fell
  to you: just hit *Draft* on his row.
- **Pool view** (radio above the table):
  - **Available only** — drafted players vanish; clean view of what's left.
  - **Show all (ghosted)** — drafted players stay, dimmed, with a `— MINE —` /
    `— GONE —` status and live toggle buttons so you can fix mistakes in place.
- **Focus drawer**: hit `🔍` on a row to open a panel under the table —
  plain-language **flash-tag chips** (Contract Year · Hamstring Risk · Elite
  Offense · Model Value · …, colour-coded, readable in ~2s), then `🆕`
  draft-day news, `🔒` handcuff linkage, `⚡` offensive-environment note, `🚑`
  monitor flag, metrics (ADP / Value Δ / SoS Wk 1-4 / SoS full / bye week),
  snipe-risk, bye overlap with your roster, cautions, and the researched note.
  `✕` closes it.
- **Column legend**: an `ℹ️ What the columns & tags mean` expander above the
  table, plus a `help` tooltip on **every** column header (G.Rk, Pos ECR, NFL
  Tier, ADP, Value Δ, CY, Flag, SoS, ⚑, Tags).
- **Tier-cliff warnings** (sidebar): `⛰️ Global Tier N — only X left` when a
  value tier is about to dry up before your next turn, plus a positional
  `last-in-tier` line (`RB tier 1: 1 left`) so a run doesn't catch you asleep.
- **Bye weeks**: the `Bye` column (2026 NFL bye) is on the master table, in the
  focus drawer, in *My roster* / *Draft history*, and in both exports. It's a
  **filter** (`Bye week` multiselect) and a **Sort by** option. Byes come from the
  FantasyPros cheatsheet, mapped team→bye by cross-reference; if that file is
  missing the column just shows `—` and everything else still works.
- **"Likely gone before your next pick"**: the `ADP` cell is tinted for
  un-drafted players who probably won't make it back to you — **burnt-orange** =
  gone before your very next pick, **amber** = on the bubble (ADP lands between
  your next two picks). The `⚑` column mirrors this with `🔥` / `⏳`. ADP is a
  12-team tool, so pick position is scaled to our 10-team room (× 10/12) before
  the call. The focus drawer spells it out ("ADP 6 ≈ pick 5 in a 10-team room —
  take him now"). Sidebar shows a running "`🔥 ~N` available project gone by your
  pick #X".
- **Filters / sort**: search (name or team), position, bye week, contract-year,
  hide-no-ADP, **`⚡` high-offense only**, **`🚑` injury-monitored only**,
  Sort-by + Order, max-tier slider. `Value Δ` green→red; `Flag` colour-coded
  (`HANDCUFF` = slate pill). The `⚑` column = computed risk (`⚠` reach, `🩹`
  durability, `🔥`/`⏳` snipe). The `Tags` column = curated context (`🔒` handcuff,
  `🚑` monitored, `⚡` high-offense). No-ADP / no-bye players sort to the bottom.
  Column-header re-sort is fine — the buttons still map to the right player.
- **Sidebar** (always visible): snake pick clock (your picks: 8, 13, 28, 33, …),
  live roster by slot (QB / RB1-2 / WR1-2 / TE / FLEX1-2 / SUPERFLEX / 7 bench,
  filled greedy-by-GlobalValue), your **bye spread** line, a snipe-risk count,
  guardrails (3rd-TE bloat, Superflex-QB pacing, late starter holes, **bye
  stacking** — a warning at 3 starters sharing a bye, an error at 4), and
  scarcity (remaining by position / tier, SF-caliber QB and elite TE-premium
  TE left).
- **My roster** tab: optimized starting lineup + bench, and the export buttons.
- **Draft history** tab: every logged pick (mine + opponents), plus *Fix a mistake*
  to remove a wrongly entered pick.

## Safety net

Every action autosaves to `draft_state.json` in this folder. If the browser tab
crashes or you refresh, just reload the page — it picks up exactly where you were.
**Undo last pick** and a confirm-gated **Reset entire draft** are in the sidebar.
Delete `draft_state.json` to start a fresh draft from scratch.

## Export

*My roster* tab → **Download roster CSV** (16 rows: slot, player, team, bye,
tier, ADP, Value Δ, when you drafted him) or **Download roster text summary**
(plain-text lineup card with a bye-spread line).

## Tests

Pure-logic helpers (snake math, slot assignment, guardrails, bye parsing,
snipe-risk, news overlay, full-season SoS parsing, tier cliffs, flash tags,
export) are covered — **39 tests**, incl. real-file checks that every
`draft_day_news.NEWS` name exists in `players_2026.csv` and that the full-season
SoS covers all 32 teams:

```bash
py test_draft_app.py
```

#!/usr/bin/env python3
"""
Mock draft simulator for the 2026 Fantasy Draft Kit.

Uses players_2026.csv (the real ECR + ADP + Global VBD dataset from
build_draft_kit_2026.py v5) to run a full 10-team, 16-round snake draft,
with the user picking 8th overall.

- Opponents (9 AI teams) draft off REAL ADP, with:
    * a roster-need pacing bonus (steers them toward a realistic final
      shape rather than pure best-ADP regardless of position),
    * a "positional run" bonus (if 3+ of the last 5 picks league-wide were
      the same position, that position gets a temporary bump -- this is
      what creates RB dumps, QB run panics in a Superflex format, etc.),
    * small random noise (real drafters don't follow ADP to the decimal).
- The user's team drafts off our own GLOBAL VALUE model (the Draft Kit's
  actual recommendation engine), with the same need-pacing logic and a
  reduced (not zero) run-reaction, and effectively zero random noise --
  i.e. "what would a savvy owner following our kit actually do here."

Both sides obey a hard floor: if a team's remaining picks can no longer
cover its minimum starter requirements (1 QB / 2 RB / 2 WR / 1 TE), the
next pick is forced to that position.

Run with a different seed to get a different simulated draft (different
opponent behavior -> different players fall to pick 8, 13, 28, ...).
"""

import sys
import numpy as np
import pandas as pd

df = pd.read_csv("players_2026.csv")
assert df["Name"].is_unique, "Duplicate player names in players_2026.csv!"

# Effective ADP used for drafting logic: real ADP where we have it; for the
# ~85 players with no published ADP (deepest end of the 328-player pool),
# push them below the entire real-ADP range but keep them in our model's
# own GlobalRank order relative to each other.
df["EffADP"] = df["ADP"]
missing = df["EffADP"].isna()
df.loc[missing, "EffADP"] = 500 + df.loc[missing, "GlobalRank"]

N_TEAMS = 10
N_ROUNDS = 16
USER_TEAM = 8  # 1-indexed -- pick 8 of 10, matches the real league

TARGET_COUNTS = {"QB": 3, "RB": 5, "WR": 6, "TE": 2}   # soft pacing target, sums to 16
MIN_STARTERS = {"QB": 1, "RB": 2, "WR": 2, "TE": 1}     # hard floor, enforced late

NEED_SCALE, RUN_SCALE, NOISE_STD = 22.0, 16.0, 9.0
USER_NEED_SCALE, USER_RUN_SCALE, USER_NOISE_STD = 20.0, 8.0, 0.0
USER_ADP_WEIGHT = 0.15  # how much market ADP still matters as a sanity check on our own value model

FLEX_ELIGIBLE = {"RB", "WR", "TE"}
SF_ELIGIBLE = {"QB", "RB", "WR", "TE"}


def build_pick_order(n_teams, n_rounds):
    order = []
    for rnd in range(1, n_rounds + 1):
        teams = list(range(1, n_teams + 1))
        if rnd % 2 == 0:
            teams = list(reversed(teams))
        for slot, team in enumerate(teams, start=1):
            overall = (rnd - 1) * n_teams + slot
            order.append((rnd, overall, team))
    return order


def run_simulation(seed):
    rng = np.random.default_rng(seed)
    pool = df.set_index("Name", drop=False)
    available = set(pool.index)
    rosters = {t: [] for t in range(1, N_TEAMS + 1)}          # list of (round, overall, player row)
    pos_counts = {t: {"QB": 0, "RB": 0, "WR": 0, "TE": 0} for t in range(1, N_TEAMS + 1)}
    recent_positions = []

    for rnd, overall, team in build_pick_order(N_TEAMS, N_ROUNDS):
        avail_df = pool.loc[list(available)]
        is_user = (team == USER_TEAM)
        remaining = N_ROUNDS - rnd + 1

        hard_needs = {p: max(0, MIN_STARTERS[p] - pos_counts[team][p]) for p in MIN_STARTERS}
        total_hard_need = sum(hard_needs.values())

        candidates = avail_df
        if total_hard_need > 0 and total_hard_need >= remaining:
            forced_pos = max(hard_needs, key=lambda p: hard_needs[p])
            forced_candidates = avail_df[avail_df["Pos"] == forced_pos]
            if not forced_candidates.empty:
                candidates = forced_candidates

        if is_user:
            base = -candidates["EffADP"] * USER_ADP_WEIGHT + candidates["GlobalValue"]
            need_scale, run_scale, noise_std = USER_NEED_SCALE, USER_RUN_SCALE, USER_NOISE_STD
        else:
            base = -candidates["EffADP"]
            need_scale, run_scale, noise_std = NEED_SCALE, RUN_SCALE, NOISE_STD

        pace_target = {p: TARGET_COUNTS[p] * (rnd / N_ROUNDS) for p in TARGET_COUNTS}
        need_bonus = candidates["Pos"].map(lambda p: pace_target[p] - pos_counts[team][p])

        window = recent_positions[-5:]
        run_bonus = candidates["Pos"].map(
            lambda p: (window.count(p) / max(1, len(window))) if window.count(p) >= 3 else 0.0
        )

        noise = rng.normal(0, noise_std, size=len(candidates)) if noise_std > 0 else np.zeros(len(candidates))

        score = base + need_bonus * need_scale + run_bonus * run_scale + noise
        pick_name = score.idxmax()
        pick_row = pool.loc[pick_name]

        rosters[team].append((rnd, overall, pick_row))
        pos_counts[team][pick_row["Pos"]] += 1
        available.discard(pick_name)
        recent_positions.append(pick_row["Pos"])

    return rosters


def assign_slots(user_picks):
    players = sorted([p for _, _, p in user_picks], key=lambda r: r["GlobalValue"], reverse=True)
    remaining = list(players)
    slots = {}

    def take_best(pos_set, label):
        for i, p in enumerate(remaining):
            if p["Pos"] in pos_set:
                slots[label] = remaining.pop(i)
                return
        slots[label] = None

    take_best({"QB"}, "QB")
    take_best({"RB"}, "RB1")
    take_best({"RB"}, "RB2")
    take_best({"WR"}, "WR1")
    take_best({"WR"}, "WR2")
    take_best({"TE"}, "TE")
    take_best(FLEX_ELIGIBLE, "FLEX1")
    take_best(FLEX_ELIGIBLE, "FLEX2")
    take_best(SF_ELIGIBLE, "SUPERFLEX")
    for i, p in enumerate(remaining, start=1):
        slots[f"BN{i}"] = p
    return slots


def fmt_player(p):
    return f"{p['Name']} ({p['Team']})"


def print_simulation(seed, sim_num):
    rosters = run_simulation(seed)
    user_picks = rosters[USER_TEAM]

    print(f"\n{'='*78}\nSIMULATION {sim_num}  (seed={seed})\n{'='*78}")
    print(f"{'Rd':>3} {'Pick':>5} {'Pos':<4} {'Player':<24} {'Team':<5} {'ADP':>6} {'Tier':>5} {'GVal':>7}  Note")
    for rnd, overall, p in user_picks:
        adp = f"{p['ADP']:.1f}" if pd.notna(p["ADP"]) else "--"
        note = ""
        if pd.notna(p["ADP"]):
            delta = p["ADP"] - overall
            if delta >= 10:
                note = f"SLIDE (ADP {p['ADP']:.0f} vs pick {overall})"
            elif delta <= -10:
                note = f"REACH (ADP {p['ADP']:.0f} vs pick {overall})"
        print(f"{rnd:>3} {overall:>5} {p['Pos']:<4} {p['Name']:<24} {p['Team']:<5} {adp:>6} "
              f"{p['TierLevel']:>5} {p['GlobalValue']:>7.1f}  {note}")

    counts = {"QB": 0, "RB": 0, "WR": 0, "TE": 0}
    for _, _, p in user_picks:
        counts[p["Pos"]] += 1
    print(f"\nPosition mix: QB {counts['QB']} / RB {counts['RB']} / WR {counts['WR']} / TE {counts['TE']}"
          f"  (total {sum(counts.values())})")

    slots = assign_slots(user_picks)
    print("\nProjected Week 1 lineup + bench:")
    for label in ["QB", "RB1", "RB2", "WR1", "WR2", "TE", "FLEX1", "FLEX2", "SUPERFLEX"]:
        p = slots[label]
        print(f"  {label:<10} {fmt_player(p) if p is not None else '-- (none drafted)'}")
    bench_labels = sorted([l for l in slots if l.startswith("BN")], key=lambda x: int(x[2:]))
    for label in bench_labels:
        p = slots[label]
        print(f"  {label:<10} {fmt_player(p)} ({p['Pos']})")

    return rosters


if __name__ == "__main__":
    seeds = [2026, 4242, 777]
    for i, s in enumerate(seeds, start=1):
        print_simulation(s, i)

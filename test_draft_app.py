#!/usr/bin/env python3
"""Pure-logic tests for draft_app.py -- no Streamlit needed.

    python test_draft_app.py      # plain asserts, exits non-zero on failure
    python -m pytest test_draft_app.py -q   # also works if pytest is installed
"""

import pandas as pd

from pathlib import Path

from draft_app import (
    ADP_TEAM_SCALE,
    BENCH_SLOTS,
    STARTER_SLOTS,
    _clean,
    _FLAG_BG,
    _flag_bg,
    adp_pick_risk,
    apply_news_overlay,
    assign_roster,
    build_export_df,
    build_text_summary,
    bye_stack_caution,
    candidate_cautions,
    effective_note,
    full_sos_label,
    next_my_pick,
    overall_to_round_pick,
    parse_full_sos,
    parse_team_byes,
    position_counts,
    resolve_action,
    roster_bye_counts,
    roster_guardrails,
    scarcity,
    snake_pick_slots,
    upcoming_my_picks,
)


def _p(name, pos, gv, **extra):
    row = {"Name": name, "Pos": pos, "Team": extra.get("Team", "XXX"),
           "GlobalValue": gv, "GlobalRank": extra.get("GlobalRank", 100),
           "ValueDelta": extra.get("ValueDelta", float("nan")),
           "ADP": extra.get("ADP", float("nan")),
           "Tier": extra.get("Tier", 5),
           "PosTier": extra.get("PosTier", 5),
           "Flag": extra.get("Flag", ""),
           "Notes": extra.get("Notes", ""),
           "DurabilityNote": extra.get("DurabilityNote", ""),
           "Bye": extra.get("Bye", None)}
    return row


def test_snake_pick_slots_slot8():
    slots = snake_pick_slots(draft_slot=8, n_teams=10, n_rounds=16)
    assert slots[:6] == [8, 13, 28, 33, 48, 53], slots
    assert len(slots) == 16
    # round parity: odd rounds go low->high, even rounds high->low
    assert slots[0] == 8 and slots[1] == 13


def test_snake_pick_slots_edges():
    assert snake_pick_slots(1, 10, 2) == [1, 20]
    assert snake_pick_slots(10, 10, 2) == [10, 11]


def test_overall_to_round_pick():
    assert overall_to_round_pick(1) == (1, 1)
    assert overall_to_round_pick(10) == (1, 10)
    assert overall_to_round_pick(11) == (2, 1)
    assert overall_to_round_pick(28) == (3, 8)


def test_next_my_pick():
    my = snake_pick_slots()
    assert next_my_pick(1, my) == (8, 7)
    assert next_my_pick(8, my) == (8, 0)          # on the clock
    assert next_my_pick(9, my) == (13, 4)
    assert next_my_pick(200, my) == (None, 0)     # draft over


def test_assign_roster_fills_flex_and_superflex_by_value():
    rows = [
        _p("QB1", "QB", 90), _p("QB2", "QB", 40),
        _p("RB1", "RB", 80), _p("RB2", "RB", 70), _p("RB3", "RB", 60),
        _p("WR1", "WR", 75), _p("WR2", "WR", 65),
        _p("TE1", "TE", 50),
    ]
    slots, bench = assign_roster(rows)
    assert slots["QB"]["Name"] == "QB1"
    assert slots["RB1"]["Name"] == "RB1" and slots["RB2"]["Name"] == "RB2"
    assert slots["WR1"]["Name"] == "WR1" and slots["WR2"]["Name"] == "WR2"
    assert slots["TE"]["Name"] == "TE1"
    # best leftover flex-eligible is RB3 (60) then... only one flex-eligible left
    assert slots["FLEX1"]["Name"] == "RB3"
    assert slots["FLEX2"] is None
    # SUPERFLEX takes best remaining SF-eligible == QB2
    assert slots["SUPERFLEX"]["Name"] == "QB2"
    assert bench == []


def test_assign_roster_bench_overflow():
    rows = [_p(f"WR{i}", "WR", 100 - i) for i in range(12)]
    slots, bench = assign_roster(rows)
    assert slots["WR1"]["Name"] == "WR0"
    # 2 WR slots + 2 FLEX + 1 SUPERFLEX = 5 starters filled by WRs, 7 to bench
    assert len(bench) == 7


def test_position_counts():
    rows = [_p("a", "QB", 1), _p("b", "RB", 1), _p("c", "RB", 1), _p("d", "TE", 1)]
    assert position_counts(rows) == {"QB": 1, "RB": 2, "WR": 0, "TE": 1}


def test_guardrail_third_te_is_error():
    rows = [_p("TE1", "TE", 50), _p("TE2", "TE", 40), _p("TE3", "TE", 30)]
    levels = [lvl for lvl, _ in roster_guardrails(rows)]
    assert "error" in levels


def test_guardrail_qb_pacing_warns_when_behind_late():
    # 9 picks, zero QBs -> must warn about Superflex QB pacing
    rows = [_p(f"WR{i}", "WR", 50 - i) for i in range(9)]
    msgs = [m for lvl, m in roster_guardrails(rows) if lvl == "warning"]
    assert any("QB" in m for m in msgs), msgs


def test_guardrail_clean_roster_reports_success():
    rows = [_p("QB1", "QB", 90), _p("RB1", "RB", 80), _p("WR1", "WR", 75)]
    levels = [lvl for lvl, _ in roster_guardrails(rows)]
    assert levels == ["success"]


def test_candidate_caution_reach_and_value():
    reach = _p("Reacher", "WR", 50, ValueDelta=-15, ADP=20, GlobalRank=35)
    assert any(c.startswith("REACH") for c in candidate_cautions(reach, {"QB": 0, "TE": 0}))
    slide = _p("Slider", "RB", 50, ValueDelta=18, ADP=40, GlobalRank=22)
    assert any(c.startswith("VALUE") for c in candidate_cautions(slide, {"QB": 0, "TE": 0}))


def test_candidate_caution_third_qb_and_third_te():
    qb = _p("QBx", "QB", 30)
    assert any("QB #3" in c for c in candidate_cautions(qb, {"QB": 2, "TE": 0}))
    te = _p("TEx", "TE", 30)
    assert any("TE #3" in c for c in candidate_cautions(te, {"QB": 0, "TE": 2}))


def test_scarcity_counts_and_cutoffs():
    df = pd.DataFrame([
        _p("QBa", "QB", 90, PosTier=1, Tier=1),
        _p("QBb", "QB", 40, PosTier=8, Tier=9),
        _p("TEa", "TE", 50, PosTier=2, Tier=3),
        _p("TEb", "TE", 20, PosTier=6, Tier=10),
        _p("RBa", "RB", 70, PosTier=3, Tier=4),
    ])
    sc = scarcity(df)
    assert sc["by_pos"] == {"QB": 2, "RB": 1, "WR": 0, "TE": 2}
    assert sc["sf_qb_left"] == 1        # only QBa has PosTier <= 4
    assert sc["elite_te_left"] == 1     # only TEa has PosTier <= 2


def test_adp_nan_sorts_last_both_directions():
    """The refinement the user asked for: no-ADP players sink in asc AND desc."""
    df = pd.DataFrame([
        {"Name": "HasADP_hi", "ADP": 5.0},
        {"Name": "NoADP", "ADP": float("nan")},
        {"Name": "HasADP_lo", "ADP": 120.0},
    ])
    asc = df.sort_values("ADP", ascending=True, na_position="last")["Name"].tolist()
    desc = df.sort_values("ADP", ascending=False, na_position="last")["Name"].tolist()
    assert asc[-1] == "NoADP", asc
    assert desc[-1] == "NoADP", desc


def test_export_df_has_16_rows_and_slots_in_order():
    rows = [_p("QB1", "QB", 90, ADP=5.0, ValueDelta=2.0, GlobalRank=2, Tier=1),
            _p("RB1", "RB", 80, GlobalRank=5),
            _p("TE1", "TE", 70, GlobalRank=8)]
    picks = [{"name": "QB1", "owner": "ME", "overall": 8},
             {"name": "RB1", "owner": "ME", "overall": 13},
             {"name": "TE1", "owner": "ME", "overall": 28}]
    df = build_export_df(rows, picks)
    assert len(df) == len(STARTER_SLOTS) + BENCH_SLOTS == 16
    assert df["Slot"].tolist()[:3] == ["QB", "RB1", "RB2"]
    qb_row = df[df["Slot"] == "QB"].iloc[0]
    assert qb_row["Player"] == "QB1"
    assert qb_row["DraftedOverall"] == 8
    assert qb_row["Round"] == 1                 # overall 8 -> round 1


def test_text_summary_is_nonempty_and_lists_players():
    rows = [_p("QB1", "QB", 90), _p("WRa", "WR", 80)]
    picks = [{"name": "QB1", "owner": "ME", "overall": 8},
             {"name": "WRa", "owner": "ME", "overall": 13}]
    txt = build_text_summary(rows, picks)
    assert "MY ROSTER" in txt and "QB1" in txt and "WRa" in txt
    assert txt.count("\n") >= 16


def test_clean_kills_the_nan_truthiness_trap():
    assert _clean(float("nan")) == ""
    assert _clean(None) == ""
    assert _clean("nan") == ""
    assert _clean("  psoas strain  ") == "psoas strain"
    assert _clean(0) == "0"


def test_candidate_caution_no_false_durability_on_nan():
    row = _p("Clean", "WR", 50)
    row["DurabilityNote"] = float("nan")
    row["Notes"] = float("nan")
    row["Flag"] = float("nan")
    assert not any(c.startswith("DURABILITY") for c in
                   candidate_cautions(row, {"QB": 0, "TE": 0}))


def test_resolve_action_add_toggle_swap():
    # undrafted + either button -> claim for that side
    assert resolve_action(None, "ME") == ("add", "ME")
    assert resolve_action(None, "OTHER") == ("add", "OTHER")
    # same button re-pressed -> toggle off
    assert resolve_action("ME", "ME") == ("remove", None)
    assert resolve_action("OTHER", "OTHER") == ("remove", None)
    # opposite button -> clean swap
    assert resolve_action("OTHER", "ME") == ("swap", "ME")
    assert resolve_action("ME", "OTHER") == ("swap", "OTHER")


def test_parse_team_byes_from_synthetic_cheatsheet():
    import tempfile
    with tempfile.TemporaryDirectory() as td:
        f = Path(td) / "bye.csv"
        f.write_text(
            '"Week 6 Bye"\n'
            '"ECR","Quarterbacks","ECR","Running Backs"\n'
            '"  Joe Burrow","4","  Chase Brown","8"\n'
            '"  Jared Goff","16","",""\n'
            '"Week 9 Bye"\n'
            '"ECR","Wide Receivers"\n'
            '"  George Pickens","40"\n',
            encoding="utf-8")
        name_to_team = {"Joe Burrow": "CIN", "Chase Brown": "CIN",
                        "Jared Goff": "DET", "George Pickens": "DAL"}
        byes = parse_team_byes(f, name_to_team)
    assert byes == {"CIN": 6, "DET": 6, "DAL": 9}, byes


def test_parse_team_byes_missing_file_is_empty():
    assert parse_team_byes(Path("nope_does_not_exist_12345.csv"), {}) == {}


def test_roster_bye_counts_groups_and_ignores_missing():
    rows = [_p("a", "RB", 1, Bye=7), _p("b", "WR", 1, Bye=7),
            _p("c", "QB", 1, Bye=11), _p("d", "TE", 1)]        # d has no bye
    assert roster_bye_counts(rows) == {7: ["a", "b"], 11: ["c"]}


def test_bye_stack_caution_fires_at_third_shared_bye():
    mine = [_p("a", "RB", 1, Bye=9), _p("b", "WR", 1, Bye=9)]
    assert bye_stack_caution(9, mine) is not None       # this pick would be #3
    assert bye_stack_caution(10, mine) is None
    assert bye_stack_caution(None, mine) is None
    assert bye_stack_caution(9, mine[:1]) is None       # only #2, not a stack yet


def test_guardrail_bye_stack_warning_and_error():
    warn_rows = [_p(f"p{i}", "WR", 50 - i, Bye=11) for i in range(3)]
    assert any("Week 11" in m and lvl == "warning"
               for lvl, m in roster_guardrails(warn_rows))
    err_rows = [_p(f"p{i}", "WR", 50 - i, Bye=11) for i in range(4)]
    assert any("Week 11" in m and lvl == "error"
               for lvl, m in roster_guardrails(err_rows))


def test_upcoming_my_picks_pair_and_end_of_draft():
    my = snake_pick_slots()                    # [8, 13, 28, 33, ...]
    assert upcoming_my_picks(1, my) == (8, 13)
    assert upcoming_my_picks(8, my) == (8, 13)     # on the clock at 8
    assert upcoming_my_picks(9, my) == (13, 28)
    assert upcoming_my_picks(my[-1], my) == (my[-1], None)   # last pick
    assert upcoming_my_picks(999, my) == (None, None)        # draft over


def test_adp_pick_risk_gone_fringe_safe():
    # next pick #8, pick after #13 (team scale ~0.833)
    assert adp_pick_risk(6, 8, 13) == "gone"       # eff 5.0  <= 8
    assert adp_pick_risk(12, 8, 13) == "fringe"    # eff 10.0 in (8, 13)
    assert adp_pick_risk(18, 8, 13) == ""          # eff 15.0 >= 13 -> safe
    assert adp_pick_risk(float("nan"), 8, 13) == ""
    assert adp_pick_risk(5, None, None) == ""       # draft over -> no signal
    assert adp_pick_risk(5, 153, None) == "gone"    # last pick: only gone/safe
    assert adp_pick_risk(400, 153, None) == ""


def test_adp_team_scale_makes_players_go_earlier():
    # a 12-team ADP of 24 is ~pick 20 in a 10-team room
    assert abs(24 * ADP_TEAM_SCALE - 20.0) < 1e-9
    # borderline at raw ADP == next_pick still counts as gone after scaling
    assert adp_pick_risk(8, 8, 13) == "gone"


def test_flag_bg_handcuff_has_colour_and_no_dead_caution():
    assert "CAUTION" not in _FLAG_BG
    assert "HANDCUFF" in _FLAG_BG
    assert _flag_bg("HANDCUFF") != ""
    assert _flag_bg("handcuff") == _flag_bg("HANDCUFF")   # case-insensitive
    assert _flag_bg("CAUTION") == ""                       # unknown -> no style
    assert _flag_bg("") == ""


def test_apply_news_overlay_adds_columns_and_maps():
    df = pd.DataFrame({
        "Name": ["Ashton Jeanty", "Mike Washington Jr.", "Nobody Special"],
        "Team": ["LV", "LV", "DET"],
        "Notes": ["ECR-only", "ECR-only", "ECR-only"],
    })
    news = {
        "Ashton Jeanty": {"monitor": True, "note": "ankle sprain, monitor"},
        "Mike Washington Jr.": {"handcuff_for": "Ashton Jeanty", "note": "direct HC"},
    }
    out = apply_news_overlay(df, news, {"DET", "KC"})
    j = out[out.Name == "Ashton Jeanty"].iloc[0]
    assert j["Monitor"] is True or j["Monitor"] == True  # noqa: E712
    assert j["NewsNote"] == "ankle sprain, monitor"
    assert j["HandcuffFor"] == ""
    assert j["HighOffense"] == False           # LV not in high-offense set  # noqa: E712
    w = out[out.Name == "Mike Washington Jr."].iloc[0]
    assert w["HandcuffFor"] == "Ashton Jeanty"
    n = out[out.Name == "Nobody Special"].iloc[0]
    assert n["NewsNote"] == "" and n["Monitor"] == False  # noqa: E712
    assert n["HighOffense"] == True            # DET is in the set  # noqa: E712
    # original df untouched (non-destructive)
    assert "NewsNote" not in df.columns


def test_effective_note_prefixes_fresh_keeps_base():
    assert effective_note({"NewsNote": "", "Notes": "researched"}) == "researched"
    assert effective_note({"NewsNote": "fresh", "Notes": ""}) == "🆕 fresh"
    combined = effective_note({"NewsNote": "fresh", "Notes": "researched"})
    assert combined.startswith("🆕 fresh") and "researched" in combined
    assert effective_note({"NewsNote": float("nan"), "Notes": "researched"}) == "researched"


def test_draft_day_news_names_all_exist_in_pool():
    """Guards against a typo'd overlay key silently doing nothing."""
    import draft_day_news
    pool = set(pd.read_csv("players_2026.csv")["Name"])
    missing = [n for n in draft_day_news.NEWS if n not in pool]
    assert not missing, f"names in draft_day_news.NEWS not in players_2026.csv: {missing}"
    bad_teams = [t for t in draft_day_news.HIGH_OFFENSE_TEAMS if len(t) not in (2, 3)]
    assert not bad_teams, f"HIGH_OFFENSE_TEAMS should be team codes: {bad_teams}"


def test_full_sos_label_scale_matches_early_vocabulary():
    assert full_sos_label(5) == "Very Soft"     # 5 stars = easiest
    assert full_sos_label(1) == "Gauntlet"      # 1 star  = toughest
    assert full_sos_label(3) == "Neutral"
    assert full_sos_label(None) == "—"
    assert full_sos_label(float("nan")) == "—"


def test_parse_full_sos_from_synthetic():
    import tempfile
    with tempfile.TemporaryDirectory() as td:
        f = Path(td) / "sos.csv"
        f.write_text(
            '"Star ratings: 1 = toughest schedule, 5 = easiest"\n'
            "\n"
            "TEAM,QB,RB,WR,TE,K,DST\n"
            "Detroit Lions,5,5,4,3,5,5\n"
            "San Francisco 49ers,1,2,2,3,5,3\n"
            "Not A Real Team,4,4,4,4,4,4\n",
            encoding="utf-8")
        m = parse_full_sos(f)
    assert m[("DET", "QB")] == 5 and m[("DET", "TE")] == 3
    assert m[("SF", "QB")] == 1
    assert not any(t == "Not A Real Team" for (t, _) in m)   # unknown name dropped
    assert ("DET", "K") not in m                              # only QB/RB/WR/TE


def test_parse_full_sos_missing_file_is_empty():
    assert parse_full_sos(Path("nope_missing_sos_98765.csv")) == {}


def test_full_sos_real_file_covers_all_32_teams_and_all_positions():
    import draft_app
    m = parse_full_sos(draft_app.FULL_SOS_PATH)
    teams = {t for (t, _) in m}
    assert len(teams) == 32, sorted(teams)
    for t in teams:
        for pos in ("QB", "RB", "WR", "TE"):
            assert (t, pos) in m, f"missing {(t, pos)}"
            assert 1 <= m[(t, pos)] <= 5
    # every non-FA player in the pool resolves to a full-season SoS
    df = pd.read_csv("players_2026.csv")
    miss = df[(df.Team != "FA") &
              df.apply(lambda r: (r.Team, r.Pos) not in m, axis=1)]
    assert miss.empty, miss[["Name", "Team", "Pos"]].to_dict("records")


def _run_all():
    fns = [v for k, v in sorted(globals().items()) if k.startswith("test_") and callable(v)]
    failed = 0
    for fn in fns:
        try:
            fn()
            print(f"  ok   {fn.__name__}")
        except AssertionError as e:
            failed += 1
            print(f"  FAIL {fn.__name__}: {e}")
        except Exception as e:  # noqa: BLE001
            failed += 1
            print(f"  ERR  {fn.__name__}: {type(e).__name__}: {e}")
    print(f"\n{len(fns) - failed}/{len(fns)} passed")
    return failed


if __name__ == "__main__":
    import sys
    sys.exit(1 if _run_all() else 0)

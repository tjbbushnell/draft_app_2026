"""
Draft-day manual overlay for draft_app.py.

`draft_app.py` layers this on top of `players_2026.csv` at load time. It is
NON-DESTRUCTIVE: players_2026.csv is never touched, and this file survives a
board rebuild by build_draft_kit_2026.py. Edit it freely right up to (and
during) the draft, then hit "Rerun" in the app.

Everything in here is hand-entered draft-day intel -- NOT model output.
"""

# ---------------------------------------------------------------------------
# Per-player overlays.  Key = exact Name as it appears in players_2026.csv.
#   note         : str  -- fresh situational note (shown 🆕-prefixed on the
#                          board Notes cell and as its own line in the drawer;
#                          the researched note is kept behind it)
#   monitor      : True -- 🚑 "injury-monitored" badge + amber name on the board
#   handcuff_for : "Starter Name" -- 🔒 badge + an explicit linkage line in the
#                  drawer (use this for handcuffs whose base Flag is something
#                  else, e.g. Jonathon Brooks is flagged BUST but is also
#                  Hubbard's contingency)
# ---------------------------------------------------------------------------
NEWS: dict[str, dict] = {
    "Ashton Jeanty": {
        "monitor": True,
        "note": "8/29 camp: RIGHT ANKLE sprain in practice (avoided a high-ankle "
                "sprain). Monitor Week 1 availability -- Mike Washington Jr. is the "
                "direct contingency.",
    },
    "Mike Washington Jr.": {
        "handcuff_for": "Ashton Jeanty",
        "note": "Rookie. Direct Jeanty handcuff and an immediate riser if the ankle "
                "costs Jeanty any Week 1 time.",
    },
    "Breece Hall": {
        "monitor": True,
        "note": "Early-Aug groin strain; trainers expect him ready for Week 1. Watch "
                "the Braelon Allen touch split early.",
    },
    "Puka Nacua": {
        "note": "8/29: psoas concern cleared, tracking full-go for Week 1 -- drop the "
                "injury fade, lock in foundational value.",
    },
    "Ja'Marr Chase": {
        "note": "8/29: hamstring scare cleared, tracking full-go for Week 1 -- "
                "foundational WR1, no panic.",
    },
    "Chuba Hubbard": {
        "monitor": True,
        "note": "Camp hamstring strain; Canales expects Week 1 but an early committee "
                "is likely -- Jonathon Brooks is the contingency.",
    },
    "Jonathon Brooks": {
        "handcuff_for": "Chuba Hubbard",
        "note": "ALSO a handcuff, not just a fade: torn ACL twice and Hubbard is the "
                "starter -- but Hubbard's camp hamstring opens an early-committee "
                "path. Speculative stash; the injury risk is still live.",
    },
    "Luther Burden III": {
        "note": "8/29: back from an early-Aug groin injury, on track for Week 1 -- "
                "explosive upside available at an ADP discount.",
    },
}

# ---------------------------------------------------------------------------
# Elite offensive environments / high weekly implied totals.
# There is NO implied-total or Vegas data in this project -- this is YOUR read.
# Seeded from a heuristic (teams with the most top-80 GlobalRank skill players);
# edit to taste before the draft. Players on these teams get an ⚡ tag and can
# be isolated with the "⚡ high-offense only" filter -- handy in rounds 11-16.
# ---------------------------------------------------------------------------
HIGH_OFFENSE_TEAMS: set[str] = {
    "DET", "BUF", "CIN", "PHI", "LAR", "KC", "GB", "DAL", "CHI", "BAL",
}

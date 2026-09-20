# ══════════════════════════════════════════════════════════════════════════════
# NBA GAME ENVIRONMENT — DAILY RUNBOOK
# Copy and paste each block as needed. Run from the project root.
# ══════════════════════════════════════════════════════════════════════════════

# ── OPTIONAL: set your Odds API key (once per session) ────────────────────────
# Without a key everything still runs; the implied team total columns are blank.
# Better: add ODDS_API_KEY=... to your .Renviron so it persists.

Sys.setenv(ODDS_API_KEY = "your_key_here")

# ── STEP 1: PULL TODAY'S GAME ENVIRONMENT ─────────────────────────────────────
# Pulls L15 team stats + opponent defensive profile + implied team totals.
# Outputs a color-coded Excel file to output/ and opens it automatically.

source("R/game_env.R")
df <- run_game_env()                      # today

# ── OPTIONAL: run for a specific date ─────────────────────────────────────────

df <- run_game_env("05/08/2026")          # swap in target date (MM/DD/YYYY)

# ── OPTIONAL: export without running the full slate printout ──────────────────

df <- build_game_env()
export_game_env(df, path = "output/nba_env_custom.xlsx")

# ── OUTPUT ────────────────────────────────────────────────────────────────────
# File saved to: output/nba_env_YYYY-MM-DD_HHMMSS.xlsx
# Open output folder:
system("open output/")

# ── COLUMN REFERENCE ──────────────────────────────────────────────────────────
#
#   Overview:
#     team             team name
#     matchup          "vs X" (home) or "at X" (away)
#     b2b              back-to-back flag (orange = second night in a row)
#     tip_off          game time
#
#   Points averages:
#     implied          implied team total for this game
#     ppg_avg          season PPG
#     ppg_l15          last 15 games PPG
#     ppg_playoffs     playoffs PPG (NA if not in postseason)
#     plus_minus       implied minus season PPG  (green = above average)
#     plus_minus_l15   implied minus L15 PPG
#     plus_minus_po    implied minus playoffs PPG
#
#   Pace (L15):
#     opp_pace_rank    opponent pace rank (1 = slowest)
#     pace_rank        team pace rank
#     proj_pace        average of both teams' pace (see key at bottom of sheet)
#
#   Points defense (L15):
#     opp_papg_rank    opponent points-allowed rank (30 = weakest defense)
#     opp_drat_rank    opponent defensive rating rank
#
#   Rebounding (L15):
#     opp_reb_rank     opponent rebounds-allowed rank
#     opp_reb_pct_rank opponent rebound % rank
#
#   Playoff mirrors (same columns, playoff sample only, 16-team pool):
#     *_po columns     same metrics, ranked 1-16 within the playoff field

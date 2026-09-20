# NBA Game Environment

A daily R script that builds a color-coded Excel "game environment" sheet for every NBA game on a given date. For each team in each game it pulls recent form, pace, and opponent defensive rankings, flags back-to-backs, and (optionally) adds the implied team total from market lines. It also builds a parallel set of playoff-only columns ranked within the 16-team playoff field.

It is a data-wrangling and presentation tool, not a predictive model. The value is in putting the relevant context for a whole slate in one sheet.

## What is in the sheet

One row per team per game.

| Section | Columns | Notes |
|---|---|---|
| Overview | `team`, `matchup`, `b2b`, `tip_off` | `b2b` is flagged when the team played the previous night |
| Scoring | `implied`, `ppg_avg`, `ppg_l15`, `ppg_playoffs`, `plus_minus`, `plus_minus_l15`, `plus_minus_po` | `plus_minus` is the implied team total minus the team's average, so positive means the game is expected to be higher scoring than the team's norm |
| Pace | `opp_pace_rank`, `pace_rank`, `proj_pace` | `proj_pace` is the average of both teams' last-15 pace |
| Points defense | `opp_papg_rank`, `opp_drat_rank` | Opponent points allowed and defensive rating ranks (1 = toughest defense) |
| Rebounding | `opp_reb_rank`, `opp_reb_pct_rank` | Opponent rebounding ranks |
| Playoff mirrors | `*_po` columns | The same metrics using playoff games only, ranked 1 to 16 |

Rank cells are colored by tier, with opponent-defense ranks inverted so green always means an easier scoring environment. `proj_pace` is colored on an absolute scale, and a pace key is written under the table.

## Sample output

From a run on May 20, 2026 (columns trimmed):

| team | matchup | implied | ppg_avg | ppg_l15 | plus_minus | proj_pace | opp_papg_rank |
|---|---|---|---|---|---|---|---|
| Oklahoma City Thunder | vs San Antonio Spurs | 111.8 | 119.0 | 121.1 | -7.2 | 99.9 | 7 |
| San Antonio Spurs | at Oklahoma City Thunder | 105.2 | 119.8 | 124.3 | -14.6 | 99.9 | 5 |

## Data sources

| Source | Provides | Access |
|---|---|---|
| stats.nba.com | League-wide team stats (base, advanced, opponent) for the regular season, last 15 games, and playoffs | Public but unofficial endpoints, no key |
| cdn.nba.com | Daily scoreboard and static season schedule | Public JSON |
| [The Odds API](https://the-odds-api.com) (optional) | Market totals and spreads, used to derive implied team totals as `(total -/+ spread) / 2`, and as a last-resort schedule source | Free API key |

The schedule lookup tries four sources in order (stats scoreboard, CDN scoreboard, static season schedule, then games listed by the odds provider) and uses the first that returns games for the date.

## Running it

```r
install.packages(c("tidyverse", "httr", "jsonlite", "openxlsx"))

# Optional: implied team totals. Put this in your .Renviron, never in code.
# ODDS_API_KEY=your_key_here

source("R/game_env.R")
df <- run_game_env()                 # today
df <- run_game_env("05/08/2026")     # a specific date (MM/DD/YYYY)
```

`RUNBOOK.R` has these steps as copy-paste blocks plus a column reference. The workbook is written to `output/nba_env_<timestamp>.xlsx`. Without an API key everything still works and the implied-total columns are simply blank.

## Limitations

- **Unofficial, unauthenticated endpoints.** The stats.nba.com endpoints are not a supported public API. They can rate limit, time out, block some networks, or change without notice, and you should check their terms of use before relying on them. When I re-tested this before publishing, stats.nba.com timed out and the static schedule returned a 403 from my network, so live results depend on where you run it.
- **Raw tip-off time on fallback paths.** When the schedule comes from the static file or the odds provider, `tip_off` is a raw UTC string or blank rather than a formatted Eastern time.
- **Hard-coded season and thresholds.** The season string, the rank tiers, and the pace color thresholds are constants in the script.
- **macOS only for auto-open.** The script opens the finished workbook with the macOS `open` command.
- **No tests and no history.** Each run is a snapshot. Nothing is stored between days and there are no automated tests.

## What I'd build next

- Format tip-off times consistently on every schedule path.
- Store each day's table so pace and defense trends can be charted across a season.
- Add a rest-days and schedule-density view across the whole league.
- Publish the sheet as a small web page instead of an Excel file.

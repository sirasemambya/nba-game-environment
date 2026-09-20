# game_env.R
# Daily NBA game environment table
# Pulls L10 team stats + opponent defensive profile + implied totals
# Outputs a color-coded Excel file for game-context analysis

library(tidyverse)
library(httr)
library(jsonlite)
library(openxlsx)

ODDS_API_KEY <- Sys.getenv("ODDS_API_KEY")   # optional; implied totals are blank without it
LINE_SOURCE  <- "pinnacle"                  # data provider key used for implied team totals
SEASON       <- "2025-26"

# ── NBA Stats API ──────────────────────────────────────────────────────────────

nba_headers <- function() {
  c(
    "Host"               = "stats.nba.com",
    "User-Agent"         = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept"             = "application/json, text/plain, */*",
    "Accept-Language"    = "en-US,en;q=0.9",
    "Accept-Encoding"    = "gzip, deflate, br",
    "x-nba-stats-origin" = "stats",
    "x-nba-stats-token"  = "true",
    "Origin"             = "https://www.nba.com",
    "Referer"            = "https://www.nba.com/",
    "Sec-Fetch-Dest"     = "empty",
    "Sec-Fetch-Mode"     = "cors",
    "Sec-Fetch-Site"     = "same-site",
    "Connection"         = "keep-alive"
  )
}

nba_get <- function(endpoint, params = list(), max_tries = 3) {
  url <- paste0("https://stats.nba.com/stats/", endpoint)
  for (i in seq_len(max_tries)) {
    resp <- tryCatch(
      GET(url, add_headers(.headers = nba_headers()), query = params,
          config = httr::timeout(60)),
      error = function(e) { message("  Attempt ", i, " failed: ", e$message); NULL }
    )
    if (!is.null(resp) && status_code(resp) == 200)
      return(content(resp, "parsed", encoding = "UTF-8"))
    if (!is.null(resp))
      message("  HTTP ", status_code(resp), " on attempt ", i)
    if (i < max_tries) Sys.sleep(3 * i)
  }
  warning("NBA stats API failed for: ", endpoint)
  NULL
}

parse_nba <- function(resp, idx = 1, name = NULL) {
  if (is.null(resp)) return(NULL)
  rs_list <- resp$resultSets
  if (!is.null(name)) {
    nm <- which(sapply(rs_list, function(x) x$name) == name)
    idx <- if (length(nm) > 0) nm[1] else idx
  }
  rs <- rs_list[[idx]]
  if (is.null(rs) || length(rs$rowSet) == 0) return(tibble())
  hdrs <- unlist(rs$headers)
  rows <- lapply(rs$rowSet, function(r) lapply(r, function(x) if (is.null(x)) NA else x))
  df   <- as.data.frame(do.call(rbind, rows), stringsAsFactors = FALSE)
  names(df) <- hdrs
  as_tibble(df) %>%
    mutate(across(where(is.list), ~ unlist(.))) %>%
    mutate(across(where(~ all(!is.na(.) & grepl("^-?[0-9.]+$", as.character(.)))), as.numeric))
}

# ── 1. Today's schedule ────────────────────────────────────────────────────────

pull_schedule_cdn <- function(date_mmddyyyy) {
  # NBA CDN live scoreboard — reliable for today's games including playoffs
  iso <- format(as.Date(date_mmddyyyy, "%m/%d/%Y"), "%Y-%m-%d")
  resp <- tryCatch(
    GET("https://cdn.nba.com/static/json/liveData/scoreboard/todaysScoreboard_00.json",
        config = httr::timeout(15)),
    error = function(e) NULL
  )
  if (is.null(resp) || status_code(resp) != 200) return(tibble())
  data <- tryCatch(fromJSON(content(resp, "text", encoding = "UTF-8"), simplifyVector = FALSE),
                   error = function(e) NULL)
  if (is.null(data)) return(tibble())
  sb <- data$scoreboard
  if (is.null(sb)) return(tibble())
  cdn_date <- if (!is.null(sb$gameDate)) sb$gameDate else ""
  if (!startsWith(cdn_date, iso)) {
    message("  CDN gameDate=", cdn_date, " (stale — not using)")
    return(tibble())
  }
  games <- sb$games
  if (length(games) == 0) { message("  CDN: no games listed for ", cdn_date); return(tibble()) }
  bind_rows(lapply(games, function(g) tibble(
    game_id = g$gameId,
    tip_off = g$gameStatusText,
    home_id = as.numeric(g$homeTeam$teamId),
    away_id = as.numeric(g$awayTeam$teamId)
  )))
}

pull_schedule_static <- function(date_mmddyyyy) {
  # NBA static season schedule JSON — includes playoff games as bracket is set
  resp <- tryCatch(
    GET("https://cdn.nba.com/static/json/staticData/scheduleLeagueV2_1.json",
        add_headers(
          "User-Agent" = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
          "Referer"    = "https://www.nba.com/",
          "Origin"     = "https://www.nba.com"
        ),
        config = httr::timeout(20)),
    error = function(e) { message("  Static fetch error: ", e$message); NULL }
  )
  if (is.null(resp)) { message("  Static: NULL response"); return(tibble()) }
  message("  Static HTTP status: ", status_code(resp))
  if (status_code(resp) != 200) return(tibble())
  data <- tryCatch(
    fromJSON(content(resp, "text", encoding = "UTF-8"), simplifyVector = FALSE),
    error = function(e) { message("  Static JSON parse error: ", e$message); NULL }
  )
  if (is.null(data)) return(tibble())
  game_dates <- data$leagueSchedule$gameDates
  message("  Static gameDates length: ", length(game_dates))
  if (is.null(game_dates)) return(tibble())

  d             <- as.Date(date_mmddyyyy, "%m/%d/%Y")
  prefix_mdy    <- format(d, "%m/%d/%Y")
  prefix_iso    <- format(d, "%Y-%m-%d")

  all_dates <- sapply(game_dates, function(gd) if (!is.null(gd$gameDate)) gd$gameDate else "")
  last_dates <- tail(all_dates[all_dates != ""], 10)
  message("  Static schedule last 10 dates: ", paste(last_dates, collapse = " | "))

  day <- NULL
  for (gd in game_dates) {
    gd_date <- if (!is.null(gd$gameDate)) gd$gameDate else ""
    if (startsWith(gd_date, prefix_mdy) || startsWith(gd_date, prefix_iso)) { day <- gd; break }
  }
  if (is.null(day) || length(day$games) == 0) return(tibble())

  bind_rows(lapply(day$games, function(g) tibble(
    game_id = g$gameId,
    tip_off = if (!is.null(g$gameTimeUTC)) g$gameTimeUTC else "",
    home_id = as.numeric(g$homeTeam$teamId),
    away_id = as.numeric(g$awayTeam$teamId)
  )))
}

pull_schedule_from_odds <- function() {
  if (ODDS_API_KEY == "") return(tibble())
  resp <- tryCatch(
    GET("https://api.the-odds-api.com/v4/sports/basketball_nba/odds/",
        query = list(apiKey = ODDS_API_KEY, regions = "us",
                     markets = "totals", bookmakers = LINE_SOURCE,
                     oddsFormat = "american")),
    error = function(e) NULL
  )
  if (is.null(resp) || status_code(resp) != 200) return(tibble())
  data <- tryCatch(fromJSON(content(resp, "text", encoding = "UTF-8"), flatten = TRUE),
                   error = function(e) NULL)
  if (is.null(data) || length(data) == 0 || nrow(data) == 0) return(tibble())
  bind_rows(lapply(seq_len(nrow(data)), function(i) {
    home_name <- data$home_team[i]
    away_name <- data$away_team[i]
    home_id   <- TEAM_MAP$team_id[TEAM_MAP$odds_name == home_name]
    away_id   <- TEAM_MAP$team_id[TEAM_MAP$odds_name == away_name]
    if (length(home_id) == 0 || length(away_id) == 0) return(NULL)
    message("  Odds API schedule: ", away_name, " at ", home_name)
    tibble(game_id = data$id[i], tip_off = "", home_id = home_id, away_id = away_id)
  }))
}

pull_schedule <- function(date = format(Sys.Date(), "%m/%d/%Y")) {
  message("Pulling schedule for ", date, "...")

  # 1. scoreboardv2
  resp  <- nba_get("scoreboardv2", list(DayOffset = 0, LeagueID = "00", gameDate = date))
  raw   <- parse_nba(resp, name = "GameHeader")
  if (!is.null(raw) && nrow(raw) > 0) {
    sched <- raw %>%
      select(GAME_ID, GAME_STATUS_TEXT, HOME_TEAM_ID, VISITOR_TEAM_ID) %>%
      rename(game_id = GAME_ID, tip_off = GAME_STATUS_TEXT,
             home_id = HOME_TEAM_ID, away_id = VISITOR_TEAM_ID) %>%
      mutate(across(c(home_id, away_id), as.numeric)) %>%
      filter(!is.na(home_id), !is.na(away_id), home_id > 0, away_id > 0)
    if (nrow(sched) > 0) return(sched)
    message("  scoreboardv2 returned rows but team IDs invalid — trying CDN...")
  }

  # 2. NBA CDN live scoreboard
  message("  scoreboardv2 empty — trying NBA CDN...")
  games <- pull_schedule_cdn(date)
  if (!is.null(games) && nrow(games) > 0) return(games)

  # 3. NBA static season schedule (covers early-morning before CDN refreshes)
  message("  CDN empty/stale — trying NBA static schedule...")
  games <- pull_schedule_static(date)
  if (!is.null(games) && nrow(games) > 0) return(games)

  # 4. Odds API fallback — derive schedule from posted lines
  message("  Static schedule empty — trying Odds API schedule fallback...")
  games <- pull_schedule_from_odds()
  if (!is.null(games) && nrow(games) > 0) return(games)

  message("No games found")
  tibble()
}

# ── Check which teams played yesterday (back-to-back detection) ───────────────

get_b2b_teams <- function(date = format(Sys.Date(), "%m/%d/%Y")) {
  yesterday <- format(as.Date(date, "%m/%d/%Y") - 1, "%m/%d/%Y")
  resp  <- nba_get("scoreboardv2", list(DayOffset = 0, LeagueID = "00", gameDate = yesterday))
  games <- parse_nba(resp, name = "GameHeader")
  if (is.null(games) || nrow(games) == 0) return(integer(0))
  c(as.numeric(games$HOME_TEAM_ID), as.numeric(games$VISITOR_TEAM_ID))
}

# ── 2. League-wide team stats (L10) ───────────────────────────────────────────

pull_team_stats <- function(measure, last_n = 10, season_type = "Regular Season") {
  Sys.sleep(0.8)
  message("  Pulling ", measure, " stats (L", last_n, ") [", season_type, "]...")
  # NBA stats API requires ALL parameters or returns HTTP 500
  resp <- nba_get("leaguedashteamstats", list(
    Conference       = "",
    DateFrom         = "",
    DateTo           = "",
    Division         = "",
    GameScope        = "",
    GameSegment      = "",
    Height           = "",
    LastNGames       = last_n,
    LeagueID         = "00",
    Location         = "",
    MeasureType      = measure,
    Month            = 0,
    OpponentTeamID   = 0,
    Outcome          = "",
    PORound          = 0,
    PaceAdjust       = "N",
    PerMode          = "PerGame",
    Period           = 0,
    PlayerExperience = "",
    PlayerPosition   = "",
    PlusMinus        = "N",
    Rank             = "Y",
    Season           = SEASON,
    SeasonSegment    = "",
    SeasonType       = season_type,
    ShotClockRange   = "",
    StarterBench     = "",
    TeamID           = 0,
    TwoWay           = 0,
    VsConference     = "",
    VsDivision       = ""
  ))
  parse_nba(resp, 1)
}

build_team_table <- function() {
  message("Pulling team stats...")
  base      <- pull_team_stats("Base",     last_n = 0)                           # full season PPG
  base_l15  <- pull_team_stats("Base",     last_n = 15)                          # L15 PPG
  base_po   <- pull_team_stats("Base",     last_n = 0, season_type = "Playoffs") # playoffs PPG
  adv       <- pull_team_stats("Advanced", last_n = 15)                          # L15 pace/rating ranks
  opp       <- pull_team_stats("Opponent", last_n = 15)                          # L15 defensive ranks
  adv_po    <- pull_team_stats("Advanced", last_n = 0, season_type = "Playoffs") # playoffs pace/rating ranks
  opp_po    <- pull_team_stats("Opponent", last_n = 0, season_type = "Playoffs") # playoffs defensive ranks

  if (is.null(base)) stop("Failed to pull base stats from NBA stats API")

  # Base: PPG season + L15
  b <- base %>%
    select(TEAM_ID, TEAM_NAME, PTS) %>%
    rename(team_id = TEAM_ID, team_name = TEAM_NAME, ppg = PTS) %>%
    mutate(team_id = as.numeric(team_id))

  b_l15 <- base_l15 %>%
    select(TEAM_ID, PTS) %>%
    rename(team_id = TEAM_ID, ppg_l15 = PTS) %>%
    mutate(team_id = as.numeric(team_id))

  # Playoffs PPG (NA for teams not in playoffs)
  b_po <- if (!is.null(base_po) && nrow(base_po) > 0) {
    base_po %>%
      select(TEAM_ID, PTS) %>%
      rename(team_id = TEAM_ID, ppg_playoffs = PTS) %>%
      mutate(team_id = as.numeric(team_id))
  } else {
    tibble(team_id = numeric(0), ppg_playoffs = numeric(0))
  }

  b <- b %>%
    left_join(b_l15, by = "team_id") %>%
    left_join(b_po,  by = "team_id")

  # Advanced: OffRtg, DefRtg, Pace, Reb%
  a <- adv %>%
    select(TEAM_ID,
           OFF_RATING, OFF_RATING_RANK,
           DEF_RATING, DEF_RATING_RANK,
           PACE,       PACE_RANK,
           REB_PCT,    REB_PCT_RANK) %>%
    rename(team_id = TEAM_ID,
           o_rat = OFF_RATING, o_rat_rank = OFF_RATING_RANK,
           d_rat = DEF_RATING, d_rat_rank = DEF_RATING_RANK,
           pace  = PACE,       pace_rank  = PACE_RANK,
           reb_pct = REB_PCT,  reb_pct_rank = REB_PCT_RANK) %>%
    mutate(team_id = as.numeric(team_id))

  # Opponent: PAPG, Opp Reb%, Opp 3PAtt, Opp 3P%
  o <- opp %>%
    select(TEAM_ID,
           OPP_PTS,     OPP_PTS_RANK,
           OPP_REB,     OPP_REB_RANK,
           OPP_FG3A,    OPP_FG3A_RANK,
           OPP_FG3_PCT, OPP_FG3_PCT_RANK) %>%
    rename(team_id = TEAM_ID,
           papg         = OPP_PTS,     papg_rank         = OPP_PTS_RANK,
           opp_reb      = OPP_REB,     opp_reb_rank      = OPP_REB_RANK,
           opp_3pa      = OPP_FG3A,    opp_3pa_rank      = OPP_FG3A_RANK,
           opp_3p_pct   = OPP_FG3_PCT, opp_3p_pct_rank   = OPP_FG3_PCT_RANK) %>%
    mutate(team_id = as.numeric(team_id))

  # Playoff pace/rating ranks (16-team pool)
  a_po <- if (!is.null(adv_po) && nrow(adv_po) > 0) {
    adv_po %>%
      select(TEAM_ID, PACE, PACE_RANK, DEF_RATING_RANK, REB_PCT_RANK) %>%
      rename(team_id = TEAM_ID,
             pace_po          = PACE,
             pace_rank_po     = PACE_RANK,
             d_rat_rank_po    = DEF_RATING_RANK,
             reb_pct_rank_po  = REB_PCT_RANK) %>%
      mutate(team_id = as.numeric(team_id))
  } else tibble(team_id = numeric(0))

  # Playoff defensive ranks
  o_po <- if (!is.null(opp_po) && nrow(opp_po) > 0) {
    opp_po %>%
      select(TEAM_ID, OPP_PTS_RANK, OPP_REB_RANK) %>%
      rename(team_id = TEAM_ID,
             papg_rank_po    = OPP_PTS_RANK,
             opp_reb_rank_po = OPP_REB_RANK) %>%
      mutate(team_id = as.numeric(team_id))
  } else tibble(team_id = numeric(0))

  b %>%
    left_join(a,    by = "team_id") %>%
    left_join(o,    by = "team_id") %>%
    left_join(a_po, by = "team_id") %>%
    left_join(o_po, by = "team_id")
}

# ── 3. Implied team totals from The Odds API ──────────────────────────────────

pull_implied_totals <- function() {
  if (ODDS_API_KEY == "") {
    message("No ODDS_API_KEY — implied totals will be blank")
    return(tibble())
  }
  message("Pulling NBA totals + spreads...")

  # Pull totals
  tot_resp <- GET(
    "https://api.the-odds-api.com/v4/sports/basketball_nba/odds/",
    query = list(apiKey = ODDS_API_KEY, regions = "us",
                 markets = "totals,spreads", bookmakers = LINE_SOURCE,
                 oddsFormat = "american")
  )
  if (status_code(tot_resp) != 200) {
    message("Odds API: HTTP ", status_code(tot_resp))
    return(tibble())
  }

  data <- fromJSON(content(tot_resp, "text", encoding = "UTF-8"), flatten = TRUE)
  if (length(data) == 0 || nrow(data) == 0) return(tibble())

  results <- list()
  for (i in seq_len(nrow(data))) {
    home_team <- data$home_team[i]
    away_team <- data$away_team[i]
    srcs      <- data$bookmakers[[i]]

    if (is.null(srcs) || length(srcs) == 0) next

    src_idx <- which(srcs$key == LINE_SOURCE)
    if (length(src_idx) == 0) next
    mkts <- srcs$markets[[src_idx[1]]]
    if (is.null(mkts)) next

    total_line  <- NA_real_
    spread_home <- NA_real_

    for (j in seq_len(nrow(mkts))) {
      outcomes <- mkts$outcomes[[j]]
      if (is.null(outcomes)) next

      if (mkts$key[j] == "totals") {
        total_line <- as.numeric(outcomes$point[1])
      }
      if (mkts$key[j] == "spreads") {
        home_row <- outcomes[outcomes$name == home_team, ]
        if (nrow(home_row) > 0) spread_home <- as.numeric(home_row$point[1])
      }
    }

    if (!is.na(total_line)) {
      sp <- coalesce(spread_home, 0)
      results[[i]] <- tibble(
        home_team    = home_team,
        away_team    = away_team,
        home_implied = round((total_line - sp) / 2, 1),
        away_implied = round((total_line + sp) / 2, 1)
      )
    }
  }
  bind_rows(results)
}

# ── Team name bridge: Odds API full names → NBA stats team IDs ────────────────

TEAM_MAP <- tribble(
  ~odds_name,                   ~team_id,
  "Atlanta Hawks",               1610612737,
  "Boston Celtics",              1610612738,
  "Brooklyn Nets",               1610612751,
  "Charlotte Hornets",           1610612766,
  "Chicago Bulls",               1610612741,
  "Cleveland Cavaliers",         1610612739,
  "Dallas Mavericks",            1610612742,
  "Denver Nuggets",              1610612743,
  "Detroit Pistons",             1610612765,
  "Golden State Warriors",       1610612744,
  "Houston Rockets",             1610612745,
  "Indiana Pacers",              1610612754,
  "Los Angeles Clippers",        1610612746,
  "Los Angeles Lakers",          1610612747,
  "Memphis Grizzlies",           1610612763,
  "Miami Heat",                  1610612748,
  "Milwaukee Bucks",             1610612749,
  "Minnesota Timberwolves",      1610612750,
  "New Orleans Pelicans",        1610612740,
  "New York Knicks",             1610612752,
  "Oklahoma City Thunder",       1610612760,
  "Orlando Magic",               1610612753,
  "Philadelphia 76ers",          1610612755,
  "Phoenix Suns",                1610612756,
  "Portland Trail Blazers",      1610612757,
  "Sacramento Kings",            1610612758,
  "San Antonio Spurs",           1610612759,
  "Toronto Raptors",             1610612761,
  "Utah Jazz",                   1610612762,
  "Washington Wizards",          1610612764
)

# ── 4. Assemble the game environment table ────────────────────────────────────

build_game_env <- function(date = format(Sys.Date(), "%m/%d/%Y")) {
  schedule   <- pull_schedule(date)
  if (nrow(schedule) == 0) return(NULL)

  team_stats <- build_team_table()
  totals     <- pull_implied_totals()
  b2b_ids    <- get_b2b_teams(date)

  rows <- list()

  for (i in seq_len(nrow(schedule))) {
    home_id <- schedule$home_id[i]
    away_id <- schedule$away_id[i]
    tip     <- schedule$tip_off[i]

    home <- team_stats %>% filter(team_id == home_id)
    away <- team_stats %>% filter(team_id == away_id)
    if (nrow(home) == 0 || nrow(away) == 0) {
      message("  Skipping game — team ID not found in stats: home=", home_id, " away=", away_id)
      next
    }

    # Get implied totals
    home_odds_name <- TEAM_MAP$odds_name[TEAM_MAP$team_id == home_id]
    home_impl      <- NA_real_
    away_impl      <- NA_real_

    if (nrow(totals) > 0 && length(home_odds_name) > 0) {
      match <- totals %>% filter(str_detect(home_team, fixed(home_odds_name, ignore_case = TRUE)))
      if (nrow(match) > 0) {
        home_impl <- match$home_implied[1]
        away_impl <- match$away_implied[1]
      }
    }

    proj_pace <- round((home$pace + away$pace) / 2, 1)

    make_row <- function(team_s, opp_s, impl, matchup_str) {
      tibble(
        # ── Section 1: Overview ──
        team       = team_s$team_name,
        matchup    = matchup_str,
        b2b        = team_s$team_id %in% b2b_ids,
        tip_off    = tip,
        implied           = impl,
        ppg_avg           = team_s$ppg,
        ppg_l15           = team_s$ppg_l15,
        ppg_playoffs      = team_s$ppg_playoffs,
        plus_minus        = if (!is.na(impl)) round(impl - team_s$ppg,          1) else NA_real_,
        plus_minus_l15    = if (!is.na(impl)) round(impl - team_s$ppg_l15,      1) else NA_real_,
        plus_minus_po     = if (!is.na(impl) && !is.na(team_s$ppg_playoffs)) round(impl - team_s$ppg_playoffs, 1) else NA_real_,
        # ── Section 2: Pace ──
        opp_pace_rank = opp_s$pace_rank,
        pace_rank     = team_s$pace_rank,
        proj_pace     = proj_pace,
        # ── Section 3: Points ──
        opp_papg_rank = opp_s$papg_rank,
        opp_drat_rank = opp_s$d_rat_rank,
        # ── Section 4: Rebounding ──
        opp_reb_rank     = opp_s$opp_reb_rank,
        opp_reb_pct_rank = opp_s$reb_pct_rank,
        # ── Section 5: Playoff mirrors ──
        opp_pace_rank_po    = opp_s$pace_rank_po,
        pace_rank_po        = team_s$pace_rank_po,
        proj_pace_po        = if (!is.na(team_s$pace_po) && !is.na(opp_s$pace_po))
                                round((team_s$pace_po + opp_s$pace_po) / 2, 1) else NA_real_,
        opp_papg_rank_po    = opp_s$papg_rank_po,
        opp_drat_rank_po    = opp_s$d_rat_rank_po,
        opp_reb_rank_po     = opp_s$opp_reb_rank_po,
        opp_reb_pct_rank_po = opp_s$reb_pct_rank_po
      )
    }

    rows[[length(rows) + 1]] <- make_row(home, away, home_impl, paste0("vs ", away$team_name))
    rows[[length(rows) + 1]] <- make_row(away, home, away_impl, paste0("at ", home$team_name))
  }

  bind_rows(rows)
}

# ── 5. Color-coded Excel export ───────────────────────────────────────────────
#
# Rank color logic (1 = best in league, 30 = worst):
#
#   OWN metrics (o_rat, pace, reb_pct):
#     Rank 1-8   = dark green  (elite)
#     Rank 9-14  = light green
#     Rank 15-20 = neutral/yellow
#     Rank 21-25 = light red
#     Rank 26-30 = dark red    (bad)
#
#   OPP defensive metrics (papg, drat, opp_reb, opp_3pa, opp_3p_pct, opp_pace):
#     Rank 1-8   = dark red    (great defense = tougher scoring matchup)
#     Rank 9-14  = light red
#     Rank 15-20 = neutral
#     Rank 21-25 = light green (soft defense = easier scoring matchup)
#     Rank 26-30 = dark green  (weakest defense = easiest scoring matchup)

export_game_env <- function(df, path = NULL) {
  if (is.null(path))
    path <- paste0("output/nba_env_", format(Sys.time(), "%Y-%m-%d_%H%M%S"), ".xlsx")

  dir.create(dirname(path), showWarnings = FALSE, recursive = TRUE)

  wb <- createWorkbook()
  addWorksheet(wb, "Game Environment")

  ci <- function(nm) which(names(df) == nm)

  # ── Section header styles ──
  hdr_overview   <- createStyle(fontName = "Calibri", fontSize = 10, textDecoration = "bold",
                                 fontColour = "#FFFFFF", halign = "center", wrapText = TRUE,
                                 fgFill = "#1D3557")   # dark navy  — Overview
  hdr_pts_avg    <- createStyle(fontName = "Calibri", fontSize = 10, textDecoration = "bold",
                                 fontColour = "#FFFFFF", halign = "center", wrapText = TRUE,
                                 fgFill = "#2E75B6")   # blue       — Points Averages
  hdr_pace       <- createStyle(fontName = "Calibri", fontSize = 10, textDecoration = "bold",
                                 fontColour = "#FFFFFF", halign = "center", wrapText = TRUE,
                                 fgFill = "#7030A0")   # purple     — Pace
  hdr_points     <- createStyle(fontName = "Calibri", fontSize = 10, textDecoration = "bold",
                                 fontColour = "#FFFFFF", halign = "center", wrapText = TRUE,
                                 fgFill = "#C00000")   # red        — Points
  hdr_rebounding <- createStyle(fontName = "Calibri", fontSize = 10, textDecoration = "bold",
                                 fontColour = "#FFFFFF", halign = "center", wrapText = TRUE,
                                 fgFill = "#538135")   # green      — Rebounding

  cell_base  <- createStyle(fontName = "Calibri", fontSize = 10, halign = "center")
  alt_row    <- createStyle(fgFill = "#F0F4FA")
  num_1dp    <- createStyle(numFmt = "0.0", halign = "center")

  # Rank colors
  dk_green <- createStyle(fgFill = "#1A7C3E", fontColour = "#FFFFFF", halign = "center", textDecoration = "bold", fontName = "Calibri", fontSize = 10)
  lt_green <- createStyle(fgFill = "#A9D18E", fontColour = "#1D4A1D", halign = "center", fontName = "Calibri", fontSize = 10)
  neutral  <- createStyle(fgFill = "#FFFFCC", fontColour = "#7F6000", halign = "center", fontName = "Calibri", fontSize = 10)
  lt_red   <- createStyle(fgFill = "#F4B8B8", fontColour = "#7B0000", halign = "center", fontName = "Calibri", fontSize = 10)
  dk_red   <- createStyle(fgFill = "#C00000", fontColour = "#FFFFFF", halign = "center", textDecoration = "bold", fontName = "Calibri", fontSize = 10)

  pos_diff <- createStyle(fgFill = "#C6EFCE", fontColour = "#276221", halign = "center", numFmt = "+0.0;-0.0", fontName = "Calibri", fontSize = 10)
  neg_diff <- createStyle(fgFill = "#FFC7CE", fontColour = "#9C0006", halign = "center", numFmt = "+0.0;-0.0", fontName = "Calibri", fontSize = 10)
  neu_diff <- createStyle(halign = "center", numFmt = "+0.0;-0.0", fontName = "Calibri", fontSize = 10)

  writeData(wb, "Game Environment", df)

  nrows     <- nrow(df)
  data_rows <- 2:(nrows + 1)
  even_rows <- seq(3, nrows + 1, by = 2)
  ncols     <- ncol(df)

  # Base styles
  addStyle(wb, "Game Environment", cell_base, rows = 1:(nrows+1), cols = 1:ncols, gridExpand = TRUE)
  if (length(even_rows) > 0)
    addStyle(wb, "Game Environment", alt_row, rows = even_rows, cols = 1:ncols, gridExpand = TRUE, stack = TRUE)

  # ── Section header coloring ──
  sec_overview   <- intersect(c("team","matchup","b2b","tip_off"), names(df))
  sec_pts_avg    <- intersect(c("implied","ppg_avg","ppg_l15","ppg_playoffs","plus_minus","plus_minus_l15","plus_minus_po"), names(df))
  sec_pace       <- intersect(c("opp_pace_rank","pace_rank","proj_pace"), names(df))
  sec_points     <- intersect(c("opp_papg_rank","opp_drat_rank"), names(df))
  sec_rebounding <- intersect(c("opp_reb_rank","opp_reb_pct_rank"), names(df))
  sec_po_pace    <- intersect(c("opp_pace_rank_po","pace_rank_po","proj_pace_po"), names(df))
  sec_po_points  <- intersect(c("opp_papg_rank_po","opp_drat_rank_po"), names(df))
  sec_po_reb     <- intersect(c("opp_reb_rank_po","opp_reb_pct_rank_po"), names(df))

  for (col in sec_overview)   addStyle(wb, "Game Environment", hdr_overview,   rows = 1, cols = ci(col), stack = TRUE)
  for (col in sec_pts_avg)    addStyle(wb, "Game Environment", hdr_pts_avg,    rows = 1, cols = ci(col), stack = TRUE)
  for (col in sec_pace)       addStyle(wb, "Game Environment", hdr_pace,       rows = 1, cols = ci(col), stack = TRUE)
  for (col in sec_points)     addStyle(wb, "Game Environment", hdr_points,     rows = 1, cols = ci(col), stack = TRUE)
  for (col in sec_rebounding) addStyle(wb, "Game Environment", hdr_rebounding, rows = 1, cols = ci(col), stack = TRUE)
  for (col in sec_po_pace)    addStyle(wb, "Game Environment", hdr_pace,       rows = 1, cols = ci(col), stack = TRUE)
  for (col in sec_po_points)  addStyle(wb, "Game Environment", hdr_points,     rows = 1, cols = ci(col), stack = TRUE)
  for (col in sec_po_reb)     addStyle(wb, "Game Environment", hdr_rebounding, rows = 1, cols = ci(col), stack = TRUE)

  # Black vertical divider before first playoff column
  if (length(sec_po_pace) > 0) {
    border_left <- createStyle(border = "left", borderColour = "#000000", borderStyle = "medium")
    addStyle(wb, "Game Environment", border_left, rows = 1:(nrows+1), cols = ci(sec_po_pace[1]), stack = TRUE)
  }

  # ── Rank coloring helper ──
  color_rank_col <- function(col_name, direction) {
    if (!col_name %in% names(df)) return(invisible(NULL))
    col_i <- ci(col_name)
    vals  <- as.integer(df[[col_name]])

    if (direction == "low_good") {
      tiers <- list(
        list(range = 1:8,   style = dk_green),
        list(range = 9:13,  style = lt_green),
        list(range = 14:17, style = neutral),
        list(range = 18:22, style = lt_red),
        list(range = 23:30, style = dk_red)
      )
    } else {  # high_good (opponent defensive metrics)
      tiers <- list(
        list(range = 1:8,   style = dk_red),
        list(range = 9:13,  style = lt_red),
        list(range = 14:17, style = neutral),
        list(range = 18:22, style = lt_green),
        list(range = 23:30, style = dk_green)
      )
    }

    for (tier in tiers) {
      r <- which(vals %in% tier$range) + 1
      if (length(r) > 0)
        addStyle(wb, "Game Environment", tier$style, rows = r, cols = col_i, stack = TRUE)
    }
  }

  # Playoff rank coloring helper (16-team pool: 1-4 / 5-8 / 9-12 / 13-16)
  color_rank_col_po <- function(col_name, direction) {
    if (!col_name %in% names(df)) return(invisible(NULL))
    col_i <- ci(col_name)
    vals  <- as.integer(df[[col_name]])
    if (direction == "low_good") {
      tiers <- list(
        list(range = 1:4,   style = dk_green),
        list(range = 5:8,   style = lt_green),
        list(range = 9:12,  style = lt_red),
        list(range = 13:16, style = dk_red)
      )
    } else {
      tiers <- list(
        list(range = 1:4,   style = dk_red),
        list(range = 5:8,   style = lt_red),
        list(range = 9:12,  style = lt_green),
        list(range = 13:16, style = dk_green)
      )
    }
    for (tier in tiers) {
      r <- which(vals %in% tier$range) + 1
      if (length(r) > 0)
        addStyle(wb, "Game Environment", tier$style, rows = r, cols = col_i, stack = TRUE)
    }
  }

  # Pace: own pace rank — low = fast (context neutral, color anyway)
  color_rank_col("pace_rank", "low_good")

  # Proj pace: color by absolute value (NBA avg ~99-100)
  pace_color_tiers <- function(pp) list(
    list(test = pp <= 97,              style = dk_red),
    list(test = pp > 97  & pp <= 99,   style = lt_red),
    list(test = pp > 99  & pp <= 101,  style = neutral),
    list(test = pp > 101 & pp <= 103,  style = lt_green),
    list(test = pp > 103,              style = dk_green)
  )
  if ("proj_pace" %in% names(df)) {
    pp_i <- ci("proj_pace")
    for (tier in pace_color_tiers(as.numeric(df$proj_pace))) {
      r <- which(tier$test) + 1
      if (length(r) > 0) addStyle(wb, "Game Environment", tier$style, rows = r, cols = pp_i, stack = TRUE)
    }
  }
  if ("proj_pace_po" %in% names(df)) {
    pp_i <- ci("proj_pace_po")
    for (tier in pace_color_tiers(as.numeric(df$proj_pace_po))) {
      r <- which(tier$test) + 1
      if (length(r) > 0) addStyle(wb, "Game Environment", tier$style, rows = r, cols = pp_i, stack = TRUE)
    }
  }

  # Opponent defensive/pace metrics: high rank = weaker defense = easier scoring matchup
  color_rank_col("opp_pace_rank", "low_good")
  for (col in c("opp_papg_rank","opp_drat_rank","opp_reb_rank","opp_reb_pct_rank"))
    color_rank_col(col, "high_good")

  # Playoff mirror ranks (16-team pool)
  color_rank_col_po("opp_pace_rank_po", "low_good")
  color_rank_col_po("pace_rank_po",     "low_good")
  for (col in c("opp_papg_rank_po","opp_drat_rank_po","opp_reb_rank_po","opp_reb_pct_rank_po"))
    color_rank_col_po(col, "high_good")

  # +/- coloring (both season and L10)
  for (pm_col in c("plus_minus", "plus_minus_l15", "plus_minus_po")) {
    pm_i <- ci(pm_col)
    if (length(pm_i) == 0) next
    pm    <- df[[pm_col]]
    pos_r <- which(!is.na(pm) & pm > 0)  + 1
    neg_r <- which(!is.na(pm) & pm < 0)  + 1
    neu_r <- which(!is.na(pm) & pm == 0) + 1
    if (length(pos_r) > 0) addStyle(wb, "Game Environment", pos_diff, rows = pos_r, cols = pm_i, stack = TRUE)
    if (length(neg_r) > 0) addStyle(wb, "Game Environment", neg_diff, rows = neg_r, cols = pm_i, stack = TRUE)
    if (length(neu_r) > 0) addStyle(wb, "Game Environment", neu_diff, rows = neu_r, cols = pm_i, stack = TRUE)
  }

  # B2B warning coloring
  if ("b2b" %in% names(df)) {
    b2b_warn <- createStyle(fgFill = "#FF6B00", fontColour = "#FFFFFF",
                             fontName = "Calibri", fontSize = 10,
                             textDecoration = "bold", halign = "center")
    b2b_clear <- createStyle(fgFill = "#FFFFFF", fontName = "Calibri",
                              fontSize = 10, halign = "center")
    b2b_rows   <- which(df$b2b == TRUE)  + 1
    clear_rows <- which(df$b2b == FALSE) + 1
    b2b_col    <- ci("b2b")
    if (length(b2b_rows)   > 0) addStyle(wb, "Game Environment", b2b_warn,  rows = b2b_rows,   cols = b2b_col, stack = TRUE)
    if (length(clear_rows) > 0) addStyle(wb, "Game Environment", b2b_clear, rows = clear_rows, cols = b2b_col, stack = TRUE)
    # Replace TRUE/FALSE with readable labels
    df$b2b <- ifelse(df$b2b, "B2B", "")
    writeData(wb, "Game Environment", df["b2b"], startRow = 2, startCol = b2b_col, colNames = FALSE)
  }

  # Number formatting
  for (col in intersect(c("implied","ppg_avg","ppg_l15","ppg_playoffs","proj_pace","proj_pace_po"), names(df)))
    addStyle(wb, "Game Environment", num_1dp, rows = data_rows, cols = ci(col), stack = TRUE)

  setColWidths(wb, "Game Environment", cols = 1:ncols, widths = "auto")
  freezePane(wb, "Game Environment", firstRow = TRUE, firstCol = FALSE)

  # ── Pace key ──
  key_row <- nrows + 3
  key <- list(
    list(label = "PROJ PACE KEY",  fill = "#1D3557", font = "#FFFFFF", bold = TRUE),
    list(label = "≤ 97  — Very Slow",   fill = "#C00000", font = "#FFFFFF", bold = FALSE),
    list(label = "98-99 — Slow",        fill = "#F4B8B8", font = "#7B0000", bold = FALSE),
    list(label = "100-101 — Average",   fill = "#FFFFCC", font = "#7F6000", bold = FALSE),
    list(label = "102-103 — Fast",      fill = "#A9D18E", font = "#1D4A1D", bold = FALSE),
    list(label = "104+ — Very Fast",    fill = "#1A7C3E", font = "#FFFFFF", bold = FALSE)
  )
  for (k in seq_along(key)) {
    writeData(wb, "Game Environment", key[[k]]$label, startRow = key_row + k - 1, startCol = 1)
    sty <- createStyle(fgFill = key[[k]]$fill, fontColour = key[[k]]$font,
                       fontName = "Calibri", fontSize = 10,
                       textDecoration = if (key[[k]]$bold) "bold" else NULL,
                       halign = "left")
    addStyle(wb, "Game Environment", sty, rows = key_row + k - 1, cols = 1, stack = FALSE)
  }

  saveWorkbook(wb, path, overwrite = TRUE)
  message("Saved: ", path)
  system(paste("open", shQuote(path)))
  invisible(df)
}

# ── Master runner ──────────────────────────────────────────────────────────────

run_game_env <- function(date = format(Sys.Date(), "%m/%d/%Y")) {
  message("\n=== NBA Game Environment — ", format(as.Date(date, "%m/%d/%Y"), "%B %d, %Y"), " ===\n")

  df <- build_game_env(date)

  if (is.null(df) || nrow(df) == 0) {
    message("No games found for ", date)
    return(invisible(NULL))
  }

  n_games <- nrow(df) / 2
  message("\n", n_games, " games on the slate — ", nrow(df), " team rows\n")
  print(df %>% select(team, matchup, implied, ppg_avg, plus_minus, opp_papg_rank, opp_drat_rank, proj_pace, tip_off), n = 50)

  export_game_env(df)
  invisible(df)
}

# ── To run ────────────────────────────────────────────────────────────────────
# From the project root:
# source("R/game_env.R")
# df <- run_game_env()                         # today
# df <- run_game_env("04/03/2026")             # specific date

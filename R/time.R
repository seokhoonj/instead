#' Measure execution time of an expression
#'
#' Evaluate an expression and return the elapsed time in `HH:MM:SS.subsec`
#' format. Intended for quick ad-hoc timing in scripts.
#'
#' @param expr An R expression to evaluate. Use `{ ... }` to time multiple lines.
#'
#' @examples
#' # Time a simple loop
#' \donttest{timeit({ for (i in 1:1e6) i })}
#'
#' @export
timeit <- function(expr) {
  stime <- as.numeric(Sys.time())
  eval(expr)
  etime <- as.numeric(Sys.time())
  sec_to_hms(etime - stime)
}

#' Convert seconds to HH:MM:SS.subsec string
#'
#' Utility to format a numeric number of seconds as a clock-style string with
#' subseconds preserved.
#'
#' @param sec Numeric seconds.
#' @return A character string in the form `HH:MM:SS.subsec`.
#'
#' @keywords internal
sec_to_hms <- function(sec) {
  op <- options(scipen = 14); on.exit(options(op))
  hh  <- sec %/% 3600
  sec <- sec %%  3600
  mm  <- sec %/% 60
  ss  <- sec %%  60
  int <- floor(ss)
  frac_str <- sub("^0\\.", "", sprintf("%.3f", ss - int))
  sprintf("%02d:%02d:%02d.%s", hh, mm, int, frac_str)
}

#' Conditionally execute functions depending on the current time
#'
#' `switch_at_time()` evaluates the current date and time and chooses
#' between two functions — `before()` or `after()` — depending on whether
#' the current time is before or after a specified cutoff `at_time`.
#'
#' Optionally, switching can be restricted to specific calendar rules
#' (e.g., only on certain months, days, or weekdays). The functions
#' `before()` and `after()` can each receive their own argument lists via
#' `before_args` and `after_args`.
#'
#' @param before A function to execute *before* the cutoff time.
#' @param after  A function to execute *after* the cutoff time.
#' @param at_time Character string of the form `"HH:MM"` (24-hour format)
#'   indicating the time boundary at which behavior switches.
#' @param months Optional integer vector (1–12). If supplied, switching
#'   occurs only when the current month is included.
#' @param days Optional integer vector (1–31). If supplied, switching
#'   occurs only when today's day-of-month is included.
#' @param wdays Optional integer vector (0–6) or weekday names
#'   (`"Mon"`, `"Tue"`, …). If supplied, switching occurs only on matching
#'   weekdays (`0 = Sunday`).
#'
#' @param tz A time-zone string passed to `Sys.time()` and used to
#'   construct the cutoff time. Defaults to `"Asia/Seoul"`.
#'
#' @param before_args A named list of arguments passed to `before()`
#'   using `do.call()`.
#' @param after_args  A named list of arguments passed to `after()`
#'   using `do.call()`.
#'
#' @return The return value of either `before()` or `after()`, depending on
#'   the current time and calendar rules.
#'
#' @details
#' The function performs three layers of logic:
#'
#' 1. **Validate time format**
#'    `at_time` must strictly follow `"HH:MM"` (e.g., `"09:30"`, `"15:05"`).
#'
#' 2. **Evaluate optional calendar rules**
#'    If any of `months`, `days`, or `wdays` are provided, switching only
#'    occurs when the current date satisfies *all* supplied rules.
#'    Otherwise, `before()` is executed regardless of time.
#'
#' 3. **Evaluate cutoff time**
#'    When rules are satisfied, the function compares the current time
#'    with today's cutoff timestamp constructed from `at_time`:
#'
#'    - If current time < cutoff -> `before()`
#'    - Otherwise -> `after()`
#'
#'
#' @examples
#' \dontrun{
#' # Simple example: switch at 15:30 every day
#' switch_at_time(
#'   before = function() "Before cutoff",
#'   after  = function() "After cutoff",
#'   at_time = "15:30"
#' )
#'
#' # Example with argument lists
#' switch_at_time(
#'   before = function(x) paste("Before:", x),
#'   after  = function(x) paste("After:", x),
#'   at_time = "10:00",
#'   before_args = list(x = 1),
#'   after_args  = list(x = 2)
#' )
#'
#' # Only switch on Mondays before/after 09:00
#' switch_at_time(
#'   before  = function() "Mon AM",
#'   after   = function() "Mon PM",
#'   at_time = "09:00",
#'   wdays   = "Mon"
#' )
#' }
#'
#' @export
switch_at_time <- function(before,
                           after,
                           at_time,
                           months      = NULL,   # 1~12
                           days        = NULL,   # 1~31
                           wdays       = NULL,   # 0~6 or "Mon"
                           tz          = "Asia/Seoul",
                           before_args = list(),
                           after_args  = list()) {

  #------------------------------------------------------------
  # Validate cutoff format: must be "HH:MM"
  #------------------------------------------------------------
  assert_time(at_time)

  #------------------------------------------------------------
  # Get current time (POSIXlt for components)
  #------------------------------------------------------------
  now <- as.POSIXlt(Sys.time(), tz = tz)

  #------------------------------------------------------------
  # Convert weekday names (e.g., "Mon") to numeric (0–6, POSIXlt$wday)
  #------------------------------------------------------------
  if (!is.null(wdays) && is.character(wdays)) {
    wmap  <- c("Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat")
    wdays <- match(wdays, wmap) - 1L
  }

  #------------------------------------------------------------
  # Calendar rule check: month / day / weekday
  #------------------------------------------------------------
  in_rule <-
    (is.null(months) || (now$mon + 1L) %in% months) &&
    (is.null(days)   ||  now$mday      %in% days  ) &&
    (is.null(wdays)  ||  now$wday      %in% wdays )

  # If calendar rule is not satisfied, run `before()` by default
  if (!in_rule) {
    return(do.call(before, before_args))
  }

  #------------------------------------------------------------
  # Build today's at_time timestamp from "HH:MM"
  #------------------------------------------------------------
  hhmm <- strsplit(at_time, ":", fixed = TRUE)[[1L]]
  hh   <- as.integer(hhmm[1L])
  mm   <- as.integer(hhmm[2L])

  cutoff <- as.POSIXlt(
    sprintf("%04d-%02d-%02d %02d:%02d:00",
            now$year + 1900L,
            now$mon  + 1L,
            now$mday,
            hh, mm),
    tz = tz
  )

  #------------------------------------------------------------
  # Switch logic: before at_time -> before(), after at_time -> after()
  #------------------------------------------------------------
  if (now < cutoff) {
    do.call(before, before_args)
  } else {
    do.call(after,  after_args)
  }
}

#' Cache-aware switching based on a rolling time window
#'
#' `cache_in_time_window()` decides whether to reuse existing cached data
#' or refresh it, based on when the cache was last updated relative to a
#' daily cutoff time and a rolling time window.
#'
#' The function compares a `cache_timestamp` (e.g., file modification time)
#' with a window around today's cutoff time `at_time`. If the timestamp
#' lies within the window, `use_cache()` is called; otherwise `refresh()`
#' is called.
#'
#' @param cache_timestamp A POSIXct/POSIXlt timestamp indicating when the
#'   cache was last updated (e.g., `file.info(path)$mtime`). May be `NA`
#'   or `NULL`, in which case the cache is treated as missing and
#'   `refresh()` is executed.
#' @param at_time Character string `"HH:MM"` (24-hour format) defining the
#'   daily cutoff time used as the center/anchor of the decision window.
#' @param window_days Integer number of days used to build the time window
#'   around (or adjacent to) `at_time`. Defaults to `1L`. See Details.
#' @param tz Time zone used for `Sys.time()` and to interpret `at_time`.
#'   Defaults to `"Asia/Seoul"`.
#' @param use_cache A function to call when `cache_timestamp` falls within
#'   the computed time window (i.e., cache is considered "fresh").
#' @param refresh A function to call when the cache is missing or when
#'   `cache_timestamp` lies outside the time window (i.e., cache is
#'   considered "stale").
#' @param use_cache_args A named list of arguments to pass to `use_cache()`
#'   via `do.call()`. Defaults to an empty list.
#' @param refresh_args A named list of arguments to pass to `refresh()`
#'   via `do.call()`. Defaults to an empty list.
#' @param verbose Logical; if `TRUE`, prints diagnostic information about
#'   the current time, cutoff time, window bounds, and cache timestamp.
#'
#' @details
#' Let `now` be the current time in `tz`, and `cutoff` be today's date at
#' `at_time` in the same time zone.
#'
#' - If `now >= cutoff`, the window is
#'   `[cutoff, cutoff + window_days days]`.
#' - If `now < cutoff`, the window is
#'   `[cutoff - window_days days, cutoff]`.
#'
#' The cache is considered "fresh" when:
#'
#' \preformatted{
#'   cache_timestamp > window_start && cache_timestamp < window_end
#' }
#'
#' (strict inequalities, matching typical "recently updated" checks).
#'
#' If `cache_timestamp` is `NULL` or `NA`, the function always falls back
#' to `refresh()`.
#'
#' This helper is useful for patterns such as:
#'
#' - "If a file was saved between yesterday 15:30 and today 15:30,
#'    reuse it; otherwise, download new data and overwrite the file."
#'
#' @return
#' The return value of either `use_cache()` or `refresh()`, depending on
#' whether the cache is within the time window.
#'
#' @examples
#' \dontrun{
#' # Suppose you compute some expensive result and want to cache it.
#' cache_path <- "cache/result.rds"
#'
#' # Load previous cache timestamp if it exists
#' ts <- if (file.exists(cache_path)) file.info(cache_path)$mtime else NA
#'
#' result <- cache_in_time_window(
#'   cache_timestamp = ts,
#'   at_time         = "09:00",   # daily cutoff at 9 AM
#'   window_days     = 1L,        # accept cache from the last 24 hours
#'   tz              = "Asia/Seoul",
#'
#'   # Function used when cache is fresh
#'   use_cache = function() readRDS(cache_path),
#'
#'   # Function used when cache is stale or missing
#'   refresh = function() {
#'     x <- runif(5)        # pretend this is an expensive computation
#'     saveRDS(x, cache_path)
#'     x
#'   },
#'
#'   verbose = TRUE
#' )
#' }
#'
#' @export
cache_in_time_window <- function(cache_timestamp,
                                 at_time,
                                 window_days   = 1L,
                                 tz            = "Asia/Seoul",
                                 use_cache,
                                 refresh,
                                 use_cache_args = list(),
                                 refresh_args   = list(),
                                 verbose        = FALSE) {
  #------------------------------------------------------------
  # Validate at_time format: must be "HH:MM"
  #------------------------------------------------------------
  assert_time(at_time)

  #------------------------------------------------------------
  # Normalize cache_timestamp
  #------------------------------------------------------------
  if (is.null(cache_timestamp) || (length(cache_timestamp) == 1L && is.na(cache_timestamp))) {
    has_cache <- FALSE
  } else if (inherits(cache_timestamp, c("POSIXct", "POSIXlt"))) {
    has_cache <- TRUE
  } else {
    stop("`cache_timestamp` must be POSIXct/POSIXlt, NA, or NULL.", call. = FALSE)
  }

  #------------------------------------------------------------
  # Current time and today's cutoff
  #------------------------------------------------------------
  now <- as.POSIXlt(Sys.time(), tz = tz)

  hhmm <- strsplit(at_time, ":", fixed = TRUE)[[1L]]
  hh   <- as.integer(hhmm[1L])
  mm   <- as.integer(hhmm[2L])

  cutoff_today <- as.POSIXlt(
    sprintf("%04d-%02d-%02d %02d:%02d:00",
            now$year + 1900L,
            now$mon  + 1L,
            now$mday,
            hh, mm),
    tz = tz
  )

  #------------------------------------------------------------
  # Build rolling window around the cutoff, direction depends on `now`
  # - if now >= cutoff: [cutoff, cutoff + window_days]
  # - if now  < cutoff: [cutoff - window_days, cutoff]
  #------------------------------------------------------------
  if (!is.numeric(window_days) || length(window_days) != 1L || window_days <= 0) {
    stop("`window_days` must be a positive numeric scalar.", call. = FALSE)
  }

  offset <- as.integer(window_days)

  cutoff_other <- cutoff_today
  cutoff_other$mday <- cutoff_other$mday + if (now >= cutoff_today) offset else -offset

  window_bounds <- sort(c(cutoff_today, cutoff_other))

  #------------------------------------------------------------
  # Decide whether cache_timestamp is within the window
  #------------------------------------------------------------
  in_window <- FALSE
  if (has_cache) {
    # Coerce to POSIXct in the target timezone for comparison
    cts <- as.POSIXct(cache_timestamp, tz = tz)
    in_window <- (cts > window_bounds[1L] && cts < window_bounds[2L])
  }

  #------------------------------------------------------------
  # Verbose diagnostics
  #------------------------------------------------------------
  if (verbose) {
    cat(sprintf(
      paste0(
        "now          : %s\n",
        "window_start : %s\n",
        "window_end   : %s\n",
        "cached_ts    : %s\n",
        "in_window    : %s\n"
      ),
      format(as.POSIXct(now              ), "%Y-%m-%d %H:%M:%S %Z"),
      format(as.POSIXct(window_bounds[1L]), "%Y-%m-%d %H:%M:%S %Z"),
      format(as.POSIXct(window_bounds[2L]), "%Y-%m-%d %H:%M:%S %Z"),
      if (has_cache) format(as.POSIXct(cache_timestamp), "%Y-%m-%d %H:%M:%S %Z") else "NA",
      in_window
    ))
  }

  #------------------------------------------------------------
  # Dispatch: use_cache() vs refresh()
  #------------------------------------------------------------
  if (in_window) {
    do.call(use_cache, use_cache_args)
  } else {
    do.call(refresh,  refresh_args)
  }
}

#' Validate an "HH:MM" 24-hour time string
#'
#' Checks that the supplied value is a non-missing character scalar
#' representing a valid 24-hour time in `"HH:MM"` format. Hours must be
#' within `00–23` and minutes within `00–59`. The function is intended for
#' validating user-supplied time specifications used by higher-level
#' scheduling or switching functions.
#'
#' @param x A character scalar expected to represent a time in `"HH:MM"`
#'   24-hour format.
#'
#' @details
#' The function performs two levels of validation:
#'
#' * **Type check** — `x` must be a non-missing character vector of length 1.
#' * **Format check** — The value must match the regular expression
#'   `^(?:[01]?\\d|2[0-3]):[0-5]\\d$`, ensuring:
#'   - hours are `0–23` (accepts `"7:05"`, `"07:05"`, `"23:59"`),
#'   - minutes are `00–59`,
#'   - a literal colon separates hours and minutes.
#'
#' If validation fails, an informative error is thrown. On success,
#' the input is returned invisibly.
#'
#' @return Invisibly returns `x` when validation succeeds.
#'
#' @examples
#' assert_time("09:30")   # OK
#' assert_time("7:05")    # OK
#'
#' \dontrun{
#' assert_time("25:10")   # Error: invalid hour
#' assert_time("09:77")   # Error: invalid minute
#' assert_time(c("09:30","10:00")) # Error: must be scalar
#' }
#'
#' @export
assert_time <- function(x) {

  #------------------------------------------------------------
  # Validate type and length
  #------------------------------------------------------------
  # Must be a non-missing character scalar such as "HH:MM"
  if (!is.character(x) || length(x) != 1L || is.na(x)) {
    stop("`x` must be a non-missing character scalar like 'HH:MM'.",
         call. = FALSE)
  }

  #------------------------------------------------------------
  # Validate format and numeric range
  #------------------------------------------------------------
  # Accept only valid 24-hour time:
  #   - Hours:   00–23 (allow "0X", "XX")
  #   - Minutes: 00–59
  # Regex explanation:
  #   (?:[01]?\\d|2[0-3])  -> hours 0–23
  #   :                    -> literal colon
  #   [0-5]\\d             -> minutes 00–59
  if (!grepl("^(?:[01]?\\d|2[0-3]):[0-5]\\d$", x)) {
    stop(sprintf(
      "`x` must be valid 'HH:MM' 24-hour format (e.g. '09:30', '15:05'). Got: '%s'",
      x
    ), call. = FALSE)
  }

  #------------------------------------------------------------
  # Passed validation — return invisibly
  #------------------------------------------------------------
  invisible(x)
}

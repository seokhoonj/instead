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
#'    - If current time < cutoff → `before()`
#'    - Otherwise → `after()`
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
  .assert_time(at_time)

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

.assert_time <- function(x, arg = "at_time") {

  #------------------------------------------------------------
  # Validate type and length
  #------------------------------------------------------------
  # Must be a non-missing character scalar such as "HH:MM"
  if (!is.character(x) || length(x) != 1L || is.na(x)) {
    stop(sprintf("`%s` must be a non-missing character scalar like 'HH:MM'.", arg),
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
      "`%s` must be valid 'HH:MM' 24-hour format (e.g. '09:30', '15:05'). Got: '%s'",
      arg, x
    ), call. = FALSE)
  }

  #------------------------------------------------------------
  # Passed validation — return invisibly
  #------------------------------------------------------------
  invisible(x)
}

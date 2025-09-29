#' Fast parser for ISO-like date formats
#'
#' A lightweight, vectorized helper to parse ISO-like dates:
#' - "YYYY-MM-DD"
#' - "YYYY/MM/DD"
#' - "YYYYMMDD"
#'
#' This function is fast-path only: it does not attempt to resolve
#' ambiguous formats (e.g., "01/08/2017") and will warn on unsupported
#' patterns. Intended for ETL pipelines, logs, or databases where input
#' is guaranteed to be ISO-like.
#'
#' @param x A character, factor, Date, or POSIXt vector.
#'
#' @return A Date vector the same length as `x`.
#'
#' @examples
#' \donttest{
#' # ISO-like formats are parsed safely
#' as_date_iso(c("2024-01-01", "20240102", "2024/01/03"))
#' #> [1] "2024-01-01" "2024-01-02" "2024-01-03"
#'
#' # Already Date objects pass through unchanged
#' d <- as.Date("2024-01-05")
#' as_date_iso(d)
#' #> [1] "2024-01-05"
#'
#' # POSIXt values drop the time component
#' ts <- as.POSIXct("2024-01-07 12:34:56")
#' as_date_iso(ts)
#' #> [1] "2024-01-07"
#'
#' # Unsupported formats produce a warning
#' as_date_iso("01/08/2024")
#' #> Warning: Unsupported date format(s): 01/08/2024
#' #> [1] NA
#' }
#'
#' @seealso [as_date_safe()]
#'
#' @export
as_date_iso <- function(x) {
  # fast path for Date objects
  if (inherits(x, "Date")) return(x)

  # drop time component if POSIXt
  if (inherits(x, "POSIXt")) return(as.Date(x))

  # coerce factors to character
  if (is.factor(x)) x <- as.character(x)

  # enforce character
  if (!is.character(x)) {
    stop("Input must be Date, POSIXt, character, or factor.", call. = FALSE)
  }

  # normalize empty strings as NA
  x[x == ""] <- NA_character_

  # pre-allocate output
  out <- as.Date(rep(NA_character_, length(x)))

  # patterns & formats (ISO-like only)
  pats <- c("^\\d{4}-\\d{2}-\\d{2}$",  # YYYY-MM-DD
            "^\\d{8}$",                # YYYYMMDD
            "^\\d{4}/\\d{2}/\\d{2}$")  # YYYY/MM/DD
  fmts <- c("%Y-%m-%d", "%Y%m%d", "%Y/%m/%d")

  # parse by pattern (fast and safe)
  for (k in seq_along(pats)) {
    idx <- !is.na(x) & grepl(pats[k], x)
    if (any(idx)) {
      out[idx] <- as.Date(x[idx], format = fmts[k])
    }
  }

  # warn if unsupported formats remain
  bad <- !is.na(x) & is.na(out)
  if (any(bad)) {
    warn_vals <- unique(x[bad])
    n_show <- min(5L, length(warn_vals))
    warning(sprintf(
      "Unsupported date format(s): %s%s",
      paste(warn_vals[seq_len(n_show)], collapse = ", "),
      if (length(warn_vals) > n_show) ", ..." else ""
    ), call. = FALSE)
  }

  out
}

#' Safely coerce to Date with ambiguity checks (ISO-first)
#'
#' Vectorized date parser that accepts character, factor, Date, or POSIXt.
#' Supported formats:
#' - ISO-like: "YYYY-MM-DD", "YYYY/MM/DD", "YYYYMMDD"
#' - Two-digit day/month with "/" or "-": "DD/MM/YYYY" vs "MM/DD/YYYY"
#'   (and "DD-MM-YYYY" vs "MM-DD-YYYY"); ambiguous cases error.
#'
#' Rules:
#' - Dates returned unchanged.
#' - POSIXt drops time (converted via as.Date()).
#' - "" treated as NA.
#' - First parse ISO-like formats (fast, unambiguous), then handle 2-digit day/month.
#' - Anything not parsed -> error with a short preview.
#'
#' @param x character | factor | Date | POSIXt vector
#'
#' @return Date vector the same length as `x`
#'
#' @examples
#' # ISO-like formats
#' as_date_safe(c("2024-01-01", "20240102", "2024/01/03"))
#' #> [1] "2024-01-01" "2024-01-02" "2024-01-03"
#'
#' \dontrun{
#' # Ambiguous two-digit cases raise an error
#' as_date_safe("01/08/2017")
#' #> Error: Ambiguous date(s): 01/08/2017 (could be DD/MM/YYYY or MM/DD/YYYY).
#' }
#'
#' # Unambiguous two-digit cases are parsed
#' as_date_safe(c("13/02/2017", "02/13/2017"))
#' #> [1] "2017-02-13" "2017-02-13"
#'
#' # POSIXt is allowed
#' as_date_safe(as.POSIXct("2024-01-07 12:34:56"))
#' #> [1] "2024-01-07"
#'
#' # Factor input is coerced
#' f <- factor("2024-01-10")
#' as_date_safe(f)
#' #> [1] "2024-01-10"
#'
#' \dontrun{
#' # Unsupported formats -> error
#' as_date_safe("2024.01.01")
#' #> Error: Unsupported date format(s): 2024.01.01.
#' }
#'
#' @seealso [as_date_iso()]
#'
#' @export
as_date_safe <- function(x) {
  # fast-path passthroughs
  if (inherits(x, "Date"))   return(x)
  if (inherits(x, "POSIXt")) return(as.Date(x))
  if (is.factor(x)) x <- as.character(x)
  if (!is.character(x)) {
    stop("Input must be character, factor, Date, or POSIXt.", call. = FALSE)
  }

  # treat "" as NA
  x[x == ""] <- NA_character_

  n   <- length(x)
  out <- as.Date(rep(NA_character_, n))
  if (all(is.na(x))) return(out)

  xs  <- trimws(x)
  idx <- which(!is.na(xs))
  if (!length(idx)) return(out)

  v <- xs[idx]

  # ISO-like first (safe & fast)
  is_ymd_dash  <- grepl("^\\d{4}-\\d{2}-\\d{2}$", v)  # YYYY-MM-DD
  is_ymd_slash <- grepl("^\\d{4}/\\d{2}/\\d{2}$", v)  # YYYY/MM/DD
  is_ymd       <- grepl("^\\d{8}$", v)                # YYYYMMDD

  if (any(is_ymd_dash)) {
    i <- which(is_ymd_dash)
    out[idx[i]] <- as.Date(v[i], "%Y-%m-%d")
  }
  if (any(is_ymd_slash)) {
    i <- which(is_ymd_slash)
    out[idx[i]] <- as.Date(v[i], "%Y/%m/%d")
  }
  if (any(is_ymd)) {
    i <- which(is_ymd)
    out[idx[i]] <- as.Date(v[i], "%Y%m%d")
  }

  # Two-digit day/month with "/" or "-" (disambiguation via helpers)
  is_dmy_slash <- grepl("^\\d{2}/\\d{2}/\\d{4}$", v)
  is_dmy_dash  <- grepl("^\\d{2}-\\d{2}-\\d{4}$", v)

  out <- .handle_two_digit(v, out, idx, is_dmy_slash, "/", "%d/%m/%Y", "%m/%d/%Y")
  out <- .handle_two_digit(v, out, idx, is_dmy_dash,  "-", "%d-%m-%Y", "%m-%d-%Y")

  # Anything still unparsed (non-NA input) is unsupported
  still_bad <- is.na(out[idx]) & !is.na(v)
  if (any(still_bad)) {
    bad_vals <- unique(v[still_bad])
    n_show <- min(5L, length(bad_vals))
    stop(sprintf(
      "Unsupported date format(s): %s%s.",
      paste(bad_vals[seq_len(n_show)], collapse = ", "),
      if (length(bad_vals) > n_show) ", ..." else ""
    ), call. = FALSE)
  }

  out
}

#' Add months to a date
#'
#' Add a specified number of months to a date-like object. Negative values
#' subtract months. Uses base rollover via POSIXlt.
#'
#' @param date A date-like object (coercible via [as.POSIXlt()]).
#' @param mon Integer number of months to add (can be negative).
#'
#' @return A Date with the months added.
#'
#' @examples
#' \donttest{
#' # Add months
#' date <- Sys.Date()
#' add_mon(date, 3)
#' }
#'
#' @export
add_mon <- function(date, mon) {
  date <- as.POSIXlt(date)
  date$mon <- date$mon + mon
  as.Date(date)
}

#' Add years to a date
#'
#' Add a specified number of years to a date-like object. Negative values
#' subtract years. Uses base rollover via POSIXlt.
#'
#' @param date A date-like object (coercible via [as.POSIXlt()]).
#' @param year Integer number of years to add (can be negative).
#'
#' @return A Date with the years added.
#'
#' @examples
#' \donttest{
#' # Add years
#' date <- Sys.Date()
#' add_year(date, 3)
#' }
#'
#' @export
add_year <- function(date, year) {
  date <- as.POSIXlt(date)
  date$year <- date$year + year
  as.Date(date)
}

#' Beginning and end of the month
#'
#' Get the beginning date (first day) or end date (last day) of the month.
#'
#' @param date A Date.
#'
#' @return A Date object representing the beginning (bmonth) or end (emonth)
#'   of the month
#'
#' @examples
#' \donttest{
#' # Beginning of month
#' bmonth(Sys.Date())
#'
#' # End of month
#' emonth(Sys.Date())
#' }
#'
#' @export
bmonth <- function(date) {
  as.Date(format(date, format = "%Y-%m-01"))
}

#' @rdname bmonth
#' @export
emonth <- function(date) {
  add_mon(date, 1L) - 1L
}

#' Check if input is in a recognizable date format
#'
#' Test whether an object is a date or can be interpreted as a date according to
#' package-configured formats.
#'
#' @param x Object to test (Date, POSIXt, character, or factor).
#'
#' @return Logical scalar: `TRUE` if `x` is Date/POSIXt or all elements of `x`
#'   match the configured date format; `FALSE` otherwise.
#'
#' @export
is_date_format <- function(x) {
  if (inherits(x, c("Date", "POSIXt"))) {
    return(TRUE)
  }
  if (inherits(x, c("character", "factor"))) {
    return(all(grepl(local(.DATE_FORMAT, envir = .INSTEAD_ENV), x[!is.na(x)])))
  }
  FALSE
}

#' Format dates as YYYYMM
#'
#' Extract year and month from dates as a compact `YYYYMM` character string.
#'
#' @param x A Date vector.
#'
#' @return A character vector in `YYYYMM` format (e.g., "202312").
#'
#' @examples
#' \donttest{
#' # Get year-month
#' x <- as.Date(c("1999-12-31", "2000-01-01"))
#' yearmon(x)
#' }
#'
#' @export
yearmon <- function(x) {
  substr(format(x, format = "%Y%m%d"), 1L, 6L)
}

#' Difference in months between two dates
#'
#' Compute the number of months between two dates with an optional day
#' threshold. If the first element of `day_limit` is non-zero, the start date
#' counts as a full month when its day is `>= day_limit[1]`, and the end date
#' counts as a full month when its day is `< day_limit[1]`.
#'
#' @param sdate A start date vector
#' @param edate An end date vector
#' @param day_limit Day threshold for counting full months. If the start date's
#'   day is less than this value, it counts as a full month. If the end date's
#'   day is greater than or equal to this value, it counts as a full month.
#'
#' @return A numeric vector representing the number of months between the dates
#'
#' @examples
#' \donttest{
#' # Month difference
#' sdate <- as.Date("1999-12-31")
#' edate <- as.Date("2000-01-01")
#' mondiff(sdate, edate)
#' }
#'
#' @export
mondiff <- function(sdate, edate, day_limit = c(0:31)) {
  # if the sdate month day <  day_limit: count
  # if the edate month day >= day_limit: count
  assert_class(sdate, "Date")
  assert_class(edate, "Date")
  ys <- data.table::year(sdate)
  ye <- data.table::year(edate)
  ms <- data.table::month(sdate)
  me <- data.table::month(edate)
  ds <- de <- 0
  if (day_limit[1L]) {
    ds <- ifelse(data.table::mday(sdate) >= day_limit[1L], 1, 0)
    de <- ifelse(data.table::mday(edate) <  day_limit[1L], 1, 0)
  }
  z <- (ye - ys) * 12 + (me - ms) + 1 - ds - de

  z
}

#' Generate sequences of Dates for each (from, to) range
#'
#' Expands paired start (`from`) and end (`to`) dates into full
#' daily sequences. Each element of the output corresponds to
#' one row of input, and contains a sequence of `Date` values
#' from `from[i]` to `to[i]` (inclusive).
#'
#' Optionally, a `labels` vector can be provided to tag each
#' sequence in the output.
#'
#' @param from A Date vector of start dates.
#' @param to A Date vector of end dates. Must be the same length as `from`.
#' @param labels Optional vector of labels (same length as `from`).
#'
#' @return A list of `Date` vectors, each element representing a full daily
#'   sequence between corresponding `from` and `to`.
#'
#' @details
#' - Both `from` and `to` must be of class `Date`.
#' - Sequences are inclusive of both endpoints.
#' - If `labels` is supplied, it is attached to the resulting list
#'   as element names.
#'
#' @examples
#' from <- as.Date(c("2024-01-01", "2024-02-01"))
#' to   <- as.Date(c("2024-01-03", "2024-02-02"))
#'
#' # Expand to daily sequences
#' seq_dates(from, to)
#'
#' # With labels
#' seq_dates(from, to, labels = c("A", "B"))
#'
#' @export
seq_dates <- function(from, to, labels = NULL) {
  stopifnot(inherits(from, "Date"), inherits(to, "Date"))
  .Call(SeqDates, from, to, labels)
}

#' Collapse overlapping (or near-adjacent) date ranges by (id, group)
#'
#' Merges multiple date ranges that either overlap or are within a user-defined
#' gap (`interval` days) *within each (id, group) block*. Optionally collapses
#' one or more columns (`merge_var`) by concatenating unique values.
#'
#' @param df A data.frame or data.table containing date ranges.
#' @param id_var One or more ID columns (bare names). Ranges are merged within each unique ID.
#' @param group_var Optional additional grouping columns (bare names).
#' @param merge_var Optional columns (bare names) whose values will be collapsed
#'   over the merged range; duplicates and `NA` are removed before collapsing.
#' @param from_var Start-date column (bare name). Must be `Date` (days).
#' @param to_var End-date column (bare name). Must be `Date` (days, inclusive).
#' @param interval Integer gap (in days) allowed between consecutive ranges to still merge.
#'   Use `0` to merge touching ranges, positive values to allow gaps, and **`-1` to require
#'   actual overlap** (touching ranges are *not* merged).
#' @param collapse Character delimiter used to combine `merge_var` values (default `"|"`).
#'
#' @return A data.table with merged, non-overlapping date ranges. The output contains:
#' * the grouping keys (`id_var`, `group_var`),
#' * the merged date columns (`from_var`, `to_var`),
#' * the `stay` column = number of inclusive days in each merged range,
#' * and the collapsed `merge_var` columns if provided.
#'
#' @details
#' * Input is internally sorted by `id_var`, `group_var`, `from_var`, `to_var`.
#' * Dates are treated as **inclusive**; adjacency is tested against `to + 1`.
#' * `interval` is a single integer and must be `>= -1`.
#' * Using `data.table` is recommended for performance; base data.frame also works.
#' * The merging index is computed in C (via `.Call(IndexOverlappingDateRanges, ...)`) for speed.
#'
#' @examples
#' \donttest{
#' id <- c("A","A","B","B")
#' group <- c("x","x","y","y")
#' work  <- c("cleansing", "analysis", "cleansing", "QA")
#' from <- as.Date(c("2022-03-01","2022-03-05","2022-03-08","2022-03-12"))
#' to <- as.Date(c("2022-03-06","2022-03-09","2022-03-10","2022-03-15"))
#' dt <- data.table::data.table(id = id, group = group, work = work,
#'                              from = from, to = to)
#'
#' # Merge touching ranges (interval = 0)
#' collapse_date_ranges(
#'   dt, id_var = id, group_var = group, merge_var = work,
#'   from_var = from, to_var = to,
#'   interval = 0
#' )
#'
#' # Require actual overlap (touching ranges NOT merged): interval = -1
#' collapse_date_ranges(
#'   dt, id_var = id, group_var = group, merge_var = work,
#'   from_var = from, to_var = to,
#'   interval = -1
#' )
#'
#' # Allow up to 2-day gaps
#' collapse_date_ranges(
#'   dt, id_var = id, group_var = group, merge_var = work,
#'   from_var = from, to_var = to,
#'   interval = 2, collapse = ", "
#' )
#' }
#'
#' @export
collapse_date_ranges <- function(df, id_var, group_var, merge_var,
                                 from_var, to_var, interval = 0L,
                                 collapse = "|") {
  assert_class(df, "data.frame")
  env <- ensure_dt_env(df)
  dt  <- env$dt

  id_var    <- capture_names(dt, !!rlang::enquo(id_var))
  group_var <- capture_names(dt, !!rlang::enquo(group_var))
  merge_var <- capture_names(dt, !!rlang::enquo(merge_var))
  from_var  <- capture_names(dt, !!rlang::enquo(from_var))
  to_var    <- capture_names(dt, !!rlang::enquo(to_var))

  # ensure Date type
  stopifnot(inherits(dt[[from_var]], "Date"), inherits(dt[[to_var]], "Date"))

  # basic sanity: from <= to
  if (any(dt[[to_var]] - dt[[from_var]] < 0, na.rm = TRUE))
    stop("Some ranges have `from_var` > `to_var`.", call. = FALSE)

  # interval: allow -1 (require overlap), 0 (touching ok), >0 (allow gaps)
  interval <- as.integer(interval)
  if (length(interval) != 1L || is.na(interval) || interval < -1L)
    stop("`interval` must be a single integer >= -1.", call. = FALSE)

  # materialize working table (only needed columns), rename date cols
  id_group_var <- c(id_var, group_var)
  all_var <- c(id_var, group_var, merge_var, from_var, to_var)
  dt <- dt[, .SD, .SDcols = all_var]
  data.table::setnames(dt, c(id_var, group_var, merge_var, "from", "to"))
  data.table::setorderv(dt, c(id_var, group_var, "from", "to"))

  # bookkeeping column for later adjustments (kept numeric)
  data.table::set(dt, j = "sub_stay", value = 0)

  # compute merging index in C
  index <- .Call(IndexOverlappingDateRanges,
                 dt[, .SD, .SDcols = id_group_var],
                 dt$from, dt$to, interval)

  data.table::set(dt, j = "loc", value = index$loc)  # group index (within id/group)
  data.table::set(dt, j = "sub", value = index$sub)  # subtract days when gaps allowed

  id_group_loc_var <- c(id_group_var, "loc")

  # collapse merge_var (if present) or just keep keys
  if (length(merge_var)) {
    m <- dt[
      ,
      lapply(.SD, function(x) paste(unique(x[!is.na(x)]), collapse = collapse)),
      keyby = id_group_loc_var, .SDcols = merge_var
    ]
  } else {
    m <- unique(dt[, .SD, .SDcols = id_group_loc_var])
    data.table::setkeyv(m, id_group_loc_var)
  }

  # compute merged bounds and stay
  from <- to <- sub_stay <- sub <- NULL
  s <- dt[
    ,
    list(from = min(from),
         to   = max(to),
         sub_stay = sum(sub_stay) + sum(sub)),
    keyby = id_group_loc_var
  ]

  z <- m[s, on = id_group_loc_var]
  data.table::set(z, j = "loc", value = NULL)
  data.table::set(z, j = "stay", value = as.numeric(z$to - z$from + 1 - z$sub_stay))
  data.table::set(z, j = "sub_stay", value = NULL)

  # restore original column names order
  data.table::setnames(z, c(id_var, group_var, merge_var, from_var, to_var, "stay"))

  env$restore(z)
}

#' Split stay intervals by cut dates
#'
#' Splits rows of a data frame into smaller intervals whenever a given
#' set of split dates falls within the `[from, to]` range.
#' Each affected row is divided into two intervals:
#' - Left interval: `[from, split_date - 1]`
#' - Right interval: `[split_date, to]`
#'
#' Useful for scenarios such as hospitalization episodes or claim periods,
#' where intervals need to be aligned with administrative cut-off dates.
#'
#' @param df A data.frame or data.table containing stay intervals.
#' @param from_var Column specifying the start date of the stay.
#'   Can be provided unquoted.
#' @param to_var Column specifying the end date of the stay.
#'   Can be provided unquoted.
#' @param dates A vector of Date objects. Each date is used as a split
#'   point if it falls strictly inside an interval (`from < d <= to`).
#' @param all Logical; if `TRUE` (default), keeps rows without any splits
#'   alongside split rows. If `FALSE`, only the split pieces are returned.
#' @param verbose Logical; if `TRUE` (default), prints progress messages
#'   showing how many rows were affected at each split and a final summary.
#'
#' @return A data.table with the same columns as `df`, but possibly with
#'   more rows due to interval splitting. The original class of `df` is
#'   preserved if wrapped by [ensure_dt_env()].
#'
#' @examples
#' \donttest{
#' df <- data.frame(
#'   id   = 1:3,
#'   from = as.Date(c("2024-01-01","2024-01-10","2024-01-20")),
#'   to   = as.Date(c("2024-01-15","2024-01-20","2024-01-25"))
#' )
#'
#' # Split stays at Jan 5 and Jan 18
#' split_stay_by_date(df, from, to, dates = as.Date(c("2024-01-05","2024-01-18")))
#' }
#'
#' @export
split_stay_by_date <- function(df, from_var, to_var, dates, all = TRUE,
                               verbose = TRUE) {
  assert_class(df, "data.frame")

  env <- ensure_dt_env(df)
  dt  <- env$dt

  # resolve column names
  fv <- capture_names(dt, !!rlang::enquo(from_var))
  tv <- capture_names(dt, !!rlang::enquo(to_var))

  # unique sorted split dates
  dates <- sort(unique(as_date_safe(dates)))

  # for final summary
  n_before <- nrow(dt)

  applied <- 0L
  for (d in dates) {
    # find rows where from < d <= to
    idx <- which(dt[[fv]] < d & dt[[tv]] >= d)
    if (!length(idx)) next

    # shallow subset
    left  <- dt[idx]                 # shallow copy for left interval
    right <- data.table::copy(left)  # deep copy for right interval

    # adjust interval boundaries
    data.table::set(left,  j = tv, value = d - 1L)  # [from, d-1]
    data.table::set(right, j = fv, value = d)       # [d, to]

    if (all) {
      dt <- data.table::rbindlist(
        list(dt[-idx], left, right),
        use.names = TRUE
      )
    } else {
      dt <- data.table::rbindlist(
        list(left, right),
        use.names = TRUE
      )
    }

    applied <- applied + 1L
    if (verbose) {
      affected <- length(idx)
      cat(sprintf(
        "Stay split at %s (affected rows: %s)\n",
        as.Date(d),
        scales::comma(affected)
      ))
    }
  }

  if (verbose) {
    n_after <- nrow(dt)
    delta   <- n_after - n_before

    cat(sprintf(
      "Completed: %s split date(s) applied across %s rows -> %s (%s) rows.\n",
      scales::comma(applied),
      scales::comma(n_before),
      scales::comma(n_after),
      if (delta >= 0) paste0("+", scales::comma(delta)) else scales::comma(delta)
    ))

    cat("! Please review variables derived from stay totals.\n",
        "! Re-calculation may be required after splitting.\n", sep = "")
  }

  env$restore(dt)
}


# Deprecated function -----------------------------------------------------

#' Deprecated: combine_overlapping_date_range
#'
#' `r lifecycle::badge("deprecated")`
#'
#' Use [collapse_date_ranges()] instead.
#'
#' @param ...	Additional arguments passed to combine_overlapping_date_range().
#'
#' @return Same return value as [collapse_date_ranges()]
#'
#' @seealso [collapse_date_ranges()]
#'
#' @export
combine_overlapping_date_range <- function(...) {
  lifecycle::deprecate_warn("0.0.0.9000", "combine_overlapping_date_ranges()", "collapse_date_ranges()")
  collapse_date_ranges(...)
}


# Internal helper functions -----------------------------------------------

#' Split "dd<sep>mm<sep>YYYY" or "mm<sep>dd<sep>YYYY" into numeric parts
#'
#' @keywords internal
#' @noRd
.split_dmy <- function(strings, sep) {
  parts <- strsplit(strings, sep, fixed = TRUE)
  m <- matrix(unlist(parts, use.names = FALSE), ncol = 3, byrow = TRUE)
  data.frame(
    a = as.integer(m[, 1]),
    b = as.integer(m[, 2]),
    Y = as.integer(m[, 3]),
    stringsAsFactors = FALSE
  )
}

#' Disambiguate two-digit day/month dates and fill parsed results
#'
#' Handle strings like "DD/MM/YYYY" vs "MM/DD/YYYY" (or '-' separator).
#' - If only one interpretation is valid, parse with that format.
#' - If both are valid *and* (a <= 12 & b <= 12), stop with an ambiguity error.
#' - If neither valid, leave as NA (caller will handle unsupported cases).
#'
#' @param v    character vector of non-NA values (a subset of original input)
#' @param out  Date vector (same length as original) to be updated
#' @param idx  integer index positions in the original vector corresponding to `v`
#' @param mask logical mask over `v` indicating which elements match the pattern
#' @param sep  separator ("/" or "-")
#' @param fmt_dmy format string for DD<sep>MM<sep>YYYY
#' @param fmt_mdy format string for MM<sep>DD<sep>YYYY
#'
#' @return updated Date vector `out`
#' @keywords internal
#' @noRd
.handle_two_digit <- function(v, out, idx, mask, sep, fmt_dmy, fmt_mdy) {
  if (!any(mask)) return(out)

  # only positions still NA in 'out' and matching the mask
  i <- which(mask & is.na(out[idx]))
  if (!length(i)) return(out)

  s <- v[i]
  d <- .split_dmy(s, sep = sep)

  # validity checks
  valid_dmy <- (d$a >= 1 & d$a <= 31) & (d$b >= 1 & d$b <= 12)
  valid_mdy <- (d$a >= 1 & d$a <= 12) & (d$b >= 1 & d$b <= 31)

  # ambiguous when both valid and both parts <= 12
  ambiguous <- valid_dmy & valid_mdy & (d$a <= 12 & d$b <= 12)

  # assign unambiguous cases
  only_dmy <- valid_dmy & !valid_mdy
  if (any(only_dmy)) {
    j <- i[only_dmy]                 # positions within v
    out[idx[j]] <- as.Date(s[only_dmy], format = fmt_dmy)
  }

  only_mdy <- valid_mdy & !valid_dmy
  if (any(only_mdy)) {
    j <- i[only_mdy]
    out[idx[j]] <- as.Date(s[only_mdy], format = fmt_mdy)
  }

  # ambiguous -> stop with message
  if (any(ambiguous)) {
    amb_vals <- unique(s[ambiguous])
    n_show <- min(5L, length(amb_vals))
    stop(sprintf(
      "Ambiguous date(s): %s%s (could be DD%sMM%sYYYY or MM%sDD%sYYYY).",
      paste(amb_vals[seq_len(n_show)], collapse = ", "),
      if (length(amb_vals) > n_show) ", ..." else "",
      sep, sep, sep, sep
    ), call. = FALSE)
  }

  out
}

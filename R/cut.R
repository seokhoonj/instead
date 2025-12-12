#' Create binned "bands" from a numeric column (class-preserving)
#'
#' Convert a numeric variable into ordered bands (factor levels) using
#' pretty, human-readable labels. Works with `data.frame`, tibble, or
#' data.table. Internally converts to `data.table` (via
#' [ensure_dt_env()]) to compute efficiently, then restores the original
#' class on return.
#'
#' @param df A data frame, tibble, or data.table.
#' @param num_var Unquoted column name of the numeric variable to bin (e.g., `age`).
#' @param breaks Optional numeric vector of cut points **in ascending order**.
#'   If missing, breaks are generated from the observed range using
#'   `interval`. (Note: unlike [base::cut()], a single integer here is
#'   **not** interpreted as “number of bins”.)
#' @param interval Integer width for bins when `breaks` is missing. Default `5`.
#' @param right Logical; passed to [base::cut()]. If `TRUE`, intervals are
#'   right-closed `(a, b]`; if `FALSE` (default), left-closed `[a, b)`.
#' @param col_nm Optional name for the new band column. Default is
#'   `"<num_var>_band"`.
#' @param cutoff Logical; if `TRUE`, the upper break is forced to the
#'   observed maximum + 1, so the top band ends right above the maximum.
#'   Default `FALSE`.
#' @param label_style Label style for the band levels; one of:
#'   - `"close"` (default): `"a-b"` for all bands
#'   - `"open"`: first band `"-b"`, last band `"a-"`
#'   - `"open.start"`: first band `"-b"` only
#'   - `"open.end"`: last band `"a-"` only
#'
#' @return An object of the **same class as `df`**, with an ordered factor
#'   column appended (at `col_nm`). The new column is placed immediately
#'   after `num_var`.
#'
#' @details
#' - When `breaks` is missing, the function computes a lower bound at
#'   `floor(min(num_var)/interval) * interval` and an upper bound at
#'   `ceiling(max(num_var)/interval) * interval` (or `max(num_var)+1` if
#'   `cutoff = TRUE`), then sequences by `interval`.
#' - Labels are derived from the numeric breaks and formatted as
#'   `"a-b"` (with `b` shown as `b-1` when `right = FALSE`, to reflect
#'   `[a, b)` semantics). `"open"` variants replace the first/last label
#'   with `"-b"` / `"a-"` respectively.
#' - The resulting band column is an **ordered factor**.
#'
#' @examples
#' \donttest{
#' # Basic: 5-year bands, left-closed intervals
#' df <- mtcars
#' add_band(df, hp)
#'
#' # Custom width and right-closed intervals
#' add_band(df, mpg, interval = 10, right = TRUE, col_nm = "mpg_band")
#'
#' # Provide explicit cut points
#' add_band(df, disp, breaks = c(0, 150, 250, 400, 600))
#'
#' # Open-ended labels at both ends
#' add_band(df, qsec, interval = 2, label_style = "open")
#' }
#'
#' @export
add_band <- function(df, num_var, breaks, interval = 5, right = FALSE,
                     col_nm, cutoff = FALSE,
                     label_style = c("close", "open.start", "open", "open.end")) {
  assert_class(df, "data.frame")
  label_style <- match.arg(label_style)

  env <- ensure_dt_env(df)
  dt  <- env$dt

  num_var <- capture_names(dt, !!rlang::enquo(num_var))
  col <- dt[[num_var]]

  mn <- floor(min(col)/interval) * interval
  if (missing(breaks)) {
    if (!cutoff) {
      mx <- ceiling(max(col)/interval) * interval
      if (max(col) == mx)
        mx <- ceiling(max(col)/interval + 1) * interval
      breaks <- seq(mn, mx, interval)
    } else {
      mx <- max(col) + 1
      breaks <- seq(mn, mx, interval)
      breaks <- sort(unique(c(breaks, mx)))
      breaks <- breaks[breaks <= mx]
    }
  }
  col_band <- cut(col, breaks = breaks, right = right, dig.lab = 5)

  # build compact labels like "a-b"
  l <- levels(col_band)
  r <- gregexpr("[0-9]+", l, perl = TRUE)
  m <- regmatches(l, r)
  s <- as.integer(vapply(m, function(x) x[1L], character(1L)))
  e <- as.integer(vapply(m, function(x) x[2L], character(1L))) - 1
  labels <- sprintf("%d-%d", s, e)

  # apply open/closed label style
  if (label_style == "open") {
    labels[1L] <- get_pattern("-[0-9]+", labels[1L])
    labels[length(labels)] <- get_pattern("[0-9]+-", labels[length(labels)])
  } else if (label_style == "open.start") {
    labels[1L] <- get_pattern("-[0-9]+", labels[1L])
  } else if (label_style == "open.end") {
    labels[length(labels)] <- get_pattern("[0-9]+-", labels[length(labels)])
  }
  levels(col_band) <- labels
  if (missing(col_nm)) col_nm <- sprintf("%s_band", num_var)
  data.table::set(dt, j = col_nm, value = ordered(col_band))
  data.table::setcolorder(dt, col_nm, after = num_var)

  env$restore(dt[])
}

#' Infer band widths from ordered factor levels
#'
#' @description
#' Parse range-style factor levels and return the *width* (interval length) for
#' each band. This is useful when your bands are created by `add_band()` and
#' stored as an ordered factor.
#'
#' @details
#' Supported level formats are **strictly**:
#' \itemize{
#'   \item `-9` : open start band, interpreted as \eqn{(-\infty, 9]} → width = `Inf`
#'   \item `10-19` : closed band, interpreted as \eqn{[10, 19]} → width = `19 - 10 + 1 = 10`
#'   \item `20-` : open end band, interpreted as \eqn{[20, \infty)} → width = `Inf`
#' }
#'
#' Any other format (including spaces like `"10 - 19"`, non-integers, extra dashes,
#' etc.) will raise an error.
#'
#' @param x An **ordered factor** whose levels encode bands in one of the supported
#'   formats (`-9`, `10-19`, `20-`).
#'
#' @return A named numeric vector of the same length as `levels(x)`, where each
#'   element is the inferred band width. Open-ended bands return `Inf`.
#'
#' @examples
#' x <- ordered(c("-9", "10-19", "20-"), c("-9", "10-19", "20-"))
#' infer_band_widths(x)
#' #   -9 10-19   20-
#' #  Inf    10   Inf
#'
#' @export
infer_band_widths <- function(x) {
  if (!is.ordered(x))
    stop("`x` must be an ordered factor.")

  lvls <- levels(x)

  if (length(lvls) == 0L)
    return(stats::setNames(numeric(0L), character(0L)))

  ok <- grepl(
    "^(?:-(?:0|[1-9]\\d*)|(?:0|[1-9]\\d*)-(?:0|[1-9]\\d*)|(?:0|[1-9]\\d*)-)$",
    lvls
  )
  if (any(!ok)) {
    bad <- paste0("'", lvls[!ok], "'", collapse = ", ")
    stop(
      "Invalid level format: ", bad,
      ". Allowed formats: -9, 10-19, 20-"
    )
  }

  widths <- vapply(lvls, function(s) {
    # Open start
    if (grepl("^-(?:0|[1-9]\\d*)$", s)) {
      return(Inf)
    }

    # Open end
    if (grepl("^(?:0|[1-9]\\d*)-$", s)) {
      return(Inf)
    }

    # Closed range
    parts <- strsplit(s, "-", fixed = TRUE)[[1L]]
    start <- as.numeric(parts[1L])
    end   <- as.numeric(parts[2L])

    if (end < start)
      stop("Invalid range (end < start): '", s, "'.")

    end - start + 1
  }, numeric(1L))

  names(widths) <- lvls
  widths
}

#' Deprecated: set_band()
#'
#' @description
#' `r lifecycle::badge("deprecated")`
#'
#' Use [add_band()] instead.
#'
#' @param ... Additional arguments passed to [add_band()].
#'
#' @return Same return value as [add_band()].
#'
#' @seealso [add_band()]
#'
#' @export
set_band <- function(...) {
  lifecycle::deprecate_warn("0.0.0.9000", "set_band()", "add_band()")
  add_band(...)
}


#' Split a table into prefix tables
#'
#' Splits a table into an ordered list of prefix tables based on `by`.
#' Each prefix table is created by sequentially adding blocks defined by `by`.
#'
#' @inheritParams split_suffix_tables
#'
#' @return A list of data.tables.
#'
#' @examples
#' x <- data.frame(
#'   gender = c("M", "F", "M", "F", "M", "F"),
#'   age_band = factor(
#'     c("20-29", "20-29", "30-39", "30-39", "40-49", "40-49"),
#'     levels = c("20-29", "30-39", "40-49"),
#'     ordered = TRUE
#'   ),
#'   value = 1:6
#' )
#'
#' # prefix tables by age band
#' split_prefix_tables(x, by = "age_band")
#'
#' # prefix tables computed separately by gender
#' split_prefix_tables(x, by = "age_band", split_by = "gender")
#'
#' @export
split_prefix_tables <- function(x, by, split_by = NULL, min_nrow = 1L) {
  .split_ordered_tables(
    x = x,
    by = by,
    split_by = split_by,
    type = "prefix",
    min_nrow = min_nrow
  )
}

#' Split a table into suffix tables
#'
#' Splits a table into an ordered list of suffix tables based on `by`.
#' Each suffix table is created by sequentially removing the earliest
#' block defined by `by`.
#'
#' If `split_by` is supplied, the operation is performed separately within
#' each group and the resulting tables are concatenated.
#'
#' @param x A data.frame or data.table.
#' @param by Character vector of column names defining ordered blocks.
#' @param split_by Optional character vector of column names used to split
#'   `x` before generating suffix tables.
#' @param min_nrow Integer; minimum number of rows required to keep a
#'   generated table. Default `1L`.
#'
#' @return A list of data.tables.
#'
#' @examples
#' x <- data.frame(
#'   gender = c("M", "F", "M", "F", "M", "F"),
#'   age_band = factor(
#'     c("20-29", "20-29", "30-39", "30-39", "40-49", "40-49"),
#'     levels = c("20-29", "30-39", "40-49"),
#'     ordered = TRUE
#'   ),
#'   value = 1:6
#' )
#'
#' # suffix tables by age band
#' split_suffix_tables(x, by = "age_band")
#'
#' # suffix tables computed separately by gender
#' split_suffix_tables(x, by = "age_band", split_by = "gender")
#'
#' # keep only tables with at least 2 rows
#' split_suffix_tables(x, by = "age_band", min_nrow = 2L)
#'
#' @export
split_suffix_tables <- function(x, by, split_by = NULL, min_nrow = 1L) {
  .split_ordered_tables(
    x = x,
    by = by,
    split_by = split_by,
    type = "suffix",
    min_nrow = min_nrow
  )
}

#' Split a table into ordered prefix/suffix tables (internal)
#'
#' @keywords internal
.split_ordered_tables <- function(x, by, split_by = NULL,
                                  type = c("suffix", "prefix"),
                                  min_nrow = 1L) {
  type <- match.arg(type)

  x <- data.table::as.data.table(x)

  if (!length(by))
    stop("`by` must contain at least one column name.", call. = FALSE)

  if (!all(by %in% names(x))) {
    bad <- by[!by %in% names(x)]
    stop(
      sprintf("Unknown `by` column(s): %s", paste(bad, collapse = ", ")),
      call. = FALSE
    )
  }

  if (!is.null(split_by) && !all(split_by %in% names(x))) {
    bad <- split_by[!split_by %in% names(x)]
    stop(
      sprintf("Unknown `split_by` column(s): %s", paste(bad, collapse = ", ")),
      call. = FALSE
    )
  }

  if (!is.numeric(min_nrow) || length(min_nrow) != 1L || is.na(min_nrow)) {
    stop("`min_nrow` must be a single non-missing number.", call. = FALSE)
  }
  min_nrow <- as.integer(min_nrow)
  if (min_nrow < 0L) {
    stop("`min_nrow` must be >= 0.", call. = FALSE)
  }

  groups <- if (is.null(split_by)) {
    list(x)
  } else {
    split(x, by = split_by)
  }

  out <- lapply(groups, function(dt) {
    data.table::setorderv(dt, by)

    blocks <- split(dt, by = by)
    blocks <- Filter(function(z) nrow(z) > 0L, blocks)

    n <- length(blocks)
    if (n == 0L) return(list())

    res <- if (type == "suffix") {
      lapply(seq_len(n), function(i) {
        data.table::rbindlist(blocks[i:n], use.names = TRUE)
      })
    } else {
      lapply(seq_len(n), function(i) {
        data.table::rbindlist(blocks[1:i], use.names = TRUE)
      })
    }

    Filter(function(z) nrow(z) >= min_nrow, res)
  })

  unlist(out, recursive = FALSE)
}

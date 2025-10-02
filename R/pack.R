#' Pack non-missing values to the left across multiple columns
#'
#' Shift non-`NA` values leftwards across a set of columns, replacing
#' the vacated positions with type-consistent `NA`s. Useful for
#' left-packing code columns (e.g., `kcd1`, `kcd2`, ...), so that
#' values always occupy the earliest available column.
#'
#' Works in-place when `df` is a data.table. For plain data.frame,
#' a temporary data.table is used under the hood and the modified
#' object is returned (no mutation of the original).
#'
#' @param df A data.frame or data.table.
#' @param cols Columns to pack (tidy-select compatible or a character vector).
#'
#' @details
#' - Columns in `cols` are expected to share the same type when they contain
#'   non-missing values. If a column is entirely `NA` (type ambiguous), it will
#'   adopt the source column's type at the moment a non-`NA` value is shifted in.
#' - The algorithm performs `(length(cols) - 1)` passes, each time moving any
#'   available value one step to the left, until no more shifts are possible.
#' - Type safety is enforced; incompatible source/destination types stop with
#'   an informative error to prevent silent coercion.
#'
#' @return The modified object, invisibly. If `df` is a data.table, it is
#'   modified by reference; if it is a data.frame, a modified copy is returned.
#'
#' @examples
#' \dontrun{
#' library(data.table)
#'
#' df <- data.table::data.table(
#'   kcd1 = c(NA, NA, NA),
#'   kcd2 = c("W49","S925", NA),
#'   kcd3 = c("S925", NA, "D12")
#' )
#' pack_left(df, c(kcd1, kcd2, kcd3))
#' df
#' #    kcd1  kcd2  kcd3
#' # 1:  W49  S925  <NA>
#' # 2: S925  <NA>  <NA>
#' # 3:  D12  <NA>  <NA>
#' }
#'
#' @export
pack_left <- function(df, cols) {
  assert_class(df, "data.frame")

  env <- ensure_dt_env(df)
  dt  <- env$dt

  cols <- capture_names(dt, !!rlang::enquo(cols))

  p <- length(cols)
  if (p <= 1L)
    return(env$restore(dt))

  # Pre-check: all non-NA columns must share the same type
  type_keys <- vapply(cols, function(j) .col_type_key(dt[[j]]), character(1))
  observed  <- unique(na.omit(type_keys))
  if (length(observed) > 1L) {
    stop(sprintf(
      "All non-NA columns must share the same type; observed: %s",
      paste(observed, collapse = ", ")
    ))
  }

  # Perform (p-1) passes to allow multi-step shifts
  for (pass in seq_len(p - 1L)) {
    for (j in 2L:p) {
      src <- cols[j]       # source column (to the right)
      dst <- cols[j - 1L]  # destination column (to the left)

      src_col <- dt[[src]]
      dst_col <- dt[[dst]]

      # Rows where dst is NA and src has a non-NA value
      i <- is.na(dst_col) & !is.na(src_col)
      if (!any(i)) next

      # If dst is completely NA (ambiguous type), initialize with src type
      if (all(is.na(dst_col))) {
        data.table::set(dt, j = dst,
                        value = rep(.na_like(src_col), nrow(dt)))
        dst_col <- dt[[dst]]
      }

      # Type safety check: dst and src must be compatible
      dst_key <- .col_type_key(dst_col)
      src_key <- .col_type_key(src_col)
      if (!is.na(dst_key) && !is.na(src_key) && !identical(dst_key, src_key)) {
        stop(sprintf("Type mismatch between '%s' (%s) and '%s' (%s).",
                     dst, dst_key, src, src_key))
      }

      # Move non-NA values one step left
      data.table::set(dt, i = which(i), j = dst, value = src_col[i])

      # Overwrite src positions with NA of the correct type
      data.table::set(dt, i = which(i), j = src,
                      value = rep(.na_like(src_col), sum(i)))
    }
  }

  env$restore(dt)
}

# Internal helper functions -----------------------------------------------

#' Internal: return a simplified type key
#'
#' Produces a type identifier string for a vector, ignoring all-`NA` cases.
#' Used internally by [pack_left()] to check type compatibility between columns.
#'
#' @param x A vector.
#'
#' @return A character string (e.g., `"integer"`, `"numeric"`, `"Date"`, `"POSIXct"`)
#'   or `NA_character_` if the column is entirely `NA`.
#'
#' @keywords internal
.col_type_key <- function(x) {
  if (all(is.na(x))) return(NA_character_)
  cls <- class(x)[1L]
  if (cls %in% c("Date", "POSIXct")) cls else typeof(x)
}

#' Internal: return a type-consistent NA value
#'
#' Creates a length-1 `NA` matching the type/class of `x`. This prevents
#' unintended column coercion when assigning missing values via
#' [data.table::set()].
#'
#' @param x A vector whose type/class determines the NA returned.
#'
#' @return A length-1 `NA` of the same type as `x`.
#'
#' @keywords internal
.na_like <- function(x) {
  if (is.integer(x)) {
    NA_integer_
  } else if (is.numeric(x)) {
    NA_real_
  } else if (is.logical(x)) {
    NA
  } else if (inherits(x, "Date")) {
    as.Date(NA)
  } else if (inherits(x, "POSIXct")) {
    as.POSIXct(NA)
  } else if (is.factor(x)) {
    # keep factor class; NA with same levels
    factor(NA, levels = levels(x), ordered = is.ordered(x))
  } else {
    NA_character_
  }
}

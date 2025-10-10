#' Replace NA with zero (in place, type-preserving)
#'
#' Efficiently replaces `NA` values with `0` in selected numeric columns
#' of a data.table. The replacement is done *in place* using
#' [data.table::set()], modifying only rows where `NA` values exist.
#'
#' @param DT A data.table (modified in place).
#' @param cols Columns to target, passed to [capture_names()].
#'   Can be specified as unquoted names (e.g., `c(x, y)`),
#'   a character vector (e.g., `c("x","y")`), or integer indices (e.g., `c(1,2)`).
#'   If missing, all numeric columns are used.
#'
#' @return The modified data.table, returned invisibly for chaining.
#'
#' @details
#' - Only **numeric** columns are targeted by default.
#' - Columns are updated **in place** only at the affected row indices.
#' - If `cols` includes non-numeric columns, they are rejected with an error.
#'
#' @examples
#' \dontrun{
#' library(data.table)
#'
#' dt <- data.table::data.table(
#'   x = c(1, NA, 3),
#'   y = c(NA, 5, NA),
#'   z = c("a", NA, "c")
#' )
#'
#' # Replace NA with 0 in numeric columns
#' replace_na_with_zero(dt)
#' print(dt)
#'
#' # Replace NA with 0 in specific columns only
#' replace_na_with_zero(dt, cols = c("x"))
#' print(dt)
#' }
#'
#' @export
replace_na_with_zero <- function(DT, cols) {
  assert_class(DT, "data.table")

  if (missing(cols)) {
    cols <- .get_numeric_cols(DT)
  } else {
    cols_captured <- capture_names(DT, !!rlang::enquo(cols))
    cols_numeric  <- .get_numeric_cols(DT)
    cols_non_numeric <- setdiff(cols_captured, cols_numeric)
    if (length(cols_non_numeric)) {
      stop(sprintf(
        "Non-numeric column(s) not allowed: %s",
        paste(cols_non_numeric, collapse = ", ")
      ), call. = FALSE)
    }
    cols <- cols_captured
  }

  if (!length(cols))
    return(invisible(DT[]))

  for (j in cols) {
    i <- which(is.na(DT[[j]]))
    if (length(i))
      data.table::set(DT, i = i, j = j, value = 0)
  }

  invisible(DT[])
}

#' Replace zero with NA (in place, type-preserving)
#'
#' Replaces `0` with `NA` in selected **numeric** columns of a data.table.
#' Uses [data.table::set()] to modify only affected rows *in place*.
#'
#' @param DT A data.table (modified in place).
#' @param cols Columns to target, passed to [capture_names()].
#'   Can be specified as unquoted names (e.g., `c(x, y)`),
#'   a character vector (e.g., `c("x","y")`), or integer indices (e.g., `c(1,2)`).
#'   If missing, all numeric columns are used.
#'
#' @return The modified data.table, returned invisibly.
#'
#' @details
#' - Only **numeric** columns are targeted by default.
#' - Columns are updated **in place** only at the affected row indices.
#' - If `cols` includes non-numeric columns, they are rejected with an error.
#'
#' @details
#' - Only **numeric** columns are targeted by default.
#' - Columns are updated **in place** only at the affected row indices.
#' - If `cols` includes non-numeric columns, they are rejected with an error.
#'
#' @examples
#' \dontrun{
#' library(data.table)
#' dt <- data.table::data.table(x = c(0,1,0), y = c(2,0,3), z = c("a","b","c"))
#' replace_zero_with_na(dt)
#' print(dt)
#' }
#'
#' @export
replace_zero_with_na <- function(DT, cols) {
  assert_class(DT, "data.table")

  if (missing(cols)) {
    cols <- .get_numeric_cols(DT)
  } else {
    cols_captured <- capture_names(DT, !!rlang::enquo(cols))
    cols_numeric  <- .get_numeric_cols(DT)
    cols_non_numeric <- setdiff(cols_captured, cols_numeric)
    if (length(cols_non_numeric)) {
      stop(sprintf(
        "Non-numeric column(s) not allowed: %s",
        paste(cols_non_numeric, collapse = ", ")
      ), call. = FALSE)
    }
    cols <- cols_captured
  }

  if (!length(cols))
    return(invisible(DT[]))

  for (j in cols) {
    i <- which(DT[[j]] == 0)
    if (length(i))
      data.table::set(DT, i = i, j = j, value = NA)
  }

  invisible(DT[])
}

#' Replace empty strings with NA (in place)
#'
#' Replaces `""` with `NA_character_` in selected **character** columns.
#' Uses [data.table::set()] for in-place updates.
#'
#' @param DT A data.table (modified in place).
#' @param cols Columns to target, passed to [capture_names()].
#'   Can be specified as unquoted names (e.g., `c(x, y)`),
#'   a character vector (e.g., `c("x","y")`), or integer indices (e.g., `c(1,2)`).
#'   If missing, all character columns are used.
#'
#' @return The modified data.table, returned invisibly.
#'
#' @details
#' - Only **character** columns are targeted by default.
#' - Empty strings (`""`) are converted to `NA_character_` **in place**.
#' - If `cols` includes non-character columns, they are rejected with an error.
#'
#' @examples
#' \dontrun{
#' library(data.table)
#' dt <- data.table(x = c("a","",NA), y = c("", "b", "c"))
#' print(dt)
#' replace_empty_with_na(dt)
#' print(dt)
#' }
#'
#' @export
replace_empty_with_na <- function(DT, cols) {
  assert_class(DT, "data.table")

  if (missing(cols)) {
    cols <- .get_character_cols(DT)
  } else {
    cols_captured  <- capture_names(DT, !!rlang::enquo(cols))
    cols_character <- .get_character_cols(DT)
    cols_non_character <- setdiff(cols_captured, cols_character)
    if (length(cols_non_character)) {
      stop(sprintf(
        "Non-character column(s) not allowed: %s",
        paste(cols_non_character, collapse = ", ")
      ), call. = FALSE)
    }
    cols <- cols_captured
  }

  if (!length(cols))
    return(invisible(DT[]))

  for (j in cols) {
    i <- which(DT[[j]] == "")
    if (length(i))
      data.table::set(DT, i = i, j = j, value = NA_character_)
  }

  invisible(DT[])
}

#' Replace NA with empty strings (in place)
#'
#' Replaces `NA_character_` with `""` in selected **character** columns.
#' Uses [data.table::set()] for in-place updates.
#'
#' @param DT A data.table (modified in place).
#' @param cols Columns to target, passed to [capture_names()].
#'   Can be specified as unquoted names (e.g., `c(x, y)`),
#'   a character vector (e.g., `c("x","y")`), or integer indices (e.g., `c(1,2)`).
#'   If missing, all character columns are used.
#'
#' @return The modified data.table, returned invisibly.
#'
#' @details
#' - Only **character** columns are targeted by default.
#' - `NA_character_` values are converted to `""` **in place**.
#' - If `cols` includes non-character columns, they are rejected with an error.
#'
#' @examples
#' \dontrun{
#' library(data.table)
#' dt <- data.table::data.table(x = c("a", NA, "c"), y = c(NA, "b", NA))
#' replace_na_with_empty(dt)
#' print(dt)
#' }
#'
#' @export
replace_na_with_empty <- function(DT, cols) {
  assert_class(DT, "data.table")

  if (missing(cols)) {
    cols <- .get_character_cols(DT)
  } else {
    cols_captured  <- capture_names(DT, !!rlang::enquo(cols))
    cols_character <- .get_character_cols(DT)
    cols_non_character <- setdiff(cols_captured, cols_character)
    if (length(cols_non_character)) {
      stop(sprintf(
        "Non-character column(s) not allowed: %s",
        paste(cols_non_character, collapse = ", ")
      ), call. = FALSE)
    }
    cols <- cols_captured
  }

  if (!length(cols))
    return(invisible(DT[]))

  for (j in cols) {
    i <- which(is.na(DT[[j]]))
    if (length(i))
      data.table::set(DT, i = i, j = j, value = "")
  }

  invisible(DT[])
}

#' Replace a specific string with another (in place)
#'
#' Replaces occurrences of string `a` with string `b` across selected
#' **character** columns, *in place* using [data.table::set()].
#'
#' @param DT A data.table (modified in place).
#' @param cols Columns to target, passed to [capture_names()].
#'   Can be specified as unquoted names (e.g., `c(x, y)`),
#'   a character vector (e.g., `c("x","y")`), or integer indices (e.g., `c(1,2)`).
#'   If missing, all character columns are used.
#' @param a String to be replaced.
#' @param b Replacement string.
#'
#' @return The modified data.table, returned invisibly.
#'
#' @details
#' - Only **character** columns are targeted by default.
#' - Exact matches of `a` are replaced with `b` **in place** (no regex).
#' - `NA` values are left unchanged.
#' - If `cols` includes non-character columns, they are rejected with an error.
#'
#' @examples
#' \dontrun{
#' library(data.table)
#' dt <- data.table::data.table(x = c("A","B","A"), y = c("C","A","C"))
#' replace_a_with_b(dt, a = "A", b = "Z")
#' print(dt)
#' }
#'
#' @export
replace_a_with_b <- function(DT, cols, a, b) {
  assert_class(DT, "data.table")

  if (missing(cols)) {
    cols <- .get_character_cols(DT)
  } else {
    cols_captured  <- capture_names(DT, !!rlang::enquo(cols))
    cols_character <- .get_character_cols(DT)
    cols_non_character <- setdiff(cols_captured, cols_character)
    if (length(cols_non_character)) {
      stop(sprintf(
        "Non-character column(s) not allowed: %s",
        paste(cols_non_character, collapse = ", ")
      ), call. = FALSE)
    }
    cols <- cols_captured
  }

  if (!length(cols))
    return(invisible(DT[]))

  for (j in cols) {
    i <- which(DT[[j]] == a)
    if (length(i))
      data.table::set(DT, i = i, j = j, value = b)
  }

  invisible(DT[])
}

#' Replace a column vector in a matrix (in place)
#'
#' Replace one or more columns of a matrix with a new vector, directly
#' modifying the matrix memory. This avoids reallocation and can be more
#' efficient for large matrices.
#'
#' @param mat A numeric matrix (modified in place).
#' @param cols Column(s) to replace; either character names or integer indices.
#' @param vec A numeric vector to insert in place of the selected columns.
#'
#' @return No return value, called for side effects (the matrix is updated in place).
#'
#' @examples
#' \dontrun{
#' x <- matrix(as.numeric(1:9), 3, 3)
#' replace_cols_in_mat(x, c(1, 3), c(100, 200, 300)) # replace 1st and 3rd columns
#' }
#'
#' @export
replace_cols_in_mat <- function(mat, cols, vec) {
  if (is.character(cols)) cols <- index_cols(mat, col)
  invisible(.Call(ReplaceColsInMat, mat, cols, vec))
}

#' Trim leading/trailing whitespace (in place)
#'
#' Removes leading and trailing whitespace from selected **character** columns.
#' Vectorized via `gsub()` per column; assigns results *in place*.
#'
#' @param DT A data.table (modified in place).
#' @param cols Columns to target, passed to [capture_names()].
#'   Can be specified as unquoted names (e.g., `c(x, y)`),
#'   a character vector (e.g., `c("x","y")`), or integer indices (e.g., `c(1,2)`).
#'   If missing, all character columns are used.
#' @param ws A regex describing whitespace (default `"[ \\t\\r\\n]"`).
#'
#' @return The modified data.table, returned invisibly.
#'
#' @details
#' - Only **character** columns are targeted by default.
#' - Leading/trailing whitespace (as defined by `ws`) is trimmed **in place**.
#' - `NA` values are preserved as `NA`.
#' - If `cols` includes non-character columns, they are rejected with an error.
#'
#' @examples
#' \dontrun{
#' library(data.table)
#' dt <- data.table(x = c(" a", "b ", " c "), y = c("d", " e", "f "))
#' trim_ws(dt)
#' print(dt)
#' }
#'
#' @export
trim_ws <- function(DT, cols, ws = "[ \t\r\n]") {
  assert_class(DT, "data.table")

  if (missing(cols)) {
    cols <- .get_character_cols(DT)
  } else {
    cols_captured  <- capture_names(DT, !!rlang::enquo(cols))
    cols_character <- .get_character_cols(DT)
    cols_non_character <- setdiff(cols_captured, cols_character)
    if (length(cols_non_character)) {
      stop(sprintf(
        "Non-character column(s) not allowed: %s",
        paste(cols_non_character, collapse = ", ")
      ), call. = FALSE)
    }
    cols <- cols_captured
  }

  if (!length(cols)) return(invisible(DT[]))

  re <- sprintf("^%s+|%s+$", ws, ws)
  for (j in cols) {
    v <- DT[[j]]
    # keep NA as NA; only trim non-NA strings
    v2 <- ifelse(is.na(v), v, gsub(re, "", v, perl = TRUE))
    if (!identical(v, v2)) data.table::set(DT, j = j, value = v2)
  }
  invisible(DT[])
}

#' Remove punctuation (in place)
#'
#' Removes punctuation from selected **character** columns using a regex.
#' Default pattern keeps asterisks: `pattern = "(?!\\\\*)[[:punct:]]"`.
#'
#' @param DT A data.table (modified in place).
#' @param cols Columns to target, passed to [capture_names()].
#'   Can be specified as unquoted names (e.g., `c(x, y)`),
#'   a character vector (e.g., `c("x","y")`), or integer indices (e.g., `c(1,2)`).
#'   If missing, all character columns are used.
#' @param pattern Regex of punctuation to remove.
#'
#' @return The modified data.table, returned invisibly.
#'
#' @details
#' - Only **character** columns are targeted by default.
#' - Punctuation matched by `pattern` is removed **in place**; `NA` is preserved.
#' - The default pattern keeps asterisks: `pattern = "(?!\\\\*)[[:punct:]]"`.
#' - If `cols` includes non-character columns, they are rejected with an error.
#'
#' @examples
#' \dontrun{
#' library(data.table)
#' dt <- data.table(x = c("A3-", "$*_B", "C+*&"), y = c("12-3", "R&", "4q_++"))
#' rm_punct(dt)
#' print(dt)
#' }
#'
#' @export
rm_punct <- function(DT, cols, pattern = "(?!\\*)[[:punct:]]") {
  assert_class(DT, "data.table")

  if (missing(cols)) {
    cols <- .get_character_cols(DT)
  } else {
    cols_captured  <- capture_names(DT, !!rlang::enquo(cols))
    cols_character <- .get_character_cols(DT)
    cols_non_character <- setdiff(cols_captured, cols_character)
    if (length(cols_non_character)) {
      stop(sprintf(
        "Non-character column(s) not allowed: %s",
        paste(cols_non_character, collapse = ", ")
      ), call. = FALSE)
    }
    cols <- cols_captured
  }

  if (!length(cols))
    return(invisible(DT[]))

  for (j in cols) {
    v <- DT[[j]]
    value <- ifelse(is.na(v), v, gsub(pattern, "", v, perl = TRUE))
    if (!identical(v, value))
      data.table::set(DT, j = j, value = value)
  }

  invisible(DT[])
}

#' Remove columns by reference (in place)
#'
#' Removes one or more columns from a data.table **by reference** (no copy)
#'
#' @param DT A data.table (modified in place).
#' @param cols Columns to remove, passed to [capture_names()].
#'   Can be specified as unquoted names (e.g., `c(x, y)`),
#'   a character vector (e.g., `c("x","y")`), or integer indices (e.g., `c(1,2)`).
#'
#' @return The modified data.table, returned invisibly.
#'
#' @details
#' - Columns are removed **by reference** (no copy).
#' - `cols` may be provided as unquoted names, character vector, or integer indices.
#' - Silently does nothing if `cols` is empty after resolution.
#'
#' @examples
#' \dontrun{
#' library(data.table)
#'
#' dt <- data.table::as.data.table(mtcars)
#' rm_cols(dt, .(mpg, cyl))
#' head(dt)
#' }
#'
#' @export
rm_cols <- function(DT, cols) {
  assert_class(DT, "data.table")
  cols <- capture_names(DT, !!rlang::enquo(cols))
  if (!length(cols)) return(invisible(DT[]))
  DT[, `:=`((cols), NULL)]
  invisible(DT[])
}


# Internal helper functions -----------------------------------------------

.get_numeric_cols <- function(df) {
  names(df)[vapply(df, is.numeric, logical(1L))]
}

.get_character_cols <- function(df) {
  names(df)[vapply(df, is.character, logical(1L))]
}

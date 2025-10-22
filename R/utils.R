#' Normalize ... arguments into a clean list
#'
#' Collects and standardizes inputs from `...`, ensuring that the result
#' is always a list of objects. If a single list was passed, it is unwrapped
#' once unless it is a data.frame, data.table, or tibble.
#'
#' @param ... Objects or lists of objects.
#'
#' @return A list of objects.
#'
#' @examples
#' \donttest{
#' normalize_dots(1, 2)
#' normalize_dots(a = 1, b = 2)
#' normalize_dots(list(a = 1, b = 2))
#' normalize_dots(data.frame(x = 1))  # protected, not unwrapped
#' normalize_dots(head(iris, 1), head(mtcars, 1))
#' normalize_dots(iris = head(iris, 1), mtcars = head(mtcars, 1))
#' normalize_dots(list(iris = head(iris, 1), mtcars = head(mtcars, 1)))
#' }
#'
#' @export
normalize_dots <- function(...) {
  dots <- list(...)
  if (length(dots) == 1L && is.list(dots[[1L]]) &&
      !inherits(dots[[1L]], c("data.frame", "data.table"))) {
    return(dots[[1L]])
  }
  dots
}

#' Prepend class(es) without duplication
#'
#' Ensure one or more class names appear at the front of an object's class
#' vector, without duplicating existing classes. If `x` is a data.table,
#' uses `data.table::setattr()` to avoid copies and preserve selfref.
#'
#' @param x An R object.
#' @param classes A character vector of class names to prepend.
#'
#' @return The input object `x` with its class attribute updated to have
#'   `classes` at the front, followed by its original classes (excluding
#'   any duplicates).
#'
#' @examples
#' \donttest{
#' df <- data.frame(a = 1)
#'
#' # Prepend a single class
#' df <- prepend_class(df, "new_class")
#'
#' # Prepend multiple classes
#' df <- prepend_class(df, c("new_class1", "new_class2"))
#' }
#'
#' @export
prepend_class <- function(x, classes) {
  # Fast returns / validation
  if (length(classes) == 0L) return(x)
  if (!is.character(classes) || anyNA(classes))
    stop("`classes` must be a non-empty character vector without NA.",
         call. = FALSE)

  # Normalize inputs
  classes   <- unique(classes)                    # avoid duplicates in input
  org_class <- setdiff(unname(class(x)), classes) # keep original order, remove overlaps

  # Data.table-friendly path: no copy, no selfref warning
  if (inherits(x, "data.table")) {
    data.table::setattr(x, "class", c(classes, org_class))
    return(x[])
  }

  # Fallback for non-data.table objects
  class(x) <- c(classes, org_class)
  x
}

#' Remove one or more classes from an object
#'
#' Safely removes specified class name(s) from an object's class vector.
#' Mirrors the internal logic and style of [prepend_class()], but
#' removes the specified `classes` instead of prepending them.
#'
#' @param x An object with a class attribute.
#' @param classes Character vector of class names to remove.
#'
#' @return The same object `x`, with the specified class(es) removed.
#'
#' @examples
#' x <- prepend_class(mtcars, c("custom", "tag"))
#' class(x)
#' # [1] "custom" "tag" "data.frame"
#'
#' y <- remove_class(x, "custom")
#' class(y)
#' # [1] "tag" "data.frame"
#'
#' z <- remove_class(y, c("tag", "data.frame"))
#' class(z)
#' # [1] "list"
#'
#' @export
remove_class <- function(x, classes) {
  # Fast returns / validation
  if (length(classes) == 0L)
    return(x)
  if (!is.character(classes) || anyNA(classes))
    stop("`classes` must be a non-empty character vector without NA.",
         call. = FALSE)

  # Normalize inputs
  classes <- unique(classes)
  remain  <- setdiff(unname(class(x)), classes)

  # Data.table-friendly path: no copy, no selfref warning
  if (inherits(x, "data.table")) {
    data.table::setattr(x, "class", remain)
    return(x[])
  }

  # Fallback for non-data.table objects
  class(x) <- remain
  x
}

#' Update class attribute (remove then prepend)
#'
#' Convenience wrapper that combines [remove_class()] and [prepend_class()]:
#' first removes specified class names, then prepends new ones (without
#' duplication). Works seamlessly with ordinary objects and `data.table`
#' (no-copy path).
#'
#' @param x An R object.
#' @param remove Character vector of class names to remove. Ignored if `NULL` or length 0.
#' @param prepend Character vector of class names to prepend. Ignored if `NULL` or length 0.
#'
#' @return The input object `x`, with its class attribute updated accordingly.
#'
#' @examples
#' \donttest{
#' x <- prepend_class(mtcars, c("old", "tag"))
#' class(x)
#' # [1] "old" "tag" "data.frame"
#'
#' x <- update_class(x, remove = "old", prepend = c("new", "hot"))
#' class(x)
#' # [1] "new" "hot" "tag" "data.frame"
#' }
#'
#' @export
update_class <- function(x, remove = NULL, prepend = NULL) {
  # Fast path: nothing to do
  if ((is.null(remove)  || length(remove)  == 0L) &&
      (is.null(prepend) || length(prepend) == 0L)) {
    return(x)
  }

  # Remove first (if requested)
  if (!is.null(remove)) {
    if (!is.character(remove) || anyNA(remove))
      stop("`remove` must be a character vector without NA.", call. = FALSE)
    x <- remove_class(x, remove)
  }

  # Then prepend (if requested)
  if (!is.null(prepend)) {
    if (!is.character(prepend) || anyNA(prepend))
      stop("`prepend` must be a character vector without NA.", call. = FALSE)
    x <- prepend_class(x, prepend)
  }

  x
}

#' Index columns by name
#'
#' Return column indices for the given column names from a data.frame or matrix.
#'
#' @param x A data.frame or matrix.
#' @param cols A character vector of column names (in desired order).
#' @param all_matches Logical; if `TRUE`, return indices of *all* columns whose
#'   names match each `cols` entry (duplicates allowed). If `FALSE` (default),
#'   return only the first match for each name.
#' @param error_on_missing Logical; if `TRUE` (default), error when any of
#'   `cols` are not present in `colnames(x)`. If `FALSE`, drop missing names.
#'
#' @return
#' Integer vector of column indices. Names of the vector correspond to the
#' requested `cols` that were successfully matched.
#'
#' @examples
#' # data.frame
#' index_cols(mtcars, c("disp", "drat", "qsec"))
#'
#' # drop missing columns silently
#' index_cols(mtcars, c("disp", "drat", "nope"), error_on_missing = FALSE)
#'
#' # matrix
#' mat <- as.matrix(mtcars)
#' index_cols(mat, c("disp", "hp"))
#'
#' # duplicated column names (all matches)
#' df <- data.frame(a = 1, a = 2, b = 3, check.names = FALSE)
#' index_cols(df, "a", all_matches = TRUE)
#'
#' @export
index_cols <- function(x, cols, all_matches = FALSE, error_on_missing = TRUE) {
  assert_class(x, c("data.frame", "matrix"))

  if (!is.character(cols) || length(cols) < 1L || anyNA(cols)) {
    rlang::abort(
      message = "`cols` must be a non-empty character vector without NA.",
      class   = c("instead_error_invalid_argument", "instead_error", "error"),
      arg     = "cols"
    )
  }

  nms <- colnames(x)
  if (is.null(nms)) {
    rlang::abort(
      message = "`x` has no column names.",
      class   = c("instead_error_no_colnames", "instead_error", "error")
    )
  }

  if (!all_matches) {
    idx0    <- match(cols, nms)
    na_mask <- is.na(idx0)

    if (any(na_mask) && error_on_missing) {
      rlang::abort(
        message = sprintf("Missing column(s): %s",
                          paste(cols[na_mask], collapse = ", ")),
        class   = c("instead_error_missing_columns", "instead_error", "error"),
        missing = cols[na_mask],
        columns = nms
      )
    }

    idx  <- idx0[!na_mask]
    kept <- cols[!na_mask]
    names(idx) <- kept
    return(idx)
  }

  # all_matches = TRUE — collect all hits for each requested name
  out_idx  <- integer(0)
  out_name <- character(0)

  for (nm in cols) {
    hits <- which(nms == nm)
    if (length(hits) == 0L) {
      if (error_on_missing) {
        rlang::abort(
          message = sprintf("Missing column: %s", nm),
          class   = c("instead_error_missing_columns", "instead_error", "error"),
          missing = nm,
          columns = nms
        )
      } else {
        next
      }
    }
    out_idx  <- c(out_idx, hits)
    out_name <- c(out_name, rep.int(nm, length(hits)))
  }

  names(out_idx) <- out_name
  out_idx
}


#' Match columns
#'
#' Return column names from a data frame that match a specified set.
#'
#' @param df A data.frame.
#' @param cols A character vector of column names.
#'
#' @return A character vector of matched column names (non-matching entries are dropped).
#'
#' @examples
#' \donttest{
#' df <- data.frame(x = c(1, 2, 3), y = c("A", "B", "C"), z = c(4, 5, 6))
#' match_cols(df, c("x", "z"))
#' }
#'
#' @export
match_cols <- function(df, cols) {
  assert_class(df, "data.frame")
  nms <- names(df)
  nms[match(cols, nms, 0L)]
}

#' Find columns by regular expression
#'
#' Return column names that match a regular expression.
#'
#' @param df A data.frame.
#' @param pattern A character string containing a regular expression.
#'
#' @return A character vector of matching column names.
#'
#' @examples
#' \donttest{
#' df <- data.frame(col_a = c(1, 2, 3), col_b = c("A", "B", "C"), col_c = c(4, 5, 6))
#' regex_cols(df, pattern = c("a|c"))
#' }
#'
#' @export
regex_cols <- function(df, pattern) {
  assert_class(df, "data.frame")
  nms <- names(df)
  nms[grepl(pattern, nms, perl = TRUE)]
}

#' Columns not in a set
#'
#' Return the columns of a data frame that are **not** in a specified list.
#'
#' @param df A data.frame.
#' @param cols A character vector of column names.
#'
#' @return A character vector of column names in `df` but not in `cols`.
#'
#' @examples
#' \donttest{anti_cols(mtcars, c("mpg", "cyl", "disp", "hp", "drat"))}
#'
#' @export
anti_cols <- function(df, cols) {
  assert_class(df, "data.frame")
  setdiff(names(df), cols)
}

#' Find matching attributes
#'
#' Returns the names of attributes in `x` that match the specified names.
#'
#' @param x Any R object (e.g., `list`, `data.frame`, `data.table`).
#' @param name Character vector of attribute names to match.
#'
#' @return A character vector of matching attribute names.
#'
#' @examples
#' \donttest{
#' # Find attributes by name
#' match_attr(iris, c("class", "names"))
#' }
#'
#' @export
match_attr <- function(x, name)
  names(attributes(x))[match(name, names(attributes(x)), 0L)]

#' Find attributes by regular expression
#'
#' Return the names of attributes in `x` whose names match a regular expression pattern.
#'
#' @param x Any R object (e.g., `list`, `data.frame`, `data.table`).
#' @param name A character string containing a regular expression pattern
#'   to match against attribute names.
#'
#' @return A character vector of matching attribute names.
#'
#' @examples
#' \donttest{
#' # Find attributes by regex
#' regex_attr(iris, "class|names")
#' }
#'
#' @export
regex_attr <- function(x, name) {
  names(attributes(x))[grepl(name, names(attributes(x)))]
}

#' Convert column names to lower or upper case
#'
#' Convenience functions to modify the case of all column names in a data frame.
#' These functions update the names **in place** (using `data.table`).
#'
#' @param df A data.frame.
#'
#' @return The input data frame with column names modified in place.
#'   Called for side effects.
#'
#' @examples
#' \donttest{
#' df <- mtcars
#'
#' # Convert to upper case
#' set_col_upper(df)
#'
#' # Convert to lower case
#' set_col_lower(df)
#'
#' }
#'
#' @export
set_col_lower <- function(df)
  data.table::setnames(df, colnames(df), tolower(colnames(df)))

#' @rdname set_col_lower
#' @export
set_col_upper <- function(df)
  data.table::setnames(df, colnames(df), toupper(colnames(df)))

#' Reorder columns of a data.frame or data.table by reference
#'
#' A convenience wrapper around [data.table::setcolorder()] that allows
#' column reordering with tidy-eval expressions. The function modifies
#' the input object **in place**.
#'
#' @param df A data.frame or data.table.
#' @param neworder Columns to move, specified as bare names inside `.(...)`.
#'   For example, `.(gear, carb)`.
#' @param before,after Optionally, a column (name or position) before/after which
#'   `neworder` should be inserted. Only one of `before` or `after` may be used.
#'
#' @return No return value, called for side effects (the column order of `df` is changed).
#'
#' @examples
#' \donttest{
#' # With data.frame
#' df1 <- mtcars
#' set_col_order(df1, .(gear, carb), after = mpg)
#'
#' # With data.table
#' df2 <- data.table::as.data.table(mtcars)
#' set_col_order(df2, .(gear, carb), before = am)
#' }
#'
#' @export
set_col_order <- function(df, neworder, before = NULL, after = NULL) {
  neworder <- capture_names(df, !!rlang::enquo(neworder))
  before   <- capture_names(df, !!rlang::enquo(before))
  after    <- capture_names(df, !!rlang::enquo(after))
  if (!rlang::has_length(before)) before <- NULL
  if (!rlang::has_length(after)) after <- NULL
  data.table::setcolorder(x = df, neworder = neworder, before = before, after = after)
}

#' Set or get column labels for a data frame
#'
#' Functions to attach or retrieve descriptive labels on columns
#' of a data.frame. Labels are stored as the `"label"` attribute
#' of each column.
#'
#' @param df A data.frame.
#' @param labels A character vector of labels to assign. Must be the same
#'   length as `cols`.
#' @param cols A character vector of column names to set or get labels for.
#'   Defaults to all columns in `df`.
#'
#' @return
#' * `set_labels()` returns the modified data frame, invisibly (labels
#'   are added as side effects).
#' * `get_labels()` returns a character vector of labels corresponding
#'   to the requested columns.
#'
#' @examples
#' \donttest{
#' df <- data.frame(Q1 = c(0, 1, 1), Q2 = c(1, 0, 1))
#'
#' # set labels
#' set_labels(df, labels = c("Rainy?", "Umbrella?"))
#'
#' # get labels
#' get_labels(df) # or View(df)
#' }
#'
#' @export
set_labels <- function(df, labels, cols) {
  if (missing(cols))
    cols <- names(df)
  if (length(cols) != length(labels))
    stop("The number of columns and the number of labels are different.")
  lapply(seq_along(cols),
         function(x) data.table::setattr(df[[cols[[x]]]], "label", labels[[x]]))
  invisible(df)
}

#' @rdname set_labels
#' @export
get_labels <- function(df, cols) {
  if (missing(cols))
    cols <- names(df)
  sapply(cols, function(x) attr(df[[x]], "label"), USE.NAMES = FALSE)
}

#' Convert a data.frame to data.table (experimental)
#'
#' @description
#' `r lifecycle::badge("experimental")`
#'
#' A wrapper around [data.table::setDT()] that converts a `data.frame` to a
#' `data.table` **by reference**. This version adds extra safety checks and
#' preserves naming in the calling environment.
#'
#' @param df A data.frame.
#'
#' @return The input `df`, converted to a data.table, returned invisibly.
#'   The object is modified in place (by reference).
#'
#' @seealso [data.table::setDT()]
#'
#' @examples
#' \donttest{
#' df <- data.frame(x = 1:3, y = 4:6)
#' set_dt(df)
#' class(df)  # now includes "data.table"
#' }
#'
#' @export
set_dt <- function(df) {
  lifecycle::signal_stage("experimental", "set_dt()")
  assert_class(df, "data.frame")
  if (!has_ptr(df)) {
    n <- sys.nframe()
    df_name <- trace_arg_expr(df)
    org_class <- class(df)
    data.table::setDT(df)
    assign(df_name, df, envir = parent.frame(n))
    invisible(df)
  }
  if (!inherits(df, "data.table")) {
    data.table::setattr(df, "class", c("data.table", "data.frame"))
  }
}

#' Convert a data.frame to tibble (experimental)
#'
#' @description
#' `r lifecycle::badge("experimental")`
#'
#' A wrapper that converts a `data.frame` to have tibble classes
#' (`tbl_df`, `tbl`, `data.frame`). Internally it ensures reference
#' semantics similar to [data.table::setDT()], then adjusts the
#' class attribute.
#'
#' @param df A data.frame.
#'
#' @return The input `df`, with tibble classes added, returned invisibly.
#'   The object is modified in place (by reference).
#'
#' @seealso [tibble::as_tibble()]
#'
#' @examples
#' \donttest{
#' df <- data.frame(x = 1:3, y = 4:6)
#' set_tibble(df)
#' class(df)  # "tbl_df" "tbl" "data.frame"
#' }
#'
#' @export
set_tibble <- function(df) {
  lifecycle::signal_stage("experimental", "set_tibble()")
  assert_class(df, "data.frame")
  if (!has_ptr(df)) {
    n <- sys.nframe()
    df_name <- trace_arg_expr(df)
    org_class <- class(df)
    data.table::setDT(df)
    data.table::setattr(df, "class", c("tbl_df", "tbl", "data.frame"))
    assign(df_name, df, envir = parent.frame(n))
    invisible()
  }
  if (!inherits(df, "tbl_df")) {
    data.table::setattr(df, "class", c("tbl_df", "tbl", "data.frame"))
  }
}

#' Convert column names to lowercase (in-place if possible)
#'
#' Converts all column names of a `data.frame`, `tibble`, or `data.table`
#' to lowercase.
#' The function uses [data.table::setnames()] internally, which directly modifies
#' the column names of the input object.
#'
#' If the input is a `data.table`, modification is done **by reference**
#' (no copy).
#' For `data.frame` and similar objects, the operation usually modifies
#' the object in place as well (since `setnames()` works on name attributes),
#' but may occasionally trigger a shallow copy depending on how the object
#' is referenced internally.
#'
#' @param x A `data.frame`, `tibble`, or `data.table` whose column names will be
#'   converted to lowercase.
#'
#' @return Invisibly returns the modified object, typically the same object
#' (in place for `data.table` and usually for `data.frame` as well).
#'
#' @examples
#' \dontrun{
#' library(data.table)
#' dt <- data.table(A = 1, B = 2)
#' set_lower_names(dt)
#' names(dt) # "a", "b"
#'
#' df <- data.frame(A = 1, B = 2)
#' set_lower_names(df)
#' names(df) # "a", "b"
#' }
#'
#' @export
set_lower_names <- function(x)
  data.table::setnames(x, colnames(x), tolower(colnames(x)))

#' @rdname set_lower_names
#' @export
set_upper_names <- function(x)
  data.table::setnames(x, colnames(x), toupper(colnames(x)))

#' Format numbers with commas
#'
#' Convert a numeric or integer vector into a character vector
#' formatted with commas as thousands separators.
#'
#' @param x A numeric or integer vector.
#' @param digits A desired number of digits after the decimal point
#'
#' @return A character vector of formatted numbers.
#'
#' @examples
#' # format numbers with commas
#' \donttest{as_comma(c(123456, 234567))}
#'
#' @export
as_comma <- function(x, digits = 0) {
  assert_class(x, c("integer", "numeric"))
  formatC(x, format = "f", digits = digits, big.mark = ",")
}

#' Paste vector elements with commas
#'
#' Combine elements of a vector into a string with commas.
#' Optionally insert newlines between elements for readability.
#'
#' @param x A vector.
#' @param newline Logical; if `TRUE`, each element is placed on a new line.
#'
#' @return No return value, prints to the console.
#'
#' @examples
#' # paste elements with commas
#' \donttest{paste_comma(names(mtcars))}
#'
#' @export
paste_comma <- function(x, newline = FALSE) {
  if (newline) {
    cat(paste0("c(", paste0("\"", paste(x, collapse = "\"\n, \""), "\""), "\n)"))
  }
  else {
    cat(paste0("c(", paste0("\"", paste(x, collapse = "\", \""), "\""), ")"))
  }
}

#' Quote and paste symbols with commas
#'
#' Capture unquoted symbols, wrap them in quotes, and print them
#' as a comma-separated character vector.
#'
#' @param ... Unquoted variable names or expressions.
#' @param newline Logical; if `TRUE`, print each element on a new line.
#'
#' @return No return value, prints to the console.
#'
#' @examples
#' \donttest{quote_comma(mpg, cyl, disp, hp, drat)} # c("mpg", "cyl", "disp", "hp", "drat")
#'
#' @export
quote_comma <- function(..., newline = FALSE) {
  if (newline) {
    cat(paste0("c(", paste0("\"", paste(vapply(substitute(list(...)),
                                  deparse, "character")[-1L], collapse = "\"\n, \""),
               "\""), "\n)"))
  }
  else {
    cat(paste0("c(", paste0("\"", paste(vapply(substitute(list(...)),
                                  deparse, "character")[-1L], collapse = "\", \""),
               "\""), ")"))
  }
  cat("\n")
}

#' Convert columns to numeric (in-place for data.table)
#'
#' Efficiently convert selected columns of a data frame, tibble, or
#' data.table to numeric type.
#' If the input is a data.table, the conversion is performed **in place**
#' (no copy).
#' If the input is a regular data.frame or tibble, a shallow copy is made
#' and the result is automatically restored to the original class.
#'
#' If `cols` is not provided (or is empty), all columns of type `integer`
#' are automatically selected for conversion.
#'
#' @param df A data.frame, tibble, or data.table containing the columns to convert.
#' @param cols Character vector of column names to convert to numeric. If missing
#'   or empty, all `integer`-type columns are selected automatically.
#' @param suppress_warnings Logical; if `TRUE`, suppresses warnings such as
#'   `"NAs introduced by coercion"`. Default is `FALSE`.
#'
#' @return
#' - If `df` is a data.table: the same object, modified in place.
#' - If `df` is a data.frame or tibble: a new object of the original class,
#'   with specified (or auto-selected) columns converted to numeric.
#'
#' @details
#' This function uses [data.table::set()] for direct column replacement,
#' minimizing memory use and avoiding full table copies.
#' It is internally wrapped by `ensure_dt_env()` to preserve the input's
#' original class and ensure safe restoration for non-data.table inputs.
#'
#' @examples
#' \dontrun{
#' library(data.table)
#' DT <- data.table(a = 1:3, b = c("4", "5", "x"))
#'
#' # Auto-select integer columns (here: "a") and convert in place
#' numify(DT)
#' str(DT)
#'
#' # Explicit columns
#' df <- data.frame(a = c("1", "2"), b = c("3", "x"))
#' df2 <- numify(df, "b")
#' str(df2)
#' }
#'
#' @export
numify <- function(df, cols, suppress_warnings = FALSE) {
  assert_class(df, "data.frame")

  env <- ensure_dt_env(df)
  dt  <- env$dt

  # Auto-select integer columns when cols is missing/empty
  if (missing(cols) || is.null(cols) || length(cols) == 0L) {
    cols <- names(dt)[vapply(dt, is.integer, logical(1L))]
  }

  # Nothing to do
  if (length(cols) == 0L) {
    return(env$restore(dt))
  }

  for (j in cols) {
    x <- dt[[j]]
    x <- if (suppress_warnings) suppressWarnings(as.numeric(x)) else as.numeric(x)
    data.table::set(dt, j = j, value = x)
  }

  env$restore(dt)
}

#' Convert all integer64 columns to numeric
#'
#' Detects columns of type `integer64` (from the **bit64** package)
#' and converts them to `numeric`.
#' If the input is a `data.table`, the conversion is performed **in place**
#' (by reference, no copy).
#' For data.frame or tibble inputs, the function returns a modified copy
#' restored to the original class.
#'
#' @param df A data.frame, tibble, or data.table.
#'   Columns with class `"integer64"` will be converted to numeric.
#'
#' @return
#' - For data.table input: the same object (modified in place, invisibly).
#' - For data.frame/tibble input: a new object of the same class, returned
#'   with integer64 columns converted to numeric.
#'
#' @examples
#' \dontrun{
#' library(data.table)
#' library(bit64)
#'
#' # data.table example (in-place)
#' dt <- data.table(a = as.integer64(1:3), b = letters[1:3])
#' str(dt)
#' set_i64_to_num(dt)
#' str(dt) # 'a' converted to numeric
#'
#' # data.frame example (copy returned)
#' df <- data.frame(a = as.integer64(4:6), b = letters[4:6])
#' df2 <- set_i64_to_num(df)
#' str(df2)
#' }
#'
#' @export
numify_i64 <- function(df) {
  assert_class(df, "data.frame")

  # Identify integer64 columns
  cols <- names(df)[vapply(df, inherits, logical(1L), what = "integer64")]

  if (length(cols) == 0L) {
    message("No integer64 columns found.")
    return(invisible(df))
  }

  # Use numify() to handle data.table / data.frame / tibble
  numify(df, cols = cols, suppress_warnings = TRUE)
}


# Text --------------------------------------------------------------------

#' Truncate Text with Ellipsis
#'
#' Shortens strings longer than a specified width and appends an ellipsis (`"..."`).
#' Uses `stringr::str_trunc()` for consistent behavior with multilingual text.
#'
#' @param x Character vector.
#' @param width Integer, maximum display width before truncation.
#' @param ellipsis String appended when truncating (default: `"..."`).
#'
#' @return Character vector with truncated strings.
#'
#' @examples
#' text <- c(
#'   "The quick brown fox jumps over the lazy dog.",
#'   "Sphinx of black quartz, judge my vow."
#' )
#'
#' truncate_text(text, width = 15)
#' #> [1] "The quick bro..." "Sphinx of bla..."
#'
#' truncate_text(text, width = 25)
#' #> [1] "The quick brown fox jumps..." "Sphinx of black quartz, j..."
#'
#' @export
truncate_text <- function(x, width = 10L, ellipsis = "...") {
  ifelse(nchar(x) > width,
         paste0(substr(x, 1L, width - nchar(ellipsis)), ellipsis),
         x)
}


# Math --------------------------------------------------------------------

#' Integer Division and Modulo
#'
#' Compute the integer quotient and remainder and return both as a list.
#'
#' @param x Numeric vector of values to be divided.
#' @param div Numeric vector or scalar divisor. If scalar, it will be
#'   recycled to match the length of `x`.
#'
#' @return A list with:
#' \describe{
#'   \item{quotient}{Integer division results.}
#'   \item{remainder}{Modulo results.}
#' }
#'
#' @examples
#' \donttest{
#' divmod(17, 5)
#' divmod(1:10, 3)
#' divmod(10:12, 1:3)
#' }
#'
#' @export
divmod <- function(x, div) {
  if (!is.numeric(x) || !is.numeric(div))
    stop("Both 'x' and 'div' must be numeric.", call. = FALSE)

  # handle recycling safely
  if (length(div) == 1L) div <- rep(div, length(x))
  if (length(x) != length(div))
    stop("Lengths of 'x' and 'div' must match or div must be length 1.", call. = FALSE)

  list(
    quotient  = x %/% div,
    remainder = x %%  div
  )
}

#' Tests whether each element of a numeric vector is non-integer
#' (i.e., has a fractional part) within a given tolerance.
#'
#' @param x A numeric vector to check.
#' @param tol Numeric tolerance for floating-point comparison.
#'   Defaults to `.Machine$double.eps^0.5`.
#'
#' @return A logical vector of the same length as `x`, where `TRUE`
#'   indicates that the value has a decimal (fractional) component.
#'
#' @details
#' The test is based on comparing each value to its nearest integer:
#' `abs(x - round(x)) > tol`.
#' This accounts for floating-point rounding errors.
#'
#' @examples
#' \donttest{
#' has_decimal(c(1, 2.5, 3.0, 4.75))
#' #> [1] FALSE  TRUE FALSE  TRUE
#' }
#'
#' @export
has_decimal <- function(x, tol = .Machine$double.eps^0.5) {
  if (!is.numeric(x))
    stop("`x` must be numeric.", call. = FALSE)
  abs(x - round(x)) > tol
}

# To be updated -----------------------------------------------------------

join <- function(..., by, all = FALSE, all.x = all, all.y = all, sort = TRUE) {
  Reduce(function(...) merge(..., by = by, all = all, all.x = all.x,
                             all.y = all.y, sort = sort), list(...))
}

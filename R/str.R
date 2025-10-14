#' Paste strings
#'
#' Collapse a character vector into a single string using a separator.
#'
#' @param x A character vector.
#' @param collapse A single string used to separate elements (not `NA_character_`).
#'   Default: `"|"`.
#'
#' @return A length-1 character vector (a single collapsed string).
#'
#' @examples
#' \donttest{paste_str(c("a", "b", "c"))}
#'
#' @export
paste_str <- function(x, collapse = "|")
  paste(x, collapse = collapse)

#' Paste unique strings
#'
#' Collapse unique, non-missing elements of a character vector into a single string.
#'
#' @param x A character vector.
#' @param collapse A single string used to separate elements (not `NA_character_`).
#'   Default: `"|"`.
#'
#' @return A length-1 character vector.
#'
#' @examples
#' \donttest{paste_uni_str(c("a", "a", "c", "b", "b", "d", "e"))}
#'
#' @export
paste_uni_str <- function(x, collapse = "|")
  paste(unique(x[!is.na(x)]), collapse = collapse)

#' Paste sorted unique strings
#'
#' Collapse **sorted** unique, non-missing elements into a single string.
#'
#' @param x A character vector.
#' @param collapse A single string used to separate elements (not `NA_character_`).
#'   Default: `"|"`.
#'
#' @return A length-1 character vector.
#'
#' @examples
#' # paste sorted unique string
#' \donttest{paste_sort_uni_str(c("a", "a", "c", "b", "b", "d", "e"))}
#'
#' @export
paste_sort_uni_str <- function(x, collapse = "|")
  paste(sort(unique(x[!is.na(x)])), collapse = collapse)

#' Split strings
#'
#' Split each input string on a delimiter.
#'
#' @param x A character vector.
#' @param split A single string delimiter.
#'
#' @return A list of character vectors (one element per input string).
#'
#' @examples
#' # split strings
#' \donttest{split_str(c("a|b|a", "c|d|c"), split = "|")}
#'
#' @export
split_str <- function(x, split = "|")
  strsplit(x, split = split, fixed = TRUE)

#' Split and paste unique strings
#'
#' Split each string then paste the **unique** tokens back with the same separator.
#'
#' @param x A character vector.
#' @param split A single string delimiter.
#'
#' @return A character vector of the same length as `x`.
#'
#' @examples
#' \donttest{
#' split_and_paste_uni_str(c("b|a|b", "d|c|d"), split = "|")
#' split_and_paste_sort_uni_str(c("b|a|b", "d|c|d"), split = "|")
#' }
#'
#' @export
split_and_paste_uni_str <- function(x, split = "|") {
  z <- split_str(x, split = split)
  vapply(z, function(x) paste_uni_str(x, collapse = split), character(1L))
}

#' @rdname split_and_paste_uni_str
#' @description
#' Split each string then paste the **sorted unique** tokens with the same separator.
#'
#' @export
split_and_paste_sort_uni_str <- function(x, split = "|") {
  z <- split_str(x, split = split)
  vapply(z, function(x) paste_sort_uni_str(x, collapse = split), character(1L))
}

#' Count pattern matches
#'
#' Count regex matches in each string.
#'
#' @param pattern A string containing a regular expression.
#' @param x A character vector.
#' @param ignore.case Logical; if `TRUE`, case is ignored. Default: `FALSE`.
#'
#' @return An integer vector of match counts (same length as `x`).
#'
#' @examples
#' \donttest{
#' count_pattern(pattern = "c", c("a|b|c", "a|c|c"))
#' }
#'
#' @export
count_pattern <- function(pattern, x, ignore.case = FALSE) {
  vapply(
    gregexpr(pattern, x, ignore.case = ignore.case, perl = TRUE),
    function(m) length(m[m != -1]),
    integer(1L)
  )
}

#' Get first match
#'
#' Extract the first regex match from each string (empty string if no match).
#'
#' @param pattern A string containing a regular expression.
#' @param x A character vector.
#' @param ignore.case Logical; if `TRUE`, case is ignored. Default: `TRUE`.
#'
#' @return A character vector of the same length as `x`.
#'
#' @examples
#' \donttest{
#' get_pattern(pattern = "c", c("a|b|c", "a|c|c"))
#' }
#'
#' @export
get_pattern <- function(pattern, x, ignore.case = TRUE) {
  r <- regexpr(pattern, x, ignore.case = ignore.case, perl = TRUE)
  z <- rep("", length(x))
  z[r != -1] <- regmatches(x, r)
  z
}

#' Get all matches
#'
#' Extract all regex matches from each string and collapse them with a separator.
#'
#' @param pattern A string containing a regular expression.
#' @param x A character vector.
#' @param collapse A single string used to separate matches (not `NA_character_`).
#'   Default: `"|"`.
#' @param ignore.case Logical; if `TRUE`, case is ignored. Default: `TRUE`.
#'
#' @return A character vector of the same length as `x`
#'   (empty string for elements with no matches).
#'
#' @examples
#' \donttest{get_pattern_all(pattern = "c", c("a|b|c", "a|c|c"))}
#'
#' @export
get_pattern_all <- function(pattern, x, collapse = "|", ignore.case = TRUE) {
  r <- gregexpr(pattern, x, ignore.case = ignore.case, perl = TRUE)
  z <- regmatches(x, r)
  vapply(z, function(s) paste(s, collapse = collapse), character(1L))
}

#' Delete patterns
#'
#' Remove all regex matches from each string.
#'
#' @param pattern A string containing a regular expression.
#' @param x A character vector.
#'
#' @return A character vector with matches removed.
#'
#' @examples
#' \donttest{del_pattern(pattern = "c", c("abc", "acc"))}
#'
#' @export
del_pattern <- function(pattern, x)
  gsub(pattern, "", x)

#' Escape punctuation characters in a string
#'
#' Escapes all punctuation characters in a character vector by prefixing
#' them with a backslash (`\\`). This is useful when preparing text for use
#' in regular expressions or other contexts where punctuation has special meaning.
#'
#' @param x A character vector whose punctuation characters should be escaped.
#'
#' @return A character vector of the same length as `x`, with each punctuation
#'   character replaced by an escaped version (e.g., `"."` becomes `"\\."`).
#'
#' @examples
#' \donttest{
#' escape_punct("Hello, world!")   # "Hello\\, world\\!"
#' escape_punct("a+b=c?")          # "a\\+b\\=c\\?"
#' escape_punct(c("A.B", "C/D"))   # "A\\.B" "C\\/D"
#' }
#'
#' @export
escape_punct <- function(x) {
  # Replace every punctuation character with its escaped version
  gsub("([[:punct:]])", "\\\\\\1", x, perl = TRUE)
}

#' Paste multiple columns row-wise
#'
#' Concatenate multiple columns of a data frame (or data.table) **row by row**
#' into a single character vector.
#'
#' @param df A data.frame or data.table.
#' @param cols A character vector of column names to concatenate.
#' @param sep A string used to separate the values. Default is `"|"`.
#' @param na.rm Logical; if `TRUE`, missing values are treated as empty strings
#'   before concatenation. Leading, trailing, or duplicated separators are
#'   automatically trimmed. Default is `FALSE`.
#'
#' @return A character vector of length `nrow(df)`, each element representing
#'   the row-wise concatenation of the specified columns.
#'
#' @examples
#' \donttest{
#' df <- data.frame(a = c("x", "y"), b = c(1, 2), c = c(NA, 3))
#'
#' paste_cols(df, c("a", "b"))
#' #> [1] "x|1" "y|2"
#'
#' paste_cols(df, c("a", "c"), sep = " / ", na.rm = TRUE)
#' #> [1] "x" "y / 3"
#' }
#'
#' @export
paste_cols <- function(df, cols, sep = "|", na.rm = FALSE) {
  assert_class(df, "data.frame")

  cols <- capture_names(df, !!rlang::enquo(cols))

  x <- lapply(cols, function(s) df[[s]])

  if (na.rm)
    x[] <- lapply(x, function(s) ifelse(is.na(s), "", s))

  pattern <- sprintf("^\\%s|\\%s\\%s|\\%s$", sep, sep, sep, sep)

  n <- length(x)
  if (n == 1L)
    return(x[[1L]])

  out <- do.call(function(...) paste(..., sep = sep), x)
  out <- gsub(pattern, "", out)

  if (na.rm)
    out <- ifelse(out == "", NA_character_, out)

  out
}

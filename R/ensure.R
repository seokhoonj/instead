#' Ensure a data.table execution environment
#'
#' @description
#' `r lifecycle::badge("experimental")`
#'
#' Provides a lightweight execution-environment **list** for safely using
#' data.table operations, regardless of whether the input is a base
#' data.frame, a tibble, or already a data.table.
#'
#' - If the input is a data.table, it is returned as-is and can be modified
#'   in place.
#' - If the input is a data.frame or tibble, a **copy** is made, converted to
#'   a data.table, and returned. A `restore()` function is provided to convert
#'   a result back to the original class (and to re-attach any extra S3 classes).
#'
#' This helper is intended for package functions that rely on in-place
#' data.table semantics while respecting the user's original data type.
#'
#' @param x A data.frame-like object: data.table, data.frame, or tibble.
#'
#' @return A list with three components:
#' \item{dt}{A data.table (safe to modify).}
#' \item{restore}{A function that restores a data.table result (typically `dt`)
#'   to the original input class and re-attaches any extra S3 classes.}
#' \item{inplace}{Logical: `TRUE` if input was already a data.table (modified in place),
#'   `FALSE` otherwise.}
#'
#' @examples
#' \donttest{
#' df <- data.frame(x = 1:3, y = 4:6)
#' env <- ensure_dt_env(df)
#' dt <- env$dt
#' dt[, z := x + y]      # in-place modification
#' env$restore(dt)       # back to data.frame
#' }
#' \dontrun{
#' tb <- tibble(a = 1:2, b = 3:4)
#' env2 <- ensure_dt_env(tb)
#' dt2 <- env2$dt
#' dt2[, c := a * b]
#' env2$restore(dt2)     # back to tibble
#' }
#'
#' @export
ensure_dt_env <- function(x) {
  lifecycle::signal_stage("experimental", "ensure_dt_env()")

  if (!inherits(x, "data.frame"))
    stop("`x` must be a data.frame / tibble / data.table.", call. = FALSE)

  base_classes <- c("data.table", "data.frame", "tbl_df", "tbl")

  extra_classes <- setdiff(class(x), base_classes)

  restore_with_extra_class <- function(z) { # closure doesn't need extra_classes as an argument
    if (!length(extra_classes)) return(z)
    instead::prepend_class(z, extra_classes)
  }

  if (inherits(x, "data.table")) {
    # Input is already a data.table -> modify in place
    list(
      dt = x,
      restore = restore_with_extra_class, # return unchanged
      inplace = TRUE
    )
  } else {
    # Protect original: make a shallow copy, then convert
    x_copy <- data.table::copy(x)
    data.table::setDT(x_copy)
    org_class <- class(x)

    restore <- function(z) {
      # Convert back to original type
      if ("tbl_df" %in% org_class) {
        if (!requireNamespace("tibble", quietly = TRUE))
          stop("The `tibble` package is required to restore to tibble.",
               call. = FALSE)
        z <- tibble::as_tibble(as.data.frame(z))
      } else if ("data.frame" %in% org_class) {
        z <- as.data.frame(z)
      } else {
        z
      }
      restore_with_extra_class(z)
    }

    list(
      dt = x_copy,
      restore = restore,
      inplace = FALSE
    )
  }
}

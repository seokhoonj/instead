#' Longer (S3 generic)
#'
#' A lightweight S3 generic for reshaping objects to long form,
#' conceptually similar to `data.table::melt()`. Concrete methods may
#' delegate to `data.table::melt()` or provide class-specific behavior.
#'
#' @param x An object to reshape (e.g., `data.frame`, `data.table`,
#'   or a domain-specific class).
#' @param ... Passed to methods.
#'
#' @return A long-form object as defined by the dispatched method.
#'
#' @examples
#' \dontrun{
#' longer(mtcars, id_vars = c(cyl, gear), measure_vars = c(mpg, wt))
#' }
#'
#' @export
longer <- function(x, ...) {
  UseMethod("longer")
}

#' Reshape a data.frame to long format (NSE-aware wrapper of `data.table::melt()`)
#'
#' `longer.data.frame()` reshapes a `data.frame` from wide to long format
#' using `data.table::melt()` internally.
#'
#' It supports **NSE (non-standard evaluation)** for `id_vars` and
#' `measure_vars`, allowing unquoted column names (e.g. `c(mpg, wt)`).
#' If these arguments are omitted, all remaining columns are inferred.
#'
#' The input is coerced to a `data.table`, reshaped, and the original
#' outer class (e.g. tibble) is restored on return.
#'
#' @param x A `data.frame`.
#' @param id_vars Columns to keep fixed (identifier variables). Supports
#'   unquoted names, character vectors, or numeric indices.
#' @param measure_vars Columns to gather into long form. Supports
#'   unquoted names, character vectors, or numeric indices.
#' @param variable_name Name of the output "key" column (default: `"variable"`).
#' @param value_name Name of the output "value" column (default: `"value"`).
#' @param ... Additional arguments forwarded to `data.table::melt()`.
#' @param na.rm,variable_factor,value_factor,verbose Passed directly to
#'   `data.table::melt()`.
#'
#' @return A long-form table with the original outer class restored.
#'
#' @details
#' Column symbols are resolved internally using `capture_names()`.
#' This allows flexible use of both quoted and unquoted variable names
#' without losing compatibility with `data.table::melt()`.
#'
#' @seealso [data.table::melt()]
#'
#' @examples
#' \dontrun{
#' df <- mtcars
#'
#' # Unquoted NSE columns
#' longer.data.table(dt, id_vars = c(cyl, gear), measure_vars = c(mpg, wt))
#'
#' # Quoted style (standard data.table syntax)
#' longer.data.table(dt, id_vars = c("cyl", "gear"), measure_vars = c("mpg", "wt"))
#' }
#'
#' @rdname longer
#' @method longer data.frame
#' @export
longer.data.frame <- function(x, id_vars, measure_vars,
                              variable_name = "variable", value_name = "value",
                              ..., na.rm = FALSE, variable_factor = TRUE,
                              value_factor = FALSE,
                              verbose = getOption("data.table.verbose")) {
  env <- ensure_dt_env(x)
  dt  <- env$dt

  iv <- capture_names(dt, !!rlang::enquo(id_vars))
  mv <- capture_names(dt, !!rlang::enquo(measure_vars))

  if (!length(iv)) iv <- NULL
  if (!length(mv)) mv <- NULL

  dl <- data.table::melt(
    data = dt,
    id.vars = id_vars,
    measure.vars = measure_vars,
    variable.name = variable_name,
    value.name = value_name,
    ..., na.rm = na.rm, variable.factor = variable_factor,
    value.factor = value_factor, verbose = verbose
  )

  env$restore(dl)
}

#' Reshape a data.table to long format (NSE-aware wrapper of `data.table::melt()`)
#'
#' `longer.data.table()` provides a convenient NSE interface to
#' `data.table::melt()`, supporting unquoted column selection for
#' `id_vars` and `measure_vars`.
#'
#' Both `id_vars` / `measure_vars` and `id.vars` / `measure.vars`
#' argument styles are accepted for compatibility.
#'
#' @param x A `data.table`.
#' @param id_vars Columns to keep fixed (identifier variables). Supports
#'   unquoted names, character vectors, or numeric indices.
#' @param measure_vars Columns to gather into long form. Supports
#'   unquoted names, character vectors, or numeric indices.
#' @param variable_name Name of the output "key" column (default: `"variable"`).
#' @param value_name Name of the output "value" column (default: `"value"`).
#' @param ... Additional arguments forwarded to `data.table::melt()`.
#' @param na.rm,variable_factor,value_factor,verbose Passed directly to
#'   `data.table::melt()`.
#'
#' @return A long-form `data.table`.
#'
#' @details
#' This wrapper resolves NSE expressions via `instead::capture_names()`,
#' enabling both quoted and unquoted column references while maintaining
#' full compatibility with `data.table::melt()`.
#'
#' @seealso [data.table::melt()]
#'
#' @examples
#' \dontrun{
#' dt <- data.table::as.data.table(mtcars)
#'
#' # Unquoted NSE columns
#' longer.data.table(dt, id_vars = c(cyl, gear), measure_vars = c(mpg, wt))
#'
#' # Quoted style (standard data.table syntax)
#' longer.data.table(dt, id_vars = c("cyl", "gear"), measure_vars = c("mpg", "wt"))
#' }
#'
#' @rdname longer
#' @method longer data.table
#' @export
longer.data.table <- function(x, id_vars, measure_vars,
                              variable_name = "variable", value_name = "value",
                              ..., na.rm = FALSE, variable_factor = TRUE,
                              value_factor = FALSE,
                              verbose = getOption("data.table.verbose")) {
  iv <- capture_names(x, !!rlang::enquo(id_vars))
  mv <- capture_names(x, !!rlang::enquo(measure_vars))

  if (!length(iv)) iv <- NULL
  if (!length(mv)) mv <- NULL

  data.table::melt(
    data = x,
    id.vars = iv,
    measure.vars = mv,
    variable.name = variable_name,
    value.name = value_name,
    ..., na.rm = na.rm, variable.factor = variable_factor,
    value.factor = value_factor, verbose = verbose
  )
}

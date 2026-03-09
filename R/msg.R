#' Print a section rule
#'
#' Prints a colored section rule using `cli::rule()`.
#'
#' @param x Character; section title.
#' @param line Integer; rule style passed to `cli::rule()`.
#' @param color Character; one of `"cyan"`, `"blue"`, `"red"`, `"green"`,
#' `"yellow"`, `"magenta"`, `"white"`, or `"silver"`.
#'
#' @return Invisibly returns `x`.
#'
#' @export
msg_rule <- function(x, line = 2, color = c("cyan", "blue", "red", "green",
                                            "yellow", "magenta", "white",
                                            "silver")) {
  color <- match.arg(color)

  txt <- cli::rule(x, line = line)

  txt <- switch(
    color,
    cyan    = cli::col_cyan(txt),
    blue    = cli::col_blue(txt),
    red     = cli::col_red(txt),
    green   = cli::col_green(txt),
    yellow  = cli::col_yellow(txt),
    magenta = cli::col_magenta(txt),
    white   = cli::col_white(txt),
    silver  = cli::col_silver(txt),
    none    = txt
  )

  cat(txt, "\n")
  invisible(x)
}


#' Print a progress step message
#'
#' Prints a step message such as `[01/10] Plan A`.
#'
#' @param i Current index.
#' @param n Total number of steps.
#' @param x Character; message text.
#'
#' @return Invisibly returns `x`.
#'
#' @examples
#' msg_step(1, 10, "Plan A")
#'
#' @export
msg_step <- function(i, n, x) {
  width <- nchar(as.character(n))
  cli::cli_alert_info(
    "[{sprintf(paste0('%0', width, 'd'), i)}/{sprintf(paste0('%0', width, 'd'), n)}] {x}"
  )
  invisible(x)
}

#' Print an info message
#'
#' @param x Character; message text.
#'
#' @return Invisibly returns `x`.
#'
#' @examples
#' msg_info("Loading workbook")
#'
#' @export
msg_info <- function(x) {
  cli::cli_alert_info("{x}")
  invisible(x)
}

#' Print a success message
#'
#' @param x Character; message text.
#' @param value Optional character value to append after `x`.
#'
#' @return Invisibly returns `x`.
#'
#' @examples
#' msg_done("Workbook saved")
#' msg_done("Workbook saved", "report.xlsx")
#'
#' @export
msg_done <- function(x = "Done", value = NULL) {
  if (is.null(value)) {
    cli::cli_alert_success("{x}")
  } else {
    cli::cli_alert_success("{x}: {value}")
  }
  invisible(x)
}

#' Print a warning message
#'
#' @param x Character; message text.
#'
#' @return Invisibly returns `x`.
#'
#' @examples
#' msg_warn("No rows matched")
#'
#' @export
msg_warn <- function(x) {
  cli::cli_alert_warning("{x}")
  invisible(x)
}

#' Print a failure message
#'
#' @param x Character; message text.
#'
#' @return Invisibly returns `x`.
#'
#' @examples
#' msg_fail("Workbook save failed")
#'
#' @export
msg_fail <- function(x) {
  cli::cli_alert_danger("{x}")
  invisible(x)
}

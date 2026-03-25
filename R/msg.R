#' Print a section rule
#'
#' Prints a colored section rule using `cli::rule()` and records the start
#' time so that [msg_rule_elapsed_time()] can later print the elapsed time.
#'
#' @param x Character; section title.
#' @param line Integer; rule style passed to `cli::rule()`.
#' @param color Character; one of `"cyan"`, `"blue"`, `"red"`, `"green"`,
#' `"yellow"`, `"magenta"`, `"white"`, or `"silver"`.
#' @param key Character scalar identifying the timer slot. Default `"rule"`.
#'
#' @return Invisibly returns `x`.
#'
#' @export
msg_rule <- function(x,
                     line = 2,
                     color = c("cyan", "blue", "red", "green",
                               "yellow", "magenta", "white", "silver"),
                     key = "rule") {

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
    txt
  )

  if (is.null(.INSTEAD_ENV$msg$time)) .INSTEAD_ENV$msg$time <- list()
  if (is.null(.INSTEAD_ENV$msg$text)) .INSTEAD_ENV$msg$text <- list()

  .INSTEAD_ENV$msg$time[[key]] <- Sys.time()
  .INSTEAD_ENV$msg$text[[key]] <- x

  cat(txt, "\n", sep = "")
  invisible(x)
}

#' Print elapsed time as a section rule
#'
#' Prints the elapsed time since the most recent call to [msg_rule()]
#' as a styled rule using `cli::rule()`.
#'
#' @param x Character label for the elapsed time. Default `"Elapsed Time"`.
#' @param line Integer; rule style passed to `cli::rule()`.
#' @param color Character; one of `"cyan"`, `"blue"`, `"red"`, `"green"`,
#'   `"yellow"`, `"magenta"`, `"white"`, or `"silver"`.
#' @param key Character scalar identifying the timer slot. Default `"rule"`.
#'
#' @return Invisibly returns the formatted elapsed-time string.
#'
#' @examples
#' \dontrun{
#' msg_rule("Start processing")
#' Sys.sleep(0.5)
#' msg_rule_elapsed_time()
#'
#' msg_rule("RR calculation")
#' Sys.sleep(0.3)
#' msg_rule_elapsed_time("RR Total Time")
#' }
#'
#' @export
msg_rule_elapsed_time <- function(
    x = "Elapsed Time",
    line = 2,
    color = c("cyan", "blue", "red", "green",
              "yellow", "magenta", "white", "silver"),
    key = "rule"
) {
  color <- match.arg(color)

  times <- .INSTEAD_ENV$msg$time
  stime <- if (!is.null(times)) times[[key]] else NULL

  if (is.null(stime)) {
    stop(
      sprintf("No start time recorded for key: %s", key),
      call. = FALSE
    )
  }

  secs <- as.numeric(Sys.time() - stime, units = "secs")
  time_txt <- format_elapsed_time(secs)

  txt <- cli::rule(
    sprintf("%s: %s", x, time_txt),
    line = line
  )

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
    txt
  )

  cat(txt, "\n", sep = "")
  invisible(time_txt)
}

#' Print a step message
#'
#' Prints a progress message of the form `"[i/n] x"` and records the start
#' time internally so that [msg_step_elapsed_time()] can later print the elapsed
#' time for the same step.
#'
#' @param i Current step index.
#' @param n Total number of steps.
#' @param x Message text.
#' @param key Character scalar identifying the timer slot. Default `"step"`.
#' @param newline Logical; if `FALSE` (default), keep the cursor on the same
#'   line so that [msg_step_elapsed_time()] can be appended after the message.
#'   If `TRUE`, print a normal line break.
#'
#' @return Invisibly returns `x`.
#'
#' @examples
#' \dontrun{
#' msg_step(1, 10, "Start", newline = TRUE)
#'
#' msg_step(2, 10, "Processing")
#' Sys.sleep(0.5)
#' msg_elapsed_time()
#' }
#'
#' @export
msg_step <- function(i, n, x, key = "step", newline = FALSE) {
  width <- nchar(as.character(n))
  msg <- sprintf(
    "[%0*d/%0*d] %s:",
    width, i,
    width, n,
    x
  )

  # initialize storage if needed
  if (is.null(.INSTEAD_ENV$msg$time)) .INSTEAD_ENV$msg$time <- list()
  if (is.null(.INSTEAD_ENV$msg$text)) .INSTEAD_ENV$msg$text <- list()

  .INSTEAD_ENV$msg$time[[key]] <- Sys.time()
  .INSTEAD_ENV$msg$text[[key]] <- msg

  if (newline) {
    cli::cli_alert_info("{msg}")
  } else {
    cat(cli::col_blue(cli::symbol$info), " ", msg, " ", sep = "")
  }

  invisible(x)
}

#' Print elapsed time for a step message
#'
#' Prints the elapsed time since the most recent call to [msg_step()] for
#' the given key.
#'
#' @param key Character scalar identifying the timer slot. Default `"step"`.
#' @param newline Logical; if `TRUE` (default), append a line break after the
#'   elapsed time. If `FALSE`, do not append a line break.
#'
#' @return Invisibly returns the formatted elapsed-time string.
#'
#' @examples
#' \dontrun{
#' msg_step(1, 3, "Task", newline = FALSE)
#' Sys.sleep(0.25)
#' msg_step_elapsed_time()
#' }
#'
#' @export
msg_step_elapsed_time <- function(key = "step", newline = TRUE) {
  times <- .INSTEAD_ENV$msg$time
  stime <- if (!is.null(times)) times[[key]] else NULL

  if (is.null(stime)) {
    stop(
      sprintf("No start time recorded for key: %s", key),
      call. = FALSE
    )
  }

  secs <- as.numeric(Sys.time() - stime, units = "secs")

  time_txt <- format_elapsed_time(secs)

  if (newline) {
    cat(time_txt, "\n", sep = "")
  } else {
    cat(time_txt)
  }

  invisible(time_txt)
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

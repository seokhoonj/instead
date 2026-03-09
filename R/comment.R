#' Print a comment divider line
#'
#' Prints a comment divider line commonly used to visually separate
#' sections in R scripts. The divider consists of underscores and
#' hyphens and begins with `#`, making it a valid comment line.
#'
#' @param width Integer; total width of the line including `#`.
#'   Default is `76`.
#'
#' @return Invisibly returns the generated line.
#'
#' @examples
#' comment_rule()
#' comment_rule(60)
#'
#' @export
comment_rule <- function(width = 76) {

  # "# " + space = 3 characters
  n <- width - 3

  left  <- floor(n / 2)
  right <- n - left - 1

  line <- paste0(
    "# ",
    strrep("_", left),
    " ",
    strrep("-", right)
  )

  cat(line, "\n")
  invisible(line)
}

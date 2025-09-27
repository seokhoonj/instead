#' List available fonts on the system
#'
#' Returns a list of fonts available on the current system.
#'
#' @param pattern Optional pattern to filter font names.
#'
#' @return Character vector of available font names.
#'
#' @examples
#' \donttest{
#' # List all fonts
#' list_fonts()
#'
#' # Fonts containing "Arial"
#' list_fonts("Arial")
#' }
#'
#' @export
list_fonts <- function(pattern = NULL) {
  os_type <- Sys.info()[["sysname"]]

  if (os_type == "Linux") {
    tryCatch({
      # Get font families from fc-list
      result <- system(
        "fc-list : family | cut -d, -f1 | sort -u",
        intern = TRUE, ignore.stderr = TRUE
      )
      if (!is.null(pattern)) {
        result <- result[grepl(pattern, result, ignore.case = TRUE)]
      }
      return(result)
    }, error = function(e) {
      return(character(0))
    })
  } else if (os_type == "Darwin") {
    tryCatch({
      # Use system_profiler on macOS
      result <- system(
        "system_profiler SPFontsDataType 2>/dev/null | grep 'Full Name:' | sed 's/.*Full Name: //' | sort -u",
        intern = TRUE, ignore.stderr = TRUE
      )
      if (!is.null(pattern)) {
        result <- result[grepl(pattern, result, ignore.case = TRUE)]
      }
      return(result)
    }, error = function(e) {
      return(character(0))
    })
  } else {
    # Return common fonts for other systems
    common_fonts <- c("Arial", "Times New Roman", "Calibri", "Comic Sans MS",
                      "Courier New", "Verdana", "Tahoma", "Georgia", "Impact")
    if (!is.null(pattern)) {
      common_fonts <- common_fonts[grepl(pattern, common_fonts, ignore.case = TRUE)]
    }
    return(common_fonts)
  }
}

#' Check if font is available on the system
#'
#' Checks if a font is available on the current system.
#'
#' @param font_name Character string specifying the font name.
#'
#' @return Logical value indicating if font is available.
#'
#' @examples
#' \donttest{
#' # Check if font exists
#' is_font_available("Arial")
#' is_font_available("NonExistentFont")
#' }
#'
#' @export
is_font_available <- function(font_name) {
  if (is.null(font_name) || !is.character(font_name) || length(font_name) != 1L) {
    return(FALSE)
  }

  # Try different methods based on OS
  os_type <- Sys.info()[["sysname"]]

  if (os_type == "Linux") {
    # Use fc-list command on Linux
    tryCatch({
      result <- system(paste0("fc-list 2>/dev/null | grep -i '", font_name, "' 2>/dev/null"),
                      intern = TRUE, ignore.stderr = TRUE)
      return(length(result) > 0)
    }, error = function(e) {
      return(FALSE)
    })
  } else if (os_type == "Darwin") {
    # Use system_profiler on macOS
    tryCatch({
      result <- system("system_profiler SPFontsDataType 2>/dev/null | grep -i ",
                      shQuote(font_name), intern = TRUE, ignore.stderr = TRUE)
      return(length(result) > 0)
    }, error = function(e) {
      return(FALSE)
    })
  } else if (os_type == "Windows") {
    # Check Windows registry or use PowerShell
    tryCatch({
      # Simple check - assume common fonts are available
      common_fonts <- c("Arial", "Times New Roman", "Calibri", "Comic Sans MS",
                        "Courier New", "Verdana", "Tahoma")
      return(font_name %in% common_fonts)
    }, error = function(e) {
      return(FALSE)
    })
  }

  # Fallback: assume font is available if we can't check
  return(TRUE)
}

#' Set or get the default font for instead outputs
#'
#' These helpers manage the global font family used by `instead`
#' (e.g., Excel writers, plot savers), stored in the option `"instead.font"`.
#'
#' - `set_instead_font()` updates the option, verifying that the font exists
#'   in [systemfonts::system_fonts()]. If the font is not found, the previous
#'   option value is restored and an error is thrown.
#' - `get_instead_font()` retrieves the currently set option.
#'
#' @param family A character string giving the font family to use, or `NULL`.
#'   - Use `NULL` or `""` to reset to system default.
#'   - Must match a font available in [systemfonts::system_fonts()].
#'
#' @return
#' - `set_instead_font()`: the font family (character scalar or `NULL`) that was set.
#' - `get_instead_font()`: the current font family (character scalar or `NULL`).
#'
#' @examples
#' \dontrun{
#' # Set Comic Sans MS if available
#' set_instead_font("Comic Sans MS")
#'
#' # Reset to system default (two ways)
#' set_instead_font(NULL)
#' set_instead_font("")
#'
#' # Get current font
#' get_instead_font()
#' }
#'
#' @export
set_instead_font <- function(family) {
  prev <- getOption("instead.font", default = NULL)

  # Treat both NULL and "" as reset
  if (is.null(family) || identical(family, "")) {
    options(instead.font = NULL)
    return(NULL)
  }

  # system fonts
  sf <- tryCatch(
    systemfonts::system_fonts(),
    error = function(e) NULL
  )
  has_font <- !is.null(sf) && any(tolower(sf$family) == tolower(family))
  if (!has_font) {
    options(instead.font = prev)
    stop(
      sprintf("Font '%s' not found. Reverting to previous option '%s'.",
              family, ifelse(is.null(prev), "NULL", prev)),
      call. = FALSE
    )
  }

  # try registering with sysfonts if available
  if (requireNamespace("sysfonts", quietly = TRUE)) {
    row <- sf[match(tolower(family), tolower(sf$family)), , drop = FALSE]
    path <- row$path[1L]
    if (is.character(path) && nzchar(path) && file.exists(path)) {
      try(sysfonts::font_add(family, path), silent = TRUE)
    }
  }

  options(instead.font = family)

  os_type <- Sys.info()[["sysname"]]
  if (os_type == "Windows") {
    if (requireNamespace("showtext", quietly = TRUE)) {
      showtext::showtext_auto(.mean_dpi())
    } else {
      message("showtext package not found -> skipping font auto-activation.")
    }
  }

  family
}

#' @rdname set_instead_font
#' @export
get_instead_font <- function()
  getOption("instead.font", default = NULL)

# Internal helper function ------------------------------------------------

.mean_dpi <- function() {
  dpi_x <- grDevices::dev.size("px")[1L] / grDevices::dev.size("in")[1L]
  dpi_y <- grDevices::dev.size("px")[2L] / grDevices::dev.size("in")[2L]
  mean(c(dpi_x, dpi_y))
}

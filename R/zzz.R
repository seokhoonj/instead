
.INSTEAD_ENV <- new.env(parent = emptyenv())

.onLoad <- function(libname, pkgname) {
  assign(
    ".DATE_FORMAT",
    paste0(
      "^\\d{4}(0?[13578]|1[02])(0?[1-9]|[12][0-9]|3[01])$|",
      "^\\d{4}(0?[469]|11)(0?[1-9]|[12][0-9]|30)$|",
      "^\\d{4}(02)(0?[1-9]|[12][0-9])$|",
      "^\\d{4}-(0?[13578]|1[02])-(0?[1-9]|[12][0-9]|3[01])$|",
      "^\\d{4}-(0?[469]|11)-(0?[1-9]|[12][0-9]|30)$|",
      "^\\d{4}-(02)-(0?[1-9]|[12][0-9])$|",
      "^\\d{4}/(0?[13578]|1[02])/(0?[1-9]|[12][0-9]|3[01])$|",
      "^\\d{4}/(0?[469]|11)/(0?[1-9]|[12][0-9]|30)$|",
      "^\\d{4}/(02)/(0?[1-9]|[12][0-9])$|",
      "^\\d{4}\\.(0?[13578]|1[02])\\.(0?[1-9]|[12][0-9]|3[01])$|",
      "^\\d{4}\\.(0?[469]|11)\\.(0?[1-9]|[12][0-9]|30)$|",
      "^\\d{4}\\.(02)\\.(0?[1-9]|[12][0-9])$"
    ),
    envir = .INSTEAD_ENV
  )
  lockBinding(".DATE_FORMAT", .INSTEAD_ENV)

  op <- options()
  op.instead <- list(
    instead.eps = 1e-8,
    instead.font = NULL,
    instead.guess_max = 21474836,
    instead.scipen = 14
  )
  toset <- setdiff(names(op.instead), names(op))
  if (length(toset)) options(op.instead[toset])

  invisible()
}

.onUnload <- function(libpath) {
  # if a DLL named 'instead' is registered, unload it safely
  if ("instead" %in% names(getLoadedDLLs())) {
    library.dynam.unload("instead", libpath)
  }
}

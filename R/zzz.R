
.INSTEAD_ENV <- NULL
.onLoad <- function(libname, pkgname) {
  .INSTEAD_ENV <<- new.env()
  assign(".DATE_FORMAT", "^\\d{4}(0?[13578]|1[02])(0?[1-9]|[12][0-9]|3[01])$|^\\d{4}(0?[469]|11)(0?[1-9]|[12][0-9]|30)$|^\\d{4}(02)(0?[1-9]|[12][0-9])$|^\\d{4}-(0?[13578]|1[02])-(0?[1-9]|[12][0-9]|3[01])$|^\\d{4}-(0?[469]|11)-(0?[1-9]|[12][0-9]|30)$|^\\d{4}-(02)-(0?[1-9]|[12][0-9])$|^\\d{4}/(0?[13578]|1[02])/(0?[1-9]|[12][0-9]|3[01])$|^\\d{4}/(0?[469]|11)/(0?[1-9]|[12][0-9]|30)$|^\\d{4}/(02)/(0?[1-9]|[12][0-9])$|^\\d{4}\\.(0?[13578]|1[02])\\.(0?[1-9]|[12][0-9]|3[01])$|^\\d{4}\\.(0?[469]|11)\\.(0?[1-9]|[12][0-9]|30)$|^\\d{4}\\.(02)\\.(0?[1-9]|[12][0-9])$",
         envir = .INSTEAD_ENV)
  op <- options()
  op.instead <- list(
    instead.eps = 1e-8,
    instead.font = NULL,
    instead.guess_max = 21474836,
    instead.scipen = 14
  )
  toset <- !(names(op.instead) %in% names(op))
  if (any(toset)) options(op.instead[toset])
  invisible()
}

.onUnload <- function(libpath) {
  library.dynam.unload("instead", libpath)
}

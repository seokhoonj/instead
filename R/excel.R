
# Reference ---------------------------------------------------------------

#' Convert Excel column letters to numbers
#'
#' Converts Excel column letters (A, B, C, ..., AA, AB, etc.) to column numbers.
#'
#' @param x String containing Excel cell reference or column letters.
#'
#' @return Numeric column number (A=1, B=2, ..., Z=26, AA=27, etc.)
#'
#' @examples
#' \donttest{
#' # Convert column letters to numbers
#' to_r1c1_col("C33")    # Returns 3 (C = 3rd column)
#' to_r1c1_col("ABC")    # Returns 731 (ABC = 731st column)
#' to_r1c1_col("ABC123") # Returns 731 (extracts ABC only)
#' }
#'
#' @export
to_r1c1_col <- function(x) {
  col_letters <- get_pattern("[A-Z]+", x)
  letters <- strsplit(col_letters, "")[[1L]]
  result <- 0
  for (letter in letters) {
    result <- result * 26 + (match(letter, LETTERS))
  }
  result
}

#' Convert column numbers to Excel letters
#'
#' Converts column numbers to Excel column letters (1=A, 2=B, 27=AA, etc.).
#'
#' @param x Numeric column number.
#'
#' @return String with Excel column letters.
#'
#' @examples
#' \donttest{
#' # Convert column numbers to letters
#' to_a1_col(1)   # Returns "A" (1st column)
#' to_a1_col(3)   # Returns "C" (3rd column)
#' to_a1_col(26)  # Returns "Z" (26th column)
#' to_a1_col(27)  # Returns "AA" (27th column)
#' }
#'
#' @export
to_a1_col <- function(x) {
  if (x <= 26)
    return(LETTERS[x])
  result <- NA_character_
  while (x > 0) {
    x <- x - 1
    result <- paste0(LETTERS[(x %% 26) + 1], result)
    x <- x %/% 26
  }
  result
}

# Save workbooks ----------------------------------------------------------

#' Write one or more tables into an existing openxlsx Workbook (stacked or placed)
#'
#' Inserts one or more rectangular tables into a worksheet of an **existing**
#' `openxlsx` workbook, optionally adding a **title row** above each table,
#' stacking tables vertically with a configurable spacer, and auto-sizing
#' data columns. This function **does not save** the workbook to disk; it
#' returns the modified `Workbook` object (use `openxlsx::saveWorkbook()`).
#'
#' If `rc` has **length 1**, the first table is written at that `(row, col)` and
#' subsequent tables are placed **below it** with `row_spacer` blank rows in between.
#' If `rc` has the same length as `data`, each table is written at its own
#' specified top-left cell (no auto-stacking).
#'
#' @param data A data.frame, matrix, or a list of them. Matrices are coerced
#'   to data.frames. A single object is treated as a length-1 list.
#' @param wb An existing `openxlsx` `Workbook` (e.g., from
#'   `openxlsx::createWorkbook()` or `openxlsx::loadWorkbook()`).
#' @param sheet Target worksheet name. Created if it does not exist. Default `"Data"`.
#' @param rc Either a single length-2 numeric vector `c(row, col)` (to stack
#'   tables), or a list of such vectors with the same length as `data`
#'   (to place each table individually).
#' @param row_spacer Integer (>= 0). Number of **blank rows** between stacked
#'   tables when `rc` has length 1. Default `2L`.
#' @param data_titles Optional character vector of titles, written **above**
#'   each table when provided. If `NULL`, `names(data)` are used when available;
#'   otherwise no title is written for that table.
#' @param title_size Numeric font size for `data_titles`. Default `14`.
#' @param row_names Logical; forwarded to `write_data()` to include row names.
#'   Default `FALSE`.
#' @param font_name,font_size,fg_fill,border_color,widths Styling parameters
#'   forwarded to `write_data()`. `widths` is used only when `auto_width = FALSE`.
#' @param auto_width Logical. If `TRUE`, compute per-column widths from the
#'   underlying data (and optional row-name column) and call
#'   `openxlsx::setColWidths()` for **data columns only** (title row is ignored).
#'   Default `TRUE`.
#' @param width_scope One of `"global_max"` or `"per_table"`. When `auto_width = TRUE`:
#'   - `"per_table"` sets widths independently for each table.
#'   - `"global_max"` first scans all tables that *start at the same column*,
#'     takes the **element-wise maximum** width across them, then applies that
#'     once—so later tables don’t shrink widths set by earlier ones.
#'   Default `"global_max"`.
#' @param num_fmt Optional numeric format(s) applied to **data cells** (headers
#'   excluded). Either a single character (applied to all numeric columns),
#'   or a named character vector/list mapping column names to formats
#'   (unknown names are ignored).
#'
#' @details
#' - Titles are written via `write_cell()` (bold, left-aligned by default).
#' - Tables are written via `write_data()` (header/body/footer borders, number
#'   formats, etc.).
#' - Column widths are set relative to the table’s starting column. If
#'   `row_names = TRUE`, the row-name column is included in width calculation.
#' - This function modifies `wb` **in place** and returns it (invisibly).
#'   Saving is up to the caller.
#'
#' @return The modified `Workbook` (invisibly), suitable for chaining and for
#'   passing to `openxlsx::saveWorkbook()`.
#'
#' @examples
#' \dontrun{
#' library(openxlsx)
#'
#' wb <- openxlsx::createWorkbook()
#'
#' # Write two tables stacked starting at A1, with titles and unified widths
#' save_data_wb(
#'   data        = list(iris_head = head(iris), mtcars_head = head(mtcars)),
#'   wb          = wb,
#'   sheet       = "Summary",
#'   rc          = list(c(1, 1)),          # stack from A1
#'   row_spacer  = 2L,
#'   data_titles = c("iris (head)", "mtcars (head)"),
#'   auto_width  = TRUE,
#'   width_scope = "global_max"
#' )
#' openxlsx::saveWorkbook(wb, "summary.xlsx", overwrite = TRUE)
#'
#' # Place at specific cells (no stacking)
#' wb2 <- openxlsx::createWorkbook()
#' save_data_wb(
#'   data        = list(A = head(iris), B = head(mtcars)),
#'   wb          = wb2,
#'   sheet       = "Placed",
#'   rc          = list(c(1, 1), c(1, 10)),  # A1 and J1
#'   auto_width  = TRUE,
#'   width_scope = "per_table"
#' )
#' openxlsx::saveWorkbook(wb2, "placed.xlsx", overwrite = TRUE)
#' }
#'
#' @seealso [save_data_xlsx()], [write_data()], [write_cell()]
#'
#' @export
save_data_wb <- function(data,
                         wb,
                         sheet = "Data",
                         rc = list(c(1L, 1L)),
                         row_spacer = 2L,
                         data_titles = NULL,
                         title_size = 14,
                         row_names = FALSE,
                         font_name = getOption("instead.font"),
                         font_size = 11,
                         fg_fill = "#E6E6E7",
                         border_color = "#000000",
                         widths = 8.43,
                         auto_width = TRUE,
                         width_scope = c("global_max", "per_table"),
                         num_fmt = NULL) {
  if (!inherits(wb, "Workbook"))
    stop("`wb` must be an openxlsx Workbook.", call. = FALSE)

  width_scope <- match.arg(width_scope)

  # normalize `data` to a list of data.frames
  if (is.data.frame(data) || is.matrix(data))
    data <- list(data)
  if (!is.list(data) || !length(data))
    stop("`data` must be a non-empty data.frame, matrix, or list of them.",
         call. = FALSE)

  data <- lapply(data, function(d) {
    if (is.data.frame(d)) return(d)
    if (is.matrix(d))     return(as.data.frame(d))
    stop("Each element of `data` must be a data.frame or matrix.",
         call. = FALSE)
  })

  # titles
  titles <- vector("character", length(data))
  if (!is.null(data_titles) && length(data_titles) == length(data)) {
    titles <- as.character(data_titles)
  } else if (!is.null(names(data))) {
    nms <- names(data)
    titles <- ifelse(!is.na(nms) & nzchar(nms), nms, NA_character_)
  } else {
    titles[] <- NA_character_
  }

  # ensure sheet exists
  existing_sheets <- openxlsx::sheets(wb)
  if (!(sheet %in% existing_sheets)) {
    openxlsx::addWorksheet(wb, sheetName = sheet, gridLines = FALSE)
  }

  # normalize rc -> list of c(row, col)
  if (is.numeric(rc) && length(rc) == 2L) {
    rc <- list(rc)
  } else if (is.list(rc)) {
    rc <- lapply(rc, function(x) {
      if (is.numeric(x) && length(x) == 2L) return(x)
      stop("Each `rc` must be numeric length-2: c(row, col).", call. = FALSE)
    })
  } else {
    stop("`rc` must be c(row, col) or a list of such vectors.", call. = FALSE)
  }

  if (length(row_spacer) != 1L || is.na(row_spacer) || row_spacer < 0L)
    stop("`row_spacer` must be a single non-negative integer.", call. = FALSE)

  # compute starting positions
  starts <- vector("list", length(data))
  if (length(rc) == 1L) {
    # Case 1: user provided a single rc (e.g., c(1, 1))
    # -> Tables will be placed one after another (stacked vertically).
    # We compute each table’s starting row by accumulating:
    #   - +1 for the header row
    #   - +nrow(data[[i]]) for the body
    #   - +row_spacer for spacing between tables
    #   - optionally +1 if a title is present above the table
    cur <- rc[[1L]]
    for (i in seq_along(data)) {
      starts[[i]] <- cur
      title_h <- if (!is.na(titles[[i]])) 1L else 0L
      h <- title_h + 1L + nrow(data[[i]]) + row_spacer
      cur <- c(cur[1L] + h, cur[2L])
    }
  } else if (length(rc) == length(data)) {
    # Case 2: user provided one rc for each table
    # -> Place each table explicitly at the given (row, col) coordinates.
    starts <- rc
  } else {
    # Case 3: invalid input
    # -> rc must either be:
    #    - a single c(row, col) to auto-stack tables, OR
    #    - a list of c(row, col) (one for each table)
    stop("`rc` must have length 1 (stacking) or length(data) (individual placement).",
         call. = FALSE)
  }

  # styles
  title_style <- openxlsx::createStyle(
    textDecoration = "bold",
    fontSize = title_size,
    fontName = font_name
  )

  # pre-compute global widths if requested
  global_widths <- list()
  if (auto_width && identical(width_scope, "global_max")) {
    # Case: "global_max"
    # -> For each table, calculate the required column widths,
    #    and then take the maximum across all tables that start at the same column.
    #    This ensures a consistent width across stacked tables aligned to the same starting col.
    for (i in seq_along(data)) {
      start_col <- starts[[i]][2L]
      cw <- .get_col_widths(data[[i]], row_names = row_names)
      key <- as.character(start_col)
      if (!length(global_widths[[key]])) {
        # first table at this start_col -> store widths directly
        global_widths[[key]] <- cw
      } else {
        # multiple tables share the same start_col
        # -> extend vectors to same length, fill missing with default width (8.43)
        m <- max(length(global_widths[[key]]), length(cw))
        old <- global_widths[[key]]
        length(old) <- m; old[is.na(old)] <- 8.43
        length(cw) <- m; cw[is.na(cw)] <- 8.43
        # take element-wise max so widest cell in any table determines final width
        global_widths[[key]] <- pmax(old, cw)
      }
    }
  }

  # write each table
  for (i in seq_along(data)) {
    start_rc  <- starts[[i]]
    start_row <- start_rc[1L]
    start_col <- start_rc[2L]

    # optional title
    if (!is.na(titles[[i]])) {
      write_cell(
        wb, sheet,
        rc        = c(start_row, start_col),
        value     = titles[[i]],
        bold      = TRUE,
        italic    = FALSE,
        underline = FALSE,
        strikeout = FALSE,
        font_name = font_name,
        font_size = title_size,
        font_color = NULL,
        fg_fill    = NULL,
        h_align    = "left",
        v_align    = "center"
      )
      start_rc[1L]  <- start_rc[1L] + 1L
      start_row     <- start_row    + 1L
    }

    # your existing writer (with borders, header/body/footer styles, num_fmt, etc.)
    write_data(
      wb            = wb,
      sheet         = sheet,
      data          = data[[i]],
      rc            = c(start_row, start_col),
      row_names     = row_names,
      font_size     = font_size,
      font_name     = font_name,
      fg_fill       = fg_fill,
      border_color  = border_color,
      widths        = widths,
      num_fmt       = num_fmt
    )

    # auto width
    if (auto_width) {
      cw <- if (identical(width_scope, "global_max")) {
        global_widths[[as.character(start_col)]]
      } else {
        .get_col_widths(data[[i]], row_names = row_names)
      }
      openxlsx::setColWidths(
        wb, sheet = sheet,
        cols   = start_col + seq_along(cw) - 1L,
        widths = cw
      )
    }
  }

  invisible(wb)
}

#' Write a list of data.frames to an existing Workbook (one sheet per table)
#'
#' Writes each element of `data` to its **own worksheet** in an existing
#' `openxlsx` workbook. Optionally writes a **title row** above each table and
#' auto-sizes **data columns** (title rows are excluded from width calculation).
#' This function **does not save** the workbook; it returns the modified
#' `Workbook` so you can call `openxlsx::saveWorkbook()` yourself.
#'
#' Sheet names are resolved in this order:
#' 1) `sheet_names` (if provided)
#' 2) `names(data)` (when available)
#' 3) `"Sheet 1"`, `"Sheet 2"`, ...
#'
#' @param data A data.frame, or a (named/unnamed) list of data.frames/matrices.
#'   Matrices are coerced to data.frames. A single object is treated as a
#'   length-1 list.
#' @param wb An existing `openxlsx` `Workbook` (e.g., from
#'   `openxlsx::createWorkbook()` or `openxlsx::loadWorkbook()`).
#' @param rc Integer vector `c(row, col)` giving the common start cell for all
#'   sheets. Default `c(1L, 1L)`.
#' @param row_names Logical; whether to include row names when writing tables.
#'   Default `FALSE`.
#' @param sheet_names Optional character vector of sheet names; must have
#'   `length(sheet_names) == length(data)` if supplied. If `NULL`, falls back
#'   to `names(data)` or `"Sheet {i}"`.
#' @param data_titles Optional character vector of titles to write **above**
#'   each table; must have `length(data_titles) == length(data)` if supplied.
#'   If `NULL`, no titles are written.
#' @param title_size Numeric font size for title rows. Default `14`.
#' @param font_name,font_size,fg_fill,border_color,widths Styling forwarded to
#'   [write_data()]. `widths` is only used when `auto_width = FALSE`.
#' @param auto_width Logical. If `TRUE`, compute per-column widths and set
#'   via `openxlsx::setColWidths()` for **data columns only**. Default `TRUE`.
#' @param num_fmt Optional numeric format(s) applied to data cells (headers
#'   excluded). Either a single character (applied to all numeric columns)
#'   or a named character vector/list mapping column names to formats
#'   (unknown names are ignored).
#'
#' @details
#' - Titles are written via `write_cell()` (bold, left-aligned).
#' - Tables are written via `write_data()` (borders, header/body/footer styles,
#'   numeric formats).
#' - Column widths are set relative to the table’s starting column and do not
#'   include the optional title row.
#'
#' @return The modified `Workbook` (invisibly).
#'
#' @examples
#' \dontrun{
#' wb <- openxlsx::createWorkbook()
#' save_data_wb_split(
#'   data        = list(iris_tab = head(iris), mtcars_tab = head(mtcars)),
#'   wb          = wb,
#'   rc          = c(1, 1),
#'   data_titles = c("Iris (head)", "Mtcars (head)"),
#'   auto_width  = TRUE
#' )
#' openxlsx::saveWorkbook(wb, "by_sheet.xlsx", overwrite = TRUE)
#' }
#'
#' @seealso [save_data_xlsx_split()], [write_data()], [write_cell()]
#'
#' @export
save_data_wb_split <- function(data,
                               wb,
                               rc = c(1L, 1L),
                               row_names = FALSE,
                               sheet_names = NULL,
                               data_titles = NULL,
                               title_size = 14,
                               font_name = getOption("instead.font"),
                               font_size = 11,
                               fg_fill = "#E6E6E7",
                               border_color = "#000000",
                               widths = 8.43,
                               auto_width = TRUE,
                               num_fmt = NULL) {
  if (!inherits(wb, "Workbook"))
    stop("`wb` must be an openxlsx Workbook.", call. = FALSE)

  # normalize input
  if (is.data.frame(data) || is.matrix(data)) data <- list(data)
  if (!is.list(data) || !length(data))
    stop("`data` must be a non-empty data.frame, matrix, or list of them.", call. = FALSE)

  data <- lapply(data, function(d) {
    if (is.matrix(d)) d <- as.data.frame(d)
    if (!is.data.frame(d)) stop("Each element of `data` must be a data.frame/matrix.", call. = FALSE)
    d
  })
  n <- length(data)

  # sheet names
  sheets <- NULL
  if (!is.null(sheet_names)) {
    if (length(sheet_names) != n)
      stop("`sheet_names` must have length = length(data).", call. = FALSE)
    sheets <- as.character(sheet_names)
  } else if (!is.null(names(data))) {
    nm <- names(data)
    nm <- ifelse(is.na(nm) | !nzchar(nm), NA_character_, nm)
    na_idx <- which(is.na(nm))
    if (length(na_idx)) nm[na_idx] <- sprintf("Sheet %d", na_idx)
    sheets <- nm
  } else {
    sheets <- sprintf("Sheet %d", seq_len(n))
  }

  # titles
  titles <- if (is.null(data_titles)) rep(NA_character_, n) else {
    if (length(data_titles) != n) {
      stop("`data_titles` must have length = length(data).", call. = FALSE)
    }
    as.character(data_titles)
  }

  # ensure sheets exist
  existing <- openxlsx::sheets(wb)
  to_add <- setdiff(sheets, existing)
  for (sn in to_add) openxlsx::addWorksheet(wb, sheetName = sn, gridLines = FALSE)

  # validate rc
  if (!is.numeric(rc) || length(rc) != 2L)
    stop("`rc` must be numeric length-2: c(row, col).", call. = FALSE)
  rc <- as.integer(rc)
  start_row0 <- rc[1L]; start_col0 <- rc[2L]

  # styles
  # titles are written via write_cell(); write_data() handles table styling
  # including borders and num_fmt, as in your existing helper.
  for (i in seq_along(data)) {
    sheet <- sheets[[i]]
    start_row <- start_row0
    start_col <- start_col0

    # optional title
    if (!is.na(titles[[i]])) {
      write_cell(
        wb, sheet,
        rc         = c(start_row, start_col),
        value      = titles[[i]],
        bold       = TRUE,
        font_name  = font_name,
        font_size  = title_size,
        h_align    = "left",
        v_align    = "center"
      )
      start_row <- start_row + 1L
    }

    # write table
    write_data(
      wb            = wb,
      sheet         = sheet,
      data          = data[[i]],
      rc            = c(start_row, start_col),
      row_names     = row_names,
      font_size     = font_size,
      font_name     = font_name,
      fg_fill       = fg_fill,
      border_color  = border_color,
      widths        = widths,
      num_fmt       = num_fmt
    )

    # auto width (data columns only)
    if (auto_width) {
      cw <- .get_col_widths(data[[i]], row_names = row_names)
      openxlsx::setColWidths(
        wb, sheet = sheet,
        cols   = start_col + seq_along(cw) - 1L,
        widths = cw
      )
    }
  }

  invisible(wb)
}

#' Write ggplot object(s) into an existing openxlsx Workbook (stacked or placed)
#'
#' Inserts one or more **ggplot** objects into a worksheet of an **existing**
#' `openxlsx` `Workbook`, optionally adding a **title row** above each plot,
#' stacking plots vertically with a configurable spacer, and controlling plot
#' size/resolution. This function **does not save** the workbook; it returns the
#' modified `Workbook` object (use `openxlsx::saveWorkbook()` to write to disk).
#'
#' If `rc` has **length 1**, the first plot is placed at that `(row, col)` and
#' subsequent plots are placed **below it** with `row_spacer` blank rows in
#' between. The next start row is approximated as:
#'
#' \preformatted{
#' start_row_next = start_row
#'                + (title_present ? 1 : 0)
#'                + ceiling(height * rows_per_inch)
#'                + row_spacer
#' }
#'
#' If `rc` has the same length as `plot`, each plot is placed at its own
#' specified top-left cell (no auto-stacking).
#'
#' @param plot A ggplot object, or a non-empty list of ggplot objects.
#' @param wb An existing `openxlsx::Workbook` (e.g., from
#'   `openxlsx::createWorkbook()` or `openxlsx::loadWorkbook()`).
#' @param sheet Target worksheet name. Created if it does not exist.
#'   Default `"Plots"`.
#' @param rc Either a single length-2 numeric vector `c(row, col)` (to stack
#'   plots), or a list of such vectors with the same length as `plot`
#'   (to place each plot individually).
#' @param width,height Plot size in inches passed to `openxlsx::insertPlot()`.
#'   Default `width = 8`, `height = 4`.
#' @param dpi Image resolution passed to `openxlsx::insertPlot()`. Default `300`.
#' @param plot_titles Optional character vector of titles (written **above** each
#'   plot). If `NULL`, `names(plot)` are used when available; otherwise, no title.
#'   Length must equal `length(plot)` when supplied.
#' @param title_size Numeric font size for `plot_titles`. Default `14`.
#' @param font_name Font family for title text. Default `getOption("instead.font")`.
#' @param row_spacer Integer (>= 0). Number of **blank rows** between stacked
#'   plots when `rc` has length 1. Default `2L`.
#' @param rows_per_inch Approximate number of worksheet rows consumed by
#'   **one inch** of plot height (used only in stack mode). Default `5`.
#'
#' @details
#' - Title rows are written via `write_cell()` (bold, left-aligned).
#' - Plots are inserted via `openxlsx::insertPlot()`. The function prints each
#'   ggplot immediately before insertion so the correct plot is captured as the
#'   **last drawn** plot on the active device.
#' - Sheet is created on demand if it does not exist.
#'
#' @return The modified `Workbook` (invisibly), suitable for chaining and for
#'   passing to `openxlsx::saveWorkbook()`.
#'
#' @examples
#' \dontrun{
#' library(ggplot2)
#' library(openxlsx)
#'
#' p1 <- ggplot(mtcars, aes(mpg, wt)) + geom_point()
#' p2 <- ggplot(mtcars, aes(hp, qsec)) + geom_point()
#'
#' # Stack two plots starting at A1 with spacing and titles
#' wb <- openxlsx::createWorkbook()
#' save_plot_wb(
#'   plot        = list("MPG vs WT" = p1, "HP vs QSEC" = p2),
#'   wb          = wb,
#'   sheet       = "Plots",
#'   rc          = c(1L, 1L),
#'   title_size  = 16,
#'   width = 8, height = 4, dpi = 300,
#'   row_spacer  = 2L, rows_per_inch = 5
#' )
#' openxlsx::saveWorkbook(wb, "plots_stacked.xlsx", overwrite = TRUE)
#' }
#'
#' @seealso [save_plot_xlsx()], [save_plot_wb_split()], [openxlsx::insertPlot()]
#'
#' @export
save_plot_wb <- function(plot, wb, sheet = "Plots",
                         rc = c(1L, 1L),
                         width = 8, height = 4, dpi = 300,
                         plot_titles = NULL, title_size = 14,
                         font_name = getOption("instead.font"),
                         row_spacer = 2L, rows_per_inch = 5) {
  if (!inherits(wb, "Workbook"))
    stop("`wb` must be an openxlsx Workbook.", call. = FALSE)

  # normalize plots
  if (inherits(plot, "ggplot")) {
    plots <- list(plot)
  } else if (is.list(plot) && length(plot) &&
             all(vapply(plot, function(p) inherits(p, "ggplot"), logical(1L)))) {
    plots <- plot
  } else {
    stop("`plot` must be a ggplot or a non-empty list of ggplots.", call. = FALSE)
  }

  # titles
  titles <- rep(NA_character_, length(plots))
  if (!is.null(plot_titles) && length(plot_titles) == length(plots)) {
    titles <- as.character(plot_titles)
  } else if (!is.null(names(plots))) {
    nm <- names(plots)
    titles <- ifelse(!is.na(nm) & nzchar(nm), nm, NA_character_)
  }

  # ensure sheet
  if (!(sheet %in% openxlsx::sheets(wb))) {
    openxlsx::addWorksheet(wb, sheetName = sheet, gridLines = FALSE)
  }

  # rc -> list of starts
  if (is.numeric(rc) && length(rc) == 2L) {
    starts <- vector("list", length(plots))
    cur <- as.integer(rc)
    for (i in seq_along(plots)) {
      starts[[i]] <- cur
      h_rows <- (if (!is.na(titles[[i]])) 1L else 0L) +
        ceiling(height * rows_per_inch) + as.integer(row_spacer)
      cur <- c(cur[1] + h_rows, cur[2])
    }
  } else if (is.list(rc) && length(rc) == length(plots)) {
    starts <- lapply(rc, function(x) {
      if (is.numeric(x) && length(x) == 2L) {
        as.integer(x)
      } else {
        stop("Each `rc` element must be numeric length-2: c(row, col).",
             call. = FALSE)
      }
    })
  } else {
    stop("`rc` must be c(row, col) or a list of c(row, col) of length(plot).",
         call. = FALSE)
  }

  title_style <- openxlsx::createStyle(
    textDecoration = "bold",
    fontSize = title_size,
    fontName = font_name
  )

  # write
  for (i in seq_along(plots)) {
    start_row <- starts[[i]][1L]
    start_col <- starts[[i]][2L]

    if (!is.na(titles[[i]])) {
      write_cell(
        wb, sheet,
        rc         = c(start_row, start_col),
        value      = titles[[i]],
        bold       = TRUE,
        italic     = FALSE,
        underline  = FALSE,
        strikeout  = FALSE,
        font_name  = font_name,
        font_size  = title_size,
        font_color = NULL,
        fg_fill    = NULL,
        h_align    = "left",
        v_align    = "center"
      )
      start_row <- start_row + 1L
    }

    # insertPlot uses last drawn plot on device
    print(plots[[i]])
    openxlsx::insertPlot(
      wb, sheet = sheet,
      startRow = start_row, startCol = start_col,
      width = width, height = height, dpi = dpi
    )
  }

  invisible(wb)
}

#' Write ggplot object(s) to separate worksheets in an existing Workbook
#'
#' Each ggplot is written to its **own sheet** in an existing `openxlsx`
#' `Workbook`. A common start cell `rc` is used for all sheets. Optional
#' per-sheet title rows. This function **does not save** the workbook; it
#' returns the modified `Workbook` for you to save.
#'
#' Sheet names are resolved in this order:
#' 1) `sheet_names` (if provided)
#' 2) `names(plot)` (when available)
#' 3) `"Sheet 1"`, `"Sheet 2"`, ...
#'
#' @inheritParams save_plot_wb
#' @param sheet_names Optional character vector of sheet names; must have
#'   `length(sheet_names) == length(plot)` when supplied. If `NULL`, falls back
#'   to `names(plot)` or `"Sheet {i}"`.
#'
#' @details
#' - Title rows are written via `write_cell()` (bold, left-aligned).
#' - Plots are inserted via `openxlsx::insertPlot()`.
#'
#' @return The modified `Workbook` (invisibly).
#'
#' @examples
#' \dontrun{
#' library(ggplot2)
#'
#' wb <- openxlsx::createWorkbook()
#' p1 <- ggplot(mtcars, aes(mpg, wt)) + geom_point()
#' p2 <- ggplot(mtcars, aes(hp, qsec)) + geom_point()
#'
#' save_plot_wb_split(
#'   plot        = list("Scatter MPG~WT" = p1, "Scatter HP~QSEC" = p2),
#'   wb          = wb,
#'   rc          = c(1L, 1L),
#'   plot_titles = c("MPG vs Weight", "HP vs QSEC"),
#'   title_size  = 16,
#'   width = 8, height = 4, dpi = 300
#' )
#' openxlsx::saveWorkbook(wb, "plots_split.xlsx", overwrite = TRUE)
#' }
#'
#' @seealso [save_plot_xlsx_split()], [save_plot_wb()], [openxlsx::insertPlot()]
#'
#' @export
save_plot_wb_split <- function(plot, wb,
                               rc = c(1L, 1L),
                               sheet_names = NULL,
                               width = 8, height = 4, dpi = 300,
                               plot_titles = NULL, title_size = 14,
                               font_name = getOption("instead.font")) {
  if (!inherits(wb, "Workbook"))
    stop("`wb` must be an openxlsx Workbook.", call. = FALSE)

  # normalize plots
  if (inherits(plot, "ggplot")) {
    plots <- list(plot)
  } else if (is.list(plot) && length(plot) &&
             all(vapply(plot, function(p) inherits(p, "ggplot"), logical(1L)))) {
    plots <- plot
  } else stop("`plot` must be a ggplot or a list of ggplots.", call. = FALSE)

  n <- length(plots)

  if (!(is.numeric(rc) && length(rc) == 2L))
    stop("`rc` must be numeric length-2: c(row, col).", call. = FALSE)
  rc <- as.integer(rc)
  start_row0 <- rc[1L]; start_col0 <- rc[2L]

  # sheet names
  sheets <- NULL
  if (!is.null(sheet_names)) {
    if (length(sheet_names) != n)
      stop("`sheet_names` must have length = length(plot).", call. = FALSE)
    sheets <- as.character(sheet_names)
  } else if (!is.null(names(plots))) {
    nm <- names(plots)
    nm[!nzchar(ifelse(is.na(nm), "", nm))] <- NA_character_
    na_idx <- which(is.na(nm))
    if (length(na_idx)) nm[na_idx] <- sprintf("Sheet %d", na_idx)
    sheets <- nm
  } else {
    sheets <- sprintf("Sheet %d", seq_len(n))
  }
  if (anyDuplicated(sheets))
    stop("`sheet_names` (or names(plot)) contain duplicates.", call. = FALSE)

  # titles
  titles <- if (is.null(plot_titles)) {
    rep_len(NA_character_, n)
  } else {
    if (length(plot_titles) != n) {
      stop("`plot_titles` must have length = length(plot).", call. = FALSE)
    }
    as.character(plot_titles)
  }

  title_style <- openxlsx::createStyle(
    textDecoration = "bold",
    fontSize = title_size,
    fontName = font_name
  )

  # write each plot to its sheet
  existing <- openxlsx::sheets(wb)
  to_add <- setdiff(sheets, existing)
  for (sn in to_add)
    openxlsx::addWorksheet(wb, sheetName = sn, gridLines = FALSE)

  for (i in seq_along(plots)) {
    sheet <- sheets[[i]]
    start_row <- start_row0
    start_col <- start_col0

    if (!is.na(titles[[i]])) {
      write_cell(
        wb, sheet,
        rc         = c(start_row, start_col),
        value      = titles[[i]],
        bold       = TRUE,
        italic     = FALSE,
        underline  = FALSE,
        strikeout  = FALSE,
        font_name  = font_name,
        font_size  = title_size,
        font_color = NULL,
        fg_fill    = NULL,
        h_align    = "left",
        v_align    = "center"
      )
      start_row <- start_row + 1L
    }

    print(plots[[i]])
    openxlsx::insertPlot(
      wb, sheet = sheet,
      startRow = start_row, startCol = start_col,
      width = width, height = height, dpi = dpi
    )
  }

  invisible(wb)
}

#' Write image file(s) to a single worksheet in an existing Workbook (stacked)
#'
#' Inserts one or more image files (e.g., PNG/JPG) into the same worksheet of
#' an existing `openxlsx` `Workbook`, stacked vertically. Optional **titles**
#' above each image and configurable spacing between images. This function
#' **does not save** the workbook; it returns the modified `Workbook`.
#'
#' @param image Character vector (or list) of file paths.
#' @param wb    An existing `openxlsx::Workbook`.
#' @param sheet Worksheet name (created if missing). Default `"Images"`.
#' @param rc    Length-2 integer `c(row, col)` for the first image’s top-left cell.
#'   Default `c(1L, 1L)`.
#' @param width,height,dpi Passed to `openxlsx::insertImage()`. Defaults:
#'   `width = 12`, `height = 6`, `dpi = 300`.
#' @param image_titles Optional character vector of titles (length = `length(image)`),
#'   written **above** each image. If `NULL`, uses `names(image)` when available.
#' @param title_size Numeric font size for titles. Default `14`.
#' @param font_name Font family for title text. Default `getOption("instead.font")`.
#' @param row_spacer,rows_per_inch Integers controlling vertical spacing in stack mode.
#'   `rows_per_inch` is the approximate number of worksheet rows consumed by **one inch**
#'   of image height. Defaults `row_spacer = 2L`, `rows_per_inch = 5L`.
#'
#' @details
#' - Title rows are written via `write_cell()` (bold, left-aligned).
#' - Images are inserted via `openxlsx::insertImage()`.
#' - The next start row in stack mode advances by
#'   `ceiling(height * rows_per_inch) + row_spacer` (plus 1 if a title row is present).
#'
#' @return The modified `Workbook` (invisibly).
#'
#' @examples
#' \dontrun{
#' wb <- openxlsx::createWorkbook()
#' save_image_wb(
#'   image        = c("First" = "plot1.png", "Second" = "plot2.jpg"),
#'   wb           = wb,
#'   sheet        = "MyImages",
#'   rc           = c(1L, 1L),
#'   image_titles = c("First Image", "Second Image"),
#'   width        = 6, height = 4, dpi = 300,
#'   row_spacer   = 2L
#' )
#' openxlsx::saveWorkbook(wb, "images_stacked.xlsx", overwrite = TRUE)
#' }
#'
#' @seealso [save_image_xlsx()], [save_image_wb_split()], [openxlsx::insertImage()]
#'
#' @export
save_image_wb <- function(image, wb, sheet = "Images",
                          rc = c(1L, 1L),
                          width = 12, height = 6, dpi = 300,
                          image_titles = NULL, title_size = 14,
                          font_name = getOption("instead.font"),
                          row_spacer = 2L, rows_per_inch = 5L) {
  if (!inherits(wb, "Workbook"))
    stop("`wb` must be an openxlsx Workbook.", call. = FALSE)

  if (is.list(image))
    image <- unlist(image, use.names = TRUE)

  if (!is.character(image) || !length(image))
    stop("`image` must be a character vector of file paths.", call. = FALSE)

  missing <- image[!file.exists(image)]
  if (length(missing))
    stop("File(s) not found: ", paste(shQuote(missing), collapse = ", "))

  titles <- if (is.null(image_titles)) {
    nm <- names(image)
    if (is.null(nm)) {
      rep_len(NA_character_, length(image))
    } else {
      ifelse(nzchar(nm), nm, NA_character_)
    }
  } else {
    if (length(image_titles) != length(image)) {
      stop("`image_titles` must have length = length(image) (or be NULL).",
           call. = FALSE)
    }
    as.character(image_titles)
  }

  if (!(sheet %in% openxlsx::sheets(wb))) {
    openxlsx::addWorksheet(wb, sheetName = sheet, gridLines = FALSE)
  }

  title_style <- openxlsx::createStyle(
    textDecoration = "bold",
    fontSize = title_size,
    fontName = font_name
  )

  start_row <- as.integer(rc[1L])
  start_col <- as.integer(rc[2L])
  advance_rows <- function(h_in) as.integer(ceiling(h_in * rows_per_inch))

  for (i in seq_along(image)) {
    if (!is.na(titles[[i]])) {
      write_cell(
        wb, sheet,
        rc         = c(start_row, start_col),
        value      = titles[[i]],
        bold       = TRUE,
        italic     = FALSE,
        underline  = FALSE,
        strikeout  = FALSE,
        font_name  = font_name,
        font_size  = title_size,
        font_color = NULL,
        fg_fill    = NULL,
        h_align    = "left",
        v_align    = "center"
      )
      start_row <- start_row + 1L
    }

    openxlsx::insertImage(
      wb, sheet = sheet, file = image[[i]],
      width = width, height = height,
      startRow = start_row, startCol = start_col, dpi = dpi
    )

    start_row <- start_row + advance_rows(height) + as.integer(row_spacer)
  }

  invisible(wb)
}

#' Write image file(s) to separate worksheets in an existing Workbook
#'
#' Each image file is inserted into its **own worksheet** in an existing
#' `openxlsx` `Workbook`. A common start cell `rc` is used for all sheets.
#' Optional per-sheet title rows. This function **does not save** the workbook;
#' it returns the modified `Workbook`.
#'
#' Sheet names are resolved in this order:
#' 1) `sheet_names` (if provided)
#' 2) `names(image)` (when available)
#' 3) `"Sheet 1"`, `"Sheet 2"`, ...
#'
#' @param image Character vector (or list) of file paths.
#' @param wb    An existing `openxlsx::Workbook`.
#' @param rc    Length-2 integer vector `c(row, col)` used on all sheets.
#'   Default `c(1L, 1L)`.
#' @param sheet_names Optional character vector of sheet names (length = `length(image)`).
#'   If `NULL`, falls back to `names(image)` or `"Sheet {i}"`.
#' @param image_titles Optional character vector of titles (length = `length(image)`),
#'   written **above** each image. If `NULL`, no titles are written.
#' @param title_size Numeric font size for titles. Default `14`.
#' @param font_name Font family for titles. Default `getOption("instead.font")`.
#' @param width,height,dpi Passed to `openxlsx::insertImage()`. Defaults
#'   `width = 12`, `height = 6`, `dpi = 300`.
#'
#' @details
#' - Title rows are written via `write_cell()` (bold, left-aligned).
#' - Images are inserted via `openxlsx::insertImage()`.
#'
#' @return The modified `Workbook` (invisibly).
#'
#' @examples
#' \dontrun{
#' wb <- openxlsx::createWorkbook()
#' save_image_wb_split(
#'   image        = c("Left plot" = "plot1.png", "Right plot" = "plot2.jpg"),
#'   wb           = wb,
#'   rc           = c(2L, 2L),
#'   sheet_names  = c("Left", "Right"),
#'   image_titles = c("Left Image", "Right Image"),
#'   width        = 10, height = 5, dpi = 300
#' )
#' openxlsx::saveWorkbook(wb, "images_split.xlsx", overwrite = TRUE)
#' }
#'
#' @seealso [save_image_xlsx_split()], [save_image_wb()], [openxlsx::insertImage()]
#'
#' @export
save_image_wb_split <- function(image, wb,
                                rc = c(1L, 1L),
                                sheet_names = NULL,
                                image_titles = NULL, title_size = 14,
                                font_name = getOption("instead.font"),
                                width = 12, height = 6, dpi = 300) {
  if (!inherits(wb, "Workbook"))
    stop("`wb` must be an openxlsx Workbook.", call. = FALSE)

  if (is.list(image))
    image <- unlist(image, use.names = TRUE)

  if (!is.character(image) || !length(image))
    stop("`image` must be a character vector of file paths.", call. = FALSE)

  n <- length(image)

  missing <- image[!file.exists(image)]
  if (length(missing))
    stop("File(s) not found: ", paste(shQuote(missing), collapse = ", "))

  if (!(is.numeric(rc) && length(rc) == 2L))
    stop("`rc` must be numeric length-2: c(row, col).", call. = FALSE)

  rc <- as.integer(rc)
  start_row0 <- rc[1L]; start_col0 <- rc[2L]

  # sheet names
  sheets <- NULL
  if (!is.null(sheet_names)) {
    if (length(sheet_names) != n) {
      stop("`sheet_names` must have the same length as `image`.", call. = FALSE)
    }
    sheets <- as.character(sheet_names)
  } else if (!is.null(names(image))) {
    nm <- names(image)
    nm[!nzchar(ifelse(is.na(nm), "", nm))] <- NA_character_
    na_idx <- which(is.na(nm))
    if (length(na_idx)) {
      nm[na_idx] <- sprintf("Sheet %d", na_idx)
    }
    sheets <- nm
  } else {
    sheets <- sprintf("Sheet %d", seq_len(n))
  }

  if (anyDuplicated(sheets))
    stop("`sheet_names` (or names(image)) contain duplicates.", call. = FALSE)

  # titles
  titles <- if (is.null(image_titles)) {
    rep_len(NA_character_, n)
  } else {
    if (length(image_titles) != n) {
      stop("`image_titles` must have length = length(image) (or be NULL).", call. = FALSE)
    }
    as.character(image_titles)
  }

  to_add <- setdiff(sheets, openxlsx::sheets(wb))
  for (sn in to_add)
    openxlsx::addWorksheet(wb, sheetName = sn, gridLines = FALSE)

  title_style <- openxlsx::createStyle(
    textDecoration = "bold",
    fontSize = title_size,
    fontName = font_name
  )

  for (i in seq_along(image)) {
    sheet <- sheets[[i]]
    start_row <- start_row0
    start_col <- start_col0

    if (!is.na(titles[[i]])) {
      write_cell(
        wb, sheet,
        rc         = c(start_row, start_col),
        value      = titles[[i]],
        bold       = TRUE,
        italic     = FALSE,
        underline  = FALSE,
        strikeout  = FALSE,
        font_name  = font_name,
        font_size  = title_size,
        font_color = NULL,
        fg_fill    = NULL,
        h_align    = "left",
        v_align    = "center"
      )
      start_row <- start_row + 1L
    }

    openxlsx::insertImage(
      wb, sheet = sheet, file = image[[i]],
      width = width, height = height,
      startRow = start_row, startCol = start_col, dpi = dpi
    )
  }

  invisible(wb)
}

# Save xlsx files ---------------------------------------------------------

#' Save multiple data.frames to one sheet (stacked), with optional titles and auto-widths
#'
#' Writes one or more rectangular tables to a single Excel worksheet using
#' `openxlsx`, stacking them vertically with a configurable blank spacer
#' between tables. Optionally writes a **title row** above each table and
#' auto-sizes the **data columns** (titles are excluded from width calculation).
#'
#' If `rc` has **length 1**, the first table is written at that `(row, col)`,
#' and subsequent tables are placed **below it**, leaving `row_spacer` blank
#' row(s) between tables. The next top-left row is computed as:
#'
#' \preformatted{
#' start_row_next = start_row
#'                + title_height              # 1 if a title is written, else 0
#'                + 1L                        # header row
#'                + nrow(df)                  # data rows
#'                + row_spacer                # blank rows between tables
#' }
#'
#' If `rc` has the same length as `data`, each table is written to its own
#' specified top-left cell, with no automatic stacking.
#'
#' Matrices are accepted and coerced to data.frames. Vectors or character
#' scalars are not supported.
#'
#' @param data A data.frame, matrix, or a list of them. If a single object
#'   is supplied it is treated as length-1 list.
#' @param file Path to the `.xlsx` file. Created if it does not exist.
#' @param sheet Target sheet name (created if missing). Default `"Sheet 1"`.
#' @param rc Either a single length-2 integer vector `c(row, col)` (to stack),
#'   or a list of such vectors with the same length as `data` (to place individually).
#' @param row_spacer Integer (>= 0). Number of **blank rows between stacked tables**
#'   when `rc` has length 1. Default `2L`.
#' @param data_titles Optional character vector of table titles. If provided and
#'   length matches `length(data)`, these are written **above** each table.
#'   If `data_titles` is `NULL`, `names(data)` are used when available.
#'   If neither is available, no title is written for that table.
#' @param title_size Numeric font size for titles. Default `14`.
#' @param row_names Logical; forwarded to [write_data()] to include row names
#'   in the written table. Default `FALSE`.
#' @param font_name,font_size,fg_fill,border_color,widths Styling; forwarded to
#'   [write_data()]. `widths` is used when `auto_width = FALSE`.
#' @param auto_width Logical. If `TRUE`, compute per-column widths from the
#'   underlying data (and optional row-name column) and call
#'   `openxlsx::setColWidths()` **for data columns only**. Title rows are ignored.
#'   Default `TRUE`.
#' @param width_scope One of `"per_table"` or `"global_max"`. When `auto_width = TRUE`:
#'   - `"per_table"` sets column widths independently **after each table**.
#'   - `"global_max"` first scans all tables that start at the **same start column**,
#'     takes the **maximum width per column** across them, then applies widths **once**,
#'     preventing later tables from overwriting the widths set by earlier ones.
#'   Default `"global_max"`.
#' @param num_fmt Optional numeric format(s) applied to data cells (headers
#'   excluded). Accepts either:
#'   - A **single character format string** applied to all numeric cells
#'     (e.g. `"0.00"` for 2 decimals, `"#,##0"` for thousands separator,
#'     `"0.00%"` for percentages).
#'   - A **named character vector or list** mapping column names to formats,
#'     e.g. `c(value = "#,##0", rate = "0.0%")`. Columns not listed are left
#'     unchanged. Unknown column names are silently ignored.
#' @param overwrite Logical; passed to `openxlsx::saveWorkbook()`. Default `FALSE`.
#'
#' @return (Invisibly) the path to `file`. Side-effects: creates/updates an Excel file.
#'
#' @section Notes:
#' - This function uses an existing helper [write_data()] for writing a single table
#'   with styling. Column auto-sizing (when enabled) ignores the title row.
#' - Column widths are set relative to the table's starting column. If `row_names = TRUE`,
#'   the first (row-name) column is included in auto-width calculation and width setting.
#'
#' @examples
#' \dontrun{
#' data1 <- head(iris)
#' data2 <- head(mtcars)
#'
#' # Stack two tables on one sheet (A1 start), with titles & unified widths
#' save_data_xlsx(
#'   data        = list(iris_head = data1, mtcars_head = data2),
#'   data_titles = c("iris (head)", "mtcars (head)"),
#'   file        = "data.xlsx",
#'   sheet       = "Summary",
#'   rc          = list(c(1, 1)),
#'   row_spacer  = 2L,
#'   auto_width  = TRUE,
#'   width_scope = "global_max",
#'   overwrite   = TRUE
#' )
#'
#' # Place at specific cells (no stacking)
#' save_data_xlsx(
#'   data        = list(A = data1, B = data2),
#'   data_titles = c("Block A", "Block B"),
#'   file        = "block.xlsx",
#'   sheet       = "Placed",
#'   rc          = list(c(1, 1), c(1, 10)),  # A1 and J1
#'   auto_width  = TRUE,
#'   width_scope = "per_table",
#'   overwrite   = TRUE
#' )
#' }
#'
#' @seealso [save_data_xlsx_split()], [write_data()]
#'
#' @export
save_data_xlsx <- function(data, file, sheet = "Data",
                           rc = list(c(1L, 1L)),
                           row_spacer = 2L,
                           data_titles = NULL,
                           title_size = 14,
                           row_names = FALSE,
                           font_name = getOption("instead.font"),
                           font_size = 11,
                           fg_fill = "#E6E6E7",
                           border_color = "#000000",
                           widths = 8.43,
                           auto_width = TRUE,
                           width_scope = c("global_max", "per_table"),
                           num_fmt = NULL,
                           overwrite = FALSE) {
  width_scope <- match.arg(width_scope)

  # create or load workbook
  wb <- if (file.exists(file)) {
    openxlsx::loadWorkbook(file)
  } else {
    openxlsx::createWorkbook()
  }

  # delegate to the external `save_data_wb()`
  wb <- save_data_wb(
    data          = data,
    wb            = wb,
    sheet         = sheet,
    rc            = rc,
    row_spacer    = row_spacer,
    data_titles   = data_titles,
    title_size    = title_size,
    row_names     = row_names,
    font_name     = font_name,
    font_size     = font_size,
    fg_fill       = fg_fill,
    border_color  = border_color,
    widths        = widths,
    auto_width    = auto_width,
    width_scope   = width_scope,
    num_fmt       = num_fmt
  )

  # save to disk
  openxlsx::saveWorkbook(wb, file = file, overwrite = overwrite)
  invisible(file)
}


#' Save a list of data.frames to an Excel workbook (one sheet per table)
#'
#' Write a data frame or a list of data frames to an Excel workbook using
#' `openxlsx`. When a list is supplied, each element is written to its **own
#' worksheet**. Sheet names are resolved as:
#'   1. `sheet_names` argument (if provided),
#'   2. else `names(data)` if available,
#'   3. else `"Sheet 1"`, `"Sheet 2"`, ...
#'
#' Optionally writes a title row above each table and auto-sizes data columns
#' (title excluded from width calculation).
#'
#' @param data A data.frame, or a named/unnamed list of data.frames (or matrix).
#' @param file Path to the target `.xlsx` workbook.
#' @param rc A length-2 integer vector giving the common starting row and column
#'   for every sheet (default `c(1L, 1L)`).
#' @param row_names Logical; whether to include row names. Default `FALSE`.
#' @param sheet_names Optional character vector of sheet names. If `NULL`,
#'   falls back to `names(data)` or autogenerated names.
#' @param data_titles Optional character vector of table titles, length `length(data)`.
#'   Written **above** each table. If `NULL`, no titles are written.
#' @param title_size Numeric font size for title rows. Default `14`.
#' @param font_name,font_size,fg_fill,border_color,widths Styling forwarded to
#'   [write_data()].
#' @param auto_width Logical. If `TRUE`, compute per-column widths and apply to data columns.
#'   Default `TRUE`.
#' @param num_fmt Optional numeric format(s) applied to data cells (headers
#'   excluded). Accepts either:
#'   - A **single character format string** applied to all numeric cells
#'     (e.g. `"0.00"` for 2 decimals, `"#,##0"` for thousands separator,
#'     `"0.00%"` for percentages).
#'   - A **named character vector or list** mapping column names to formats,
#'     e.g. `c(value = "#,##0", rate = "0.0%")`. Columns not listed are left
#'     unchanged. Unknown column names are silently ignored.
#' @param overwrite Logical; if `TRUE`, overwrite an existing workbook. Default `FALSE`.
#'
#' @return (Invisibly) the path to `file`. Side-effects: creates/updates an Excel file.
#'
#' @examples
#' \dontrun{
#' d1 <- head(iris)
#' d2 <- head(mtcars)
#'
#' # Case 1: explicit sheet_names
#' save_data_xlsx_split(
#'   data        = list(d1, d2),
#'   sheet_names = c("iris_data", "mtcars_data"),
#'   data_titles = c("Iris Head", "Mtcars Head"),
#'   file        = "sheets_named.xlsx",
#'   overwrite   = TRUE
#' )
#'
#' # Case 2: names(data) used as sheet names
#' save_data_xlsx_split(
#'   data       = list("iris_tab" = d1, "mtcars_tab" = d2),
#'   file       = "sheets_from_names.xlsx",
#'   overwrite  = TRUE
#' )
#'
#' # Case 3: unnamed list -> Sheet 1, Sheet 2
#' save_data_xlsx_split(
#'   data       = list(d1, d2),
#'   file       = "sheets_auto.xlsx",
#'   overwrite  = TRUE
#' )
#' }
#' @export
save_data_xlsx_split <- function(data, file, rc = c(1L, 1L),
                                 row_names = FALSE,
                                 sheet_names = NULL,
                                 data_titles = NULL,
                                 title_size = 14,
                                 font_name = getOption("instead.font"),
                                 font_size = 11,
                                 fg_fill = "#E6E6E7",
                                 border_color = "#000000",
                                 widths = 8.43,
                                 auto_width = TRUE,
                                 num_fmt = NULL,
                                 overwrite = FALSE) {
  # open or create workbook
  wb <- if (file.exists(file)) {
    openxlsx::loadWorkbook(file)
  } else {
    openxlsx::createWorkbook()
  }

  # delegate the writing work
  wb <- save_data_wb_split(
    data         = data,
    wb           = wb,
    rc           = rc,
    row_names    = row_names,
    sheet_names  = sheet_names,
    data_titles  = data_titles,
    title_size   = title_size,
    font_name    = font_name,
    font_size    = font_size,
    fg_fill      = fg_fill,
    border_color = border_color,
    widths       = widths,
    auto_width   = auto_width,
    num_fmt      = num_fmt
  )

  # save to disk
  openxlsx::saveWorkbook(wb, file = file, overwrite = overwrite)
  invisible(file)
}

#' Save ggplot object(s) to a single Excel sheet (stack or fixed positions)
#'
#' Writes one or more ggplots to a single worksheet using `openxlsx`.
#' - If `rc` is a length-2 integer vector (e.g., `c(1L, 1L)`), plots are
#'   stacked vertically with an optional title row and `row_spacer` blank rows
#'   between plots. The next start row is approximated as:
#'     1 (title if present) + ceiling(height * rows_per_inch) + row_spacer.
#' - If `rc` is a list of `c(row, col)` pairs with the same length as `plot`,
#'   each plot is placed at its specified position (no stacking).
#'
#' @param plot A ggplot object or a non-empty list of ggplot objects.
#' @param file Path to the target `.xlsx` workbook.
#' @param sheet Target sheet name (created if missing). Default `"Plots"`.
#' @param rc Either a single `c(row, col)` (stack mode) or a list of such
#'   vectors (fixed positions).
#' @param width,height Plot size in inches (passed to `openxlsx::insertPlot()`).
#' @param dpi Image resolution for `openxlsx::insertPlot()`. Default `300`.
#' @param plot_titles Optional character vector of titles (length = length(plot)).
#'   If `NULL`, `names(plot)` are used when available; otherwise no titles.
#' @param title_size Numeric font size for titles. Default `14`.
#' @param font_name Font family for title text (passed to openxlsx style).
#'   Default `getOption("instead.font")`.
#' @param row_spacer Non-negative integer. Blank rows between stacked plots.
#'   Default `2L`.
#' @param rows_per_inch Approximate number of worksheet rows consumed per one
#'   inch of plot height in stack mode. Default `5`.
#' @param overwrite Overwrite existing workbook on save? Default `FALSE`.
#'
#' @return (Invisibly) the path to `file`. Side-effects: creates/updates an Excel file.
#'
#' @examples
#' \dontrun{
#' library(ggplot2)
#'
#' p1 <- ggplot(mtcars, aes(mpg, wt)) + geom_point()
#' p2 <- ggplot(mtcars, aes(hp, qsec)) + geom_point()
#'
#' # 1) Fixed positions (list rc): place two plots at different rows
#' save_plot_xlsx(
#'   plot   = list("Scatter 1" = p1, "Scatter 2" = p2),
#'   file   = "plots_positions.xlsx",
#'   sheet  = "Plots",
#'   rc     = list(c(1, 1), c(30, 1)),
#'   plot_titles = NULL,      # uses names(plot) as titles
#'   title_size  = 16,
#'   width = 8, height = 4, dpi = 300,
#'   overwrite = TRUE
#' )
#'
#' # 2) Stack mode (single rc): place plots top-to-bottom with spacing
#' save_plot_xlsx(
#'   plot   = list(p1, p2),
#'   file   = "plots_stacked.xlsx",
#'   sheet  = "Plots",
#'   rc     = c(1L, 1L),
#'   plot_titles = c("MPG vs Weight", "HP vs QSEC"),
#'   row_spacer  = 2L,
#'   width = 8, height = 4, dpi = 300,
#'   overwrite = TRUE
#' )
#' }
#'
#' @seealso [save_plot_wb()], [openxlsx::insertPlot()]
#'
#' @export
save_plot_xlsx <- function(plot, file, sheet = "Plots",
                           rc = c(1L, 1L),
                           width = 8, height = 4, dpi = 300,
                           plot_titles = NULL, title_size = 14,
                           font_name = getOption("instead.font"),
                           row_spacer = 2L, rows_per_inch = 5,
                           overwrite = FALSE) {
  wb <- if (file.exists(file)) {
    openxlsx::loadWorkbook(file)
  } else {
    openxlsx::createWorkbook()
  }

  wb <- save_plot_wb(
    plot         = plot,
    wb           = wb,
    sheet        = sheet,
    rc           = rc,
    width        = width,
    height       = height,
    dpi          = dpi,
    plot_titles  = plot_titles,
    title_size   = title_size,
    font_name    = font_name,
    row_spacer   = row_spacer,
    rows_per_inch = rows_per_inch
  )

  openxlsx::saveWorkbook(wb, file = file, overwrite = overwrite)
  invisible(file)
}


#' Save ggplot object(s) to seperate Excel sheets (one plot per sheet)
#'
#' Writes each ggplot to its own worksheet using `openxlsx`. A common start cell
#' (`rc`) is applied to all sheets. Sheet names and per-sheet title rows are optional.
#'
#' Sheet name resolution priority:
#' 1) `sheet_names` argument (if provided),
#' 2) `names(plot)` (if present and non-empty),
#' 3) "Sheet 1", "Sheet 2", ...
#'
#' @param plot A ggplot object or a non-empty list of ggplot objects.
#' @param file Path to the target `.xlsx` workbook.
#' @param rc Common start cell c(row, col) used for all sheets. Default c(1L, 1L).
#' @param sheet_names Optional character vector of sheet names; length must equal length(plot).
#' @param width,height Plot size in inches for openxlsx::insertPlot().
#' @param dpi Image resolution for openxlsx::insertPlot(). Default 300.
#' @param plot_titles Optional character vector of titles (one per plot); written one row above the plot.
#' @param title_size Numeric font size for titles. Default 14.
#' @param font_name Font family for title text. Default getOption("instead.font").
#' @param overwrite Logical; overwrite existing workbook on save? Default FALSE.
#'
#' @return (Invisibly) the path to `file`. Side-effects: creates/updates an Excel file.
#'
#' @examples
#' \dontrun{
#' library(ggplot2)
#'
#' p1 <- ggplot(mtcars, aes(mpg, wt)) + geom_point()
#' p2 <- ggplot(mtcars, aes(hp, qsec)) + geom_point()
#'
#' save_plot_xlsx_split(
#'   plot        = list(Fit1 = p1, Fit2 = p2),
#'   file        = "plots_split.xlsx",
#'   rc          = c(1L, 1L),
#'   sheet_names = c("Scatter MPG~WT", "Scatter HP~QSEC"),
#'   plot_titles = c("MPG vs Weight", "HP vs QSEC"),
#'   title_size  = 16,
#'   width = 8, height = 4, dpi = 300,
#'   overwrite   = TRUE
#' )
#' }
#'
#' @seealso [save_plot_wb_split()], [openxlsx::insertPlot()]
#'
#' @export
save_plot_xlsx_split <- function(plot, file,
                                 rc = c(1L, 1L),
                                 sheet_names = NULL,
                                 width = 8, height = 4, dpi = 300,
                                 plot_titles = NULL, title_size = 14,
                                 font_name = getOption("instead.font"),
                                 overwrite = FALSE) {
  wb <- if (file.exists(file)) {
    openxlsx::loadWorkbook(file)
  } else {
    openxlsx::createWorkbook()
  }

  wb <- save_plot_wb_split(
    plot        = plot,
    wb          = wb,
    rc          = rc,
    sheet_names = sheet_names,
    width       = width,
    height      = height,
    dpi         = dpi,
    plot_titles = plot_titles,
    title_size  = title_size,
    font_name   = font_name
  )

  openxlsx::saveWorkbook(wb, file = file, overwrite = overwrite)
  invisible(file)
}

#' Save image file(s) to a single Excel sheet (stacked vertically)
#'
#' Inserts one or more image files into the same worksheet, stacking them.
#' Optional titles per image and row spacing between images.
#'
#' @param image Character vector (or list) of file paths.
#' @param file Path to target .xlsx.
#' @param sheet Sheet name to write into. Default "Images".
#' @param rc Length-2 integer vector c(row, col) for the first image.
#' @param width,height,dpi Passed to openxlsx::insertImage().
#' @param image_titles Optional character vector of titles (one per image); written one row above the image.
#' @param title_size Numeric font size for titles. Default 14.
#' @param font_name Font family for title text. Default getOption("instead.font").
#' @param row_spacer Non-negative integer: blank rows between images. Default 2.
#' @param rows_per_inch Approx rows per 1 inch of image height. Default 5.
#' @param overwrite Overwrite workbook if TRUE.
#'
#' @return (Invisibly) the path to `file`. Side-effects: creates/updates an Excel file.
#'
#' @examples
#' \dontrun{
#' # Example image files (replace with actual PNG/JPG paths)
#' img1 <- "plot1.png"
#' img2 <- "plot2.jpg"
#'
#' # Stack two images vertically into one sheet
#' save_image_xlsx(
#'   image        = c("First" = img1, "Second" = img2),
#'   file         = "images_stacked.xlsx",
#'   sheet        = "MyImages",
#'   rc           = c(1L, 1L),   # start at A1
#'   image_titles = c("First Image Title", "Second Image Title"),
#'   title_size   = 16,
#'   width        = 6, height = 4, dpi = 300,
#'   row_spacer   = 2L,
#'   overwrite    = TRUE
#' )
#'
#' # Single image
#' save_image_xlsx(
#'   image     = img1,
#'   file      = "single_image.xlsx",
#'   overwrite = TRUE
#' )
#' }
#'
#' @seealso [save_image_wb()], [openxlsx::insertImage()]
#'
#' @export
save_image_xlsx <- function(image, file, sheet = "Images",
                            rc = c(1L, 1L),
                            width = 12, height = 6, dpi = 300,
                            image_titles = NULL, title_size = 14,
                            font_name = getOption("instead.font"),
                            row_spacer = 2L, rows_per_inch = 5L,
                            overwrite = FALSE) {
  wb <- if (file.exists(file)) {
    openxlsx::loadWorkbook(file)
  } else {
    openxlsx::createWorkbook()
  }

  wb <- save_image_wb(
    image        = image,
    wb           = wb,
    sheet        = sheet,
    rc           = rc,
    width        = width,
    height       = height,
    dpi          = dpi,
    image_titles = image_titles,
    title_size   = title_size,
    font_name    = font_name,
    row_spacer   = row_spacer,
    rows_per_inch = rows_per_inch
  )

  openxlsx::saveWorkbook(wb, file = file, overwrite = overwrite)
  invisible(file)
}

#' Save image file(s) to Excel (one sheet per image)
#'
#' Writes each image to its own worksheet. A common start cell (`rc`) is used
#' for all sheets. Optional `sheet_names` and per-sheet `image_titles` supported.
#'
#' @param image Character vector (or list) of file paths.
#' @param file Path to target .xlsx.
#' @param rc Length-2 integer vector c(row, col) for top-left position. Default c(1L, 1L).
#' @param sheet_names Optional character vector of sheet names; length must equal length(image).
#'                    If NULL, names(image) (when present) are used; otherwise "Sheet i".
#' @param image_titles Optional character vector of titles (one per image); written one row above the image.
#' @param title_size Numeric font size for titles. Default 14.
#' @param font_name Font family for title text. Default getOption("instead.font").
#' @param width,height,dpi Passed to openxlsx::insertImage().
#' @param overwrite Overwrite workbook if TRUE.
#'
#' @return (Invisibly) the path to `file`. Side-effects: creates/updates an Excel file.
#'
#' @examples
#' \dontrun{
#' # Example images (replace with real PNG/JPG paths on your system)
#' img1 <- "plot1.png"
#' img2 <- "plot2.jpg"
#'
#' # 1) Named list/vector of images: names are used as sheet names
#' save_image_xlsx_split(
#'   image        = c("First plot" = img1, "Second plot" = img2),
#'   file         = "images_split.xlsx",
#'   rc           = c(1L, 1L),
#'   image_titles = c("Title for first image", "Title for second image"),
#'   title_size   = 16,
#'   overwrite    = TRUE
#' )
#'
#' # 2) Unnamed vector: sheets are auto-named ("Sheet 1", "Sheet 2", ...)
#' save_image_xlsx_split(
#'   image        = c(img1, img2),
#'   file         = "images_split_auto.xlsx",
#'   rc           = c(2L, 2L),   # start at B2
#'   overwrite    = TRUE
#' )
#' }
#'
#' @seealso [save_image_wb_split()], [openxlsx::insertImage()]
#'
#' @export
save_image_xlsx_split <- function(image, file,
                                  rc = c(1L, 1L),
                                  sheet_names = NULL,
                                  image_titles = NULL, title_size = 14,
                                  font_name = getOption("instead.font"),
                                  width = 12, height = 6, dpi = 300,
                                  overwrite = FALSE) {
  wb <- if (file.exists(file)) {
    openxlsx::loadWorkbook(file)
  } else {
    openxlsx::createWorkbook()
  }

  wb <- save_image_wb_split(
    image        = image,
    wb           = wb,
    rc           = rc,
    sheet_names  = sheet_names,
    image_titles = image_titles,
    title_size   = title_size,
    font_name    = font_name,
    width        = width,
    height       = height,
    dpi          = dpi
  )

  openxlsx::saveWorkbook(wb, file = file, overwrite = overwrite)
  invisible(file)
}

# Helper functions --------------------------------------------------------

#' Write a value into a single cell of an Excel worksheet
#'
#' A thin wrapper around [openxlsx::writeData()] with styling support.
#' This function writes a scalar value (character, numeric, logical, etc.)
#' into a specific `(row, col)` position of a worksheet and applies optional
#' font/border/align styles.
#'
#' @param wb A `Workbook` object created by [openxlsx::createWorkbook()] or
#'   loaded with [openxlsx::loadWorkbook()].
#' @param sheet Sheet name or index to write into.
#' @param rc A length-2 integer vector `c(row, col)` giving the target cell.
#' @param value A scalar value to write (character, numeric, logical, etc.).
#' @param bold,italic,underline,strikeout Logical; text decorations to apply.
#' @param font_name Character; font family. Default: `getOption("instead.font")`.
#' @param font_size Numeric; font size in points. Default `11`.
#' @param font_color Character; text color (hex string or named color). Default `NULL`.
#' @param fg_fill Character; background fill color. Default `NULL`.
#' @param h_align,v_align Character; horizontal and vertical alignment
#'   (`"left"`, `"center"`, `"right"` / `"top"`, `"center"`, `"bottom"`).
#'   Passed to [openxlsx::createStyle()].
#' @param ... Additional arguments forwarded to [openxlsx::writeData()].
#'
#' @return The modified `Workbook` (invisibly), allowing for chaining.
#'
#' @examples
#' \dontrun{
#' library(openxlsx)
#'
#' wb <- openxlsx::createWorkbook()
#' addWorksheet(wb, "Sheet1")
#'
#' # Write "Hello" into cell B2 with bold styling
#' write_cell(
#'   wb, sheet = "Sheet1",
#'   rc    = c(2L, 2L),
#'   value = "Hello",
#'   bold  = TRUE
#' )
#'
#' saveWorkbook(wb, "example.xlsx", overwrite = TRUE)
#' }
#'
#' @seealso [openxlsx::writeData()], [openxlsx::createStyle()]
#'
#' @export
write_cell <- function(wb, sheet, rc = c(1L, 1L), value,
                       bold = FALSE, italic = FALSE,
                       underline = FALSE, strikeout = FALSE,
                       font_name = getOption("instead.font"),
                       font_size = 11, font_color = NULL, fg_fill = NULL,
                       h_align = NULL, v_align = NULL, ...) {
  assert_class(wb, "Workbook")
  stopifnot(length(rc) == 2L)

  row <- rc[1L]
  col <- rc[2L]

  # write value
  openxlsx::writeData(
    wb      = wb,
    sheet   = sheet,
    x       = value,
    startRow = row,
    startCol = col,
    colNames = FALSE,
    rowNames = FALSE,
    ...
  )

  # build textDecoration vector
  text_decoration <- c(
    if (bold) "bold",
    if (italic) "italic",
    if (underline) "underline",
    if (strikeout) "strikeout"
  )
  if (!length(text_decoration))
    text_decoration <- NULL

  # create and apply style
  style <- openxlsx::createStyle(
    fontName       = font_name,
    fontSize       = font_size,
    fontColour     = font_color,
    textDecoration = text_decoration,
    fgFill         = fg_fill,
    halign         = h_align,
    valign         = v_align
  )

  openxlsx::addStyle(
    wb, sheet, style,
    rows = row, cols = col,
    gridExpand = FALSE, stack = TRUE
  )

  invisible(wb)
}

#' Write a styled data.frame to an Excel worksheet
#'
#' Low-level helper for writing tabular data into an Excel sheet with
#' consistent styling. Intended for use inside higher-level wrappers
#' (e.g., [save_data_xlsx()], [save_data_xlsx_split()]).
#'
#' Features:
#' - Writes data at the specified start cell (`rc`).
#' - Optionally includes row names.
#' - Applies different styles for header, body, and footer cells:
#'   * **Header**: grey fill, bold borders (thick/double at edges).
#'   * **Body**: thin borders (left/middle vs. right edge distinguished).
#'   * **Footer**: thicker bottom border.
#' - Sets column widths (fixed or vectorized).
#'
#' @param wb A Workbook object created by [openxlsx::createWorkbook()] or
#'   [openxlsx::loadWorkbook()].
#' @param sheet A single sheet name (character scalar).
#' @param data A `data.frame` (or `matrix`) to write.
#' @param rc A length-2 integer vector giving the starting row and column
#'   (`c(row, col)`). Default `c(1L, 1L)`.
#' @param row_names Logical; whether to include row names. Default `TRUE`.
#' @param font_name Character font family for all text. Default
#'   `getOption("instead.font")`.
#' @param font_size Numeric font size for all text. Default `11`.
#' @param fg_fill Background fill colour (hex string) used for header cells,
#'   and optionally for row-name cells when `row_names = TRUE`.
#'   Default `"#E6E6E7"`.
#' @param border_color Border colour (hex string). Default `"#000000"`.
#' @param widths Column width(s) to apply. Can be a single numeric (recycled)
#'   or a numeric vector matching the number of columns. Default `8.43`.
#' @param num_fmt Optional numeric format(s) applied to data cells (headers
#'   excluded). Accepts either:
#'   - A **single character format string** applied to all numeric cells
#'     (e.g. `"0.00"` for 2 decimals, `"#,##0"` for thousands separator,
#'     `"0.00%"` for percentages).
#'   - A **named character vector or list** mapping column names to formats,
#'     e.g. `c(value = "#,##0", rate = "0.0%")`. Columns not listed are left
#'     unchanged. Unknown column names are silently ignored.
#'
#' @return No return value. Called for side effects (writes into workbook).
#'
#' @details
#' Formatting strings follow Excel's built-in custom number formats. Some
#' common patterns:
#' \itemize{
#'   \item `"0"` -> integers
#'   \item `"0.00"` -> fixed 2 decimals
#'   \item `"#,##0"` -> thousands separator
#'   \item `"0.00%"` -> percentages with 2 decimals
#' }
#' @examples
#' \dontrun{
#' wb <- openxlsx::createWorkbook()
#' openxlsx::addWorksheet(wb, "Example")
#' write_data(
#'   wb, "Example",
#'   data = head(mtcars),
#'   rc = c(2, 2),
#'   row_names = FALSE,
#'   font_name = "Calibri",
#'   font_size = 11,
#'   border_color = "#333333",
#'   widths = 10
#' )
#' openxlsx::saveWorkbook(wb, "styled_data.xlsx", overwrite = TRUE)
#' }
#'
#' @export
write_data <- function(wb, sheet, data, rc = c(1L, 1L), row_names = TRUE,
                       font_name = getOption("instead.font"), font_size = 11,
                       fg_fill = "#E6E6E7", border_color = "#000000", widths = 8.43,
                       num_fmt = NULL) {
  header_style1 <- openxlsx::createStyle(
    fontName = font_name,
    fontSize = font_size,
    fontColour = "#000000",
    halign = "center",
    valign = "center",
    fgFill = fg_fill,
    border = "TopRightBottom",
    borderColour = border_color,
    borderStyle = c("medium", "thin", "double")
  )
  header_style2 <- openxlsx::createStyle(
    fontName = font_name,
    fontSize = font_size,
    fontColour = "#000000",
    halign = "center",
    valign = "center",
    fgFill = fg_fill,
    border = "TopBottom",
    borderColour = border_color,
    borderStyle = c("medium", "double")
  )
  body_style1 <- openxlsx::createStyle(
    fontName = font_name,
    fontSize = font_size,
    border = "TopRightBottom",
    borderColour = border_color
  )
  body_style2 <- openxlsx::createStyle(
    fontName = font_name,
    fontSize = font_size,
    border = "TopBottom",
    borderColour = border_color
  )
  footer_style1 <- openxlsx::createStyle(
    fontName = font_name,
    fontSize = font_size,
    border = "TopRightBottom",
    borderColour = border_color,
    borderStyle = c("thin", "thin", "medium")
  )
  footer_style2 <- openxlsx::createStyle(
    fontName = font_name,
    fontSize = font_size,
    border = "TopBottom",
    borderColour = border_color,
    borderStyle = c("thin", "medium")
  )

  openxlsx::writeData(
    wb = wb, sheet = sheet, x = data,
    xy = rev(rc), rowNames = row_names
  )

  start_cell <- rc
  end_cell   <- start_cell + dim(data)

  srow <- start_cell[1L]
  scol <- start_cell[2L]
  erow <- end_cell[1L]
  ecol <- end_cell[2L]

  if (!row_names) ecol <- ecol - 1

  # Columns for styling
  header_cols <- scol:ecol
  left_cols   <- scol:max(ecol - 1L, scol)  # left block (all but last col)
  right_col   <- ecol                       # rightmost column

  # Key: Body rows are empty if only 1 row of data
  n <- nrow(data)
  body_rows <- if (n >= 2L) seq.int(srow + 1L, erow - 1L) else integer(0)

  # Header / Footer rows
  header_row <- srow
  footer_row <- erow

  # Header styles
  openxlsx::addStyle(wb, sheet, header_style1, rows = header_row,
                     cols = left_cols, gridExpand = TRUE)
  openxlsx::addStyle(wb, sheet, header_style2, rows = header_row,
                     cols = right_col, gridExpand = TRUE)

  # Body styles (only applied if n >= 2; skipped if body_rows = integer(0))
  openxlsx::addStyle(wb, sheet, body_style1, rows = body_rows,
                     cols = left_cols, gridExpand = TRUE)
  openxlsx::addStyle(wb, sheet, body_style2, rows = body_rows,
                     cols = right_col, gridExpand = TRUE)

  # Footer styles
  # Special case: if only 1 row, remove top border to avoid double-line clash
  footer_style1_top_none <- openxlsx::createStyle(
    fontName     = font_name,
    border       = "TopRightBottom",
    borderColour = border_color,
    borderStyle  = c("none", "thin", "medium")  # no top border
  )
  footer_style2_top_none <- openxlsx::createStyle(
    fontName     = font_name,
    border       = "TopBottom",
    borderColour = border_color,
    borderStyle  = c("none", "medium")          # no top border
  )

  if (n == 1L) {
    # Single-row table: apply "no-top-border" footer styles
    openxlsx::addStyle(wb, sheet, footer_style1_top_none, rows = footer_row,
                       cols = left_cols, gridExpand = TRUE)
    openxlsx::addStyle(wb, sheet, footer_style2_top_none, rows = footer_row,
                       cols = right_col, gridExpand = TRUE)
  } else {
    # Multi-row table: apply normal footer styles
    openxlsx::addStyle(wb, sheet, footer_style1, rows = footer_row,
                       cols = left_cols, gridExpand = TRUE)
    openxlsx::addStyle(wb, sheet, footer_style2, rows = footer_row,
                       cols = right_col, gridExpand = TRUE)
  }

  # Row-name column fill (only data rows & footer, never header)
  if (row_names) {
    data_rows <- c(body_rows, footer_row)
    if (length(data_rows)) {
      rn_fill <- openxlsx::createStyle(fgFill = fg_fill)
      openxlsx::addStyle(
        wb, sheet, rn_fill,
        rows = data_rows, cols = scol,  # scol is the rownames column on sheet
        gridExpand = TRUE, stack = TRUE
      )
    }
  }

  # Applies only to data rows (body + last row), never to header.
  if (!is.null(num_fmt)) {
    fmts <- if (is.list(num_fmt)) unlist(num_fmt) else num_fmt

    # rows to format = all body rows + footer row (last data row)
    data_rows <- c(body_rows, footer_row)

    if (length(data_rows)) {
      # 1) single unnamed string -> apply to all *numeric* columns
      if (is.character(fmts) && length(fmts) == 1L && is.null(names(fmts))) {
        target_cols <- which(vapply(data, is.numeric, logical(1)))
        if (length(target_cols)) {
          fmt_style <- openxlsx::createStyle(numFmt = fmts)
          for (j in target_cols) {
            col_on_sheet <- if (row_names) scol + j else scol + j - 1L
            openxlsx::addStyle(
              wb, sheet, fmt_style,
              rows = data_rows, cols = col_on_sheet,
              gridExpand = TRUE, stack = TRUE
            )
          }
        }

      # 2) named vector/list -> apply only to matching columns
      } else {
        # keep only valid, existing column names
        if (is.null(names(fmts))) {
          warning("Ignoring `num_fmt` because it is an unnamed multi-value vector.")
        } else {
          fmts <- fmts[intersect(names(fmts), colnames(data))]
          if (length(fmts)) {
            for (nm in names(fmts)) {
              j <- match(nm, colnames(data))
              # format numeric columns only
              if (is.na(j) || !is.numeric(data[[j]])) next
              col_on_sheet <- if (row_names) scol + j else scol + j - 1L
              fmt_style <- openxlsx::createStyle(numFmt = unname(fmts[[nm]]))
              openxlsx::addStyle(
                wb, sheet, fmt_style,
                rows = data_rows, cols = col_on_sheet,
                gridExpand = TRUE, stack = TRUE
              )
            }
          }
        }
      }
    }
  }
  # openxlsx::setColWidths(wb, sheet, cols = header_cols, widths = widths)
}

# Internal helper functions -----------------------------------------------

#' @keywords internal
#' @noRd
.get_col_widths <- function(df, row_names = FALSE, min_width = 8.43, pad = 3) {
  stopifnot(is.data.frame(df))
  if (row_names) {
    df <- cbind(.row = rownames(df), df, stringsAsFactors = FALSE)
  }
  cols <- seq_along(df)
  vapply(cols, function(j) {
    header <- colnames(df)[j]
    body   <- as.character(df[[j]])
    w <- max(nchar(c(header, body), type = "width"), na.rm = TRUE)
    max(min_width, w + pad)
  }, numeric(1L))
}

#' @keywords internal
#' @noRd
.insert_plot <- function(wb, sheet, plot, rc = c(1L, 1L), width = 12, height = 6,
                         dpi = 300) {
  print(plot)
  openxlsx::insertPlot(wb, sheet = sheet, width = width, height = height,
                       startRow = rc[1L], startCol = rc[2L], dpi = dpi)
}

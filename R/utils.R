#' Internal helpers
#'
#' @keywords internal
#' @noRd
require_pkg <- function(pkg) {
  if (!requireNamespace(pkg, quietly = TRUE)) {
    stop(sprintf("Package '%s' is required but not installed.", pkg), call. = FALSE)
  }
}

#' Case-insensitive lookup of a worksheet name. Returns NULL when absent.
#'
#' @keywords internal
#' @noRd
match_sheet_name <- function(sheets, target) {
  idx <- which(tolower(sheets) == tolower(target))
  if (length(idx) == 0) {
    return(NULL)
  }
  sheets[idx[1]]
}

read_xlsx_sheet <- function(path, sheet, skip_empty = TRUE) {
  require_pkg("openxlsx")
  openxlsx::read.xlsx(
    path,
    sheet = sheet,
    colNames = TRUE,
    na.strings = c("", "NA"),
    skipEmptyRows = skip_empty
  )
}

#' Read a worksheet of an XLSForm
#'
#' When a column mixes text and numbers (a choices `value` column holding both
#' `yes` and `1`), openxlsx reads it as character and keeps each number's raw
#' cell text. Workbooks saved by Python tooling store integers as `1.0`, which
#' SurveyCTO reads as `1`, so integer-valued text of that form is normalized.
#' @keywords internal
#' @noRd
read_xlsform_sheet <- function(path, sheet, skip_empty = TRUE) {
  x <- read_xlsx_sheet(path, sheet, skip_empty = skip_empty)
  x[] <- lapply(x, function(col) {
    if (is.character(col)) sub("^(-?[0-9]+)\\.0+$", "\\1", col) else col
  })
  x
}

write_xlsx_sheets <- function(path, sheets) {
  require_pkg("openxlsx")
  wb <- openxlsx::createWorkbook()
  for (name in names(sheets)) {
    openxlsx::addWorksheet(wb, name)
    openxlsx::writeData(wb, name, sheets[[name]])
  }
  openxlsx::saveWorkbook(wb, path, overwrite = TRUE)
}

sanitize_text <- function(x) {
  if (is.null(x)) return(x)
  # Collapse line breaks so a label fits one Excel cell. Stripping $, ` and "
  # was a Stata escaping concern and has no meaning here.
  x <- gsub("\n", " ", x, fixed = TRUE)
  x <- gsub("\r", " ", x, fixed = TRUE)
  x
}

is_blank <- function(x) {
  is.na(x) | trimws(x) == ""
}

make_labelled <- function(x, labels) {
  if (requireNamespace("haven", quietly = TRUE)) {
    # haven::labelled() rebuilds the vector, so carry the variable label over.
    return(haven::labelled(
      x,
      labels = labels,
      label = attr(x, "label", exact = TRUE)
    ))
  }
  x
}

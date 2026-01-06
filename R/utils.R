#' Internal helpers
#'
#' @keywords internal
#' @noRd
require_pkg <- function(pkg) {
  if (!requireNamespace(pkg, quietly = TRUE)) {
    stop(sprintf("Package '%s' is required but not installed.", pkg), call. = FALSE)
  }
}

read_xlsx_sheet <- function(path, sheet) {
  require_pkg("openxlsx")
  openxlsx::read.xlsx(path, sheet = sheet, colNames = TRUE, na.strings = c("", "NA"))
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
  x <- gsub('"', "", x, fixed = TRUE)
  x <- gsub("`", "", x, fixed = TRUE)
  x <- gsub("$", "", x, fixed = TRUE)
  x <- gsub("\n", " ", x, fixed = TRUE)
  x <- gsub("\r", " ", x, fixed = TRUE)
  x
}

is_blank <- function(x) {
  is.na(x) | trimws(x) == ""
}

make_labelled <- function(x, labels) {
  if (requireNamespace("haven", quietly = TRUE)) {
    return(haven::labelled(x, labels = labels))
  }
  x
}

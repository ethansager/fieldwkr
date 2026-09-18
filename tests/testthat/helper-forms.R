# Build a small XLSForm on disk for tests. `survey` and `choices` are data
# frames; missing columns are filled with blanks.
make_form <- function(survey, choices = NULL, settings = NULL) {
  path <- tempfile(fileext = ".xlsx")
  wb <- openxlsx::createWorkbook()
  openxlsx::addWorksheet(wb, "survey")
  openxlsx::writeData(wb, "survey", survey)
  if (is.null(choices)) {
    choices <- data.frame(
      list_name = character(),
      value = character(),
      label = character()
    )
  }
  openxlsx::addWorksheet(wb, "choices")
  openxlsx::writeData(wb, "choices", choices)
  if (!is.null(settings)) {
    openxlsx::addWorksheet(wb, "settings")
    openxlsx::writeData(wb, "settings", settings)
  }
  openxlsx::saveWorkbook(wb, path, overwrite = TRUE)
  path
}

# Survey sheet from equal-length vectors, with blank defaults for the
# expression columns.
survey_rows <- function(type, name, label = name, ...) {
  extra <- list(...)
  n <- length(type)
  out <- data.frame(type = type, name = name, label = label, stringsAsFactors = FALSE)
  for (col in c("relevance", "constraint", "calculation", "repeat_count", "choice_filter")) {
    out[[col]] <- if (!is.null(extra[[col]])) extra[[col]] else rep("", n)
  }
  out
}

# Instance columns of a repeated field, in instance order.
instance_cols <- function(df, field) {
  grep(paste0("^", field, "__r[0-9]+$"), names(df), value = TRUE)
}

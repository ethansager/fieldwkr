#' Check a SurveyCTO/ODK XLSForm
#'
#' @param path Path to the XLSForm (Excel file).
#' @param verbose Print summary output.
#' @return A list with errors and warnings.
#' @export
test_form <- function(path, verbose = TRUE) {
  survey <- read_xlsx_sheet(path, "survey")
  choices <- read_xlsx_sheet(path, "choices")

  errors <- character()
  warnings <- character()

  required_cols <- c("type", "name")
  if (!all(required_cols %in% names(survey))) {
    missing_cols <- setdiff(required_cols, names(survey))
    errors <- c(
      errors,
      sprintf(
        "Missing required columns in survey sheet: %s",
        paste(missing_cols, collapse = ", ")
      )
    )
    return(list(errors = errors, warnings = warnings))
  }

  if (any(duplicated(survey$name))) {
    dups <- unique(survey$name[duplicated(survey$name)])
    errors <- c(
      errors,
      sprintf(
        "Duplicate names in survey sheet: %s",
        paste(dups, collapse = ", ")
      )
    )
  }

  if (!"label" %in% names(survey)) {
    warnings <- c(warnings, "Column 'label' not found in survey sheet.")
  } else {
    no_label_types <- c(
      "begin_group",
      "end_group",
      "begin_repeat",
      "end_repeat",
      "calculate"
    )
    missing_label <- is_blank(survey$label) & !survey$type %in% no_label_types
    if (any(missing_label)) {
      rows <- which(missing_label)
      warnings <- c(
        warnings,
        sprintf(
          "Missing labels in survey sheet rows: %s",
          paste(rows, collapse = ", ")
        )
      )
    }
  }

  if (nrow(choices) > 0) {
    required_choice_cols <- c("list_name", "value", "label")
    if (!all(required_choice_cols %in% names(choices))) {
      missing_cols <- setdiff(required_choice_cols, names(choices))
      errors <- c(
        errors,
        sprintf(
          "Missing required columns in choices sheet: %s",
          paste(missing_cols, collapse = ", ")
        )
      )
    } else {
      dup_choices <- duplicated(choices[, c("list_name", "value")])
      if (any(dup_choices)) {
        errors <- c(
          errors,
          "Duplicate list_name/value pairs found in choices sheet."
        )
      }
    }
  }

  select_rows <- grepl("select_one|select_multiple", survey$type)
  if (any(select_rows) && "list_name" %in% names(choices)) {
    lists <- unique(choices$list_name)
    for (i in which(select_rows)) {
      parts <- strsplit(survey$type[i], " ")[[1]]
      if (length(parts) >= 2) {
        list_name <- parts[2]
        if (!list_name %in% lists) {
          warnings <- c(
            warnings,
            sprintf(
              "Missing choices list '%s' referenced in survey row %s.",
              list_name,
              i
            )
          )
        }
      }
    }
  }

  if (verbose) {
    message(sprintf("Errors: %s", length(errors)))
    if (length(errors) > 0) {
      message(paste(errors, collapse = "\n"))
    }
    message(sprintf("Warnings: %s", length(warnings)))
    if (length(warnings) > 0) message(paste(warnings, collapse = "\n"))
  }

  list(errors = errors, warnings = warnings)
}

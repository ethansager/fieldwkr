#' Check a SurveyCTO/ODK XLSForm
#'
#' @param path Path to the XLSForm (Excel file).
#' @param verbose Print summary output.
#' @details
#' Validates high-impact programming issues in the `survey` and `choices`
#' sheets, including:
#' - required columns and duplicate question names
#' - malformed `select_one` / `select_multiple` types
#' - missing or duplicate choices
#' - begin/end group and repeat balance
#' - unresolved `${var}` references in expression columns
#'
#' Output is a list with `errors` and `warnings` vectors suitable for CI checks.
#' @return A list with errors and warnings.
#' @export
test_form <- function(path, verbose = TRUE) {
  errors <- character()
  warnings <- character()

  add_error <- function(msg) {
    errors <<- c(errors, msg)
  }
  add_warning <- function(msg) {
    warnings <<- c(warnings, msg)
  }

  if (!file.exists(path)) {
    stop(sprintf("File not found: %s", path), call. = FALSE)
  }

  require_pkg("openxlsx")
  sheets <- tryCatch(
    openxlsx::getSheetNames(path),
    error = function(e) {
      add_error(sprintf("Unable to read workbook: %s", e$message))
      character()
    }
  )

  get_sheet_name <- function(target) {
    idx <- which(tolower(sheets) == tolower(target))
    if (length(idx) == 0) {
      return(NULL)
    }
    sheets[idx[1]]
  }

  survey_sheet <- get_sheet_name("survey")
  choices_sheet <- get_sheet_name("choices")

  if (is.null(survey_sheet)) {
    add_error("Missing required sheet 'survey'.")
    return(list(errors = errors, warnings = warnings))
  }

  if (is.null(choices_sheet)) {
    add_warning("Sheet 'choices' not found.")
  }

  survey <- tryCatch(
    read_xlsx_sheet(path, survey_sheet),
    error = function(e) {
      add_error(sprintf("Failed to read survey sheet: %s", e$message))
      data.frame(stringsAsFactors = FALSE)
    }
  )
  choices <- if (is.null(choices_sheet)) {
    data.frame(stringsAsFactors = FALSE)
  } else {
    tryCatch(
      read_xlsx_sheet(path, choices_sheet),
      error = function(e) {
        add_error(sprintf("Failed to read choices sheet: %s", e$message))
        data.frame(stringsAsFactors = FALSE)
      }
    )
  }

  if (length(errors) > 0 && nrow(survey) == 0) {
    return(list(errors = errors, warnings = warnings))
  }

  required_cols <- c("type", "name")
  if (!all(required_cols %in% names(survey))) {
    missing_cols <- setdiff(required_cols, names(survey))
    add_error(sprintf(
      "Missing required columns in survey sheet: %s",
      paste(missing_cols, collapse = ", ")
    ))
    return(list(errors = errors, warnings = warnings))
  }

  as_clean_chr <- function(x) {
    out <- as.character(x)
    out[is.na(out)] <- ""
    trimws(out)
  }

  type_raw <- as_clean_chr(survey$type)
  type_norm <- normalize_xlsform_type(type_raw)
  name_raw <- as_clean_chr(survey$name)

  missing_name <- name_raw == "" &
    type_norm != "" &
    !type_norm %in% c("end group", "end repeat")
  if (any(missing_name)) {
    add_error(sprintf(
      "Missing names in survey sheet rows: %s",
      paste(which(missing_name), collapse = ", ")
    ))
  }

  non_blank_name <- name_raw != ""
  if (any(duplicated(name_raw[non_blank_name]))) {
    dups <- unique(name_raw[non_blank_name][duplicated(name_raw[non_blank_name])])
    add_error(sprintf(
      "Duplicate names in survey sheet: %s",
      paste(dups, collapse = ", ")
    ))
  }

  invalid_name <- non_blank_name &
    !grepl("^[A-Za-z][A-Za-z0-9_.-]*$", name_raw)
  if (any(invalid_name)) {
    add_warning(sprintf(
      "Potentially invalid XLSForm names in survey rows: %s",
      paste(which(invalid_name), collapse = ", ")
    ))
  }

  malformed_select <- type_norm %in% c("select_one", "select_multiple")
  if (any(malformed_select)) {
    add_error(sprintf(
      "Malformed select question types (missing list names) in survey rows: %s",
      paste(which(malformed_select), collapse = ", ")
    ))
  }

  known_simple_types <- c(
    "text",
    "integer",
    "decimal",
    "date",
    "time",
    "datetime",
    "calculate",
    "note",
    "acknowledge",
    "rank",
    "hidden",
    "start",
    "end",
    "today",
    "deviceid",
    "subscriberid",
    "simserial",
    "phonenumber",
    "username",
    "email",
    "geopoint",
    "geotrace",
    "geoshape",
    "image",
    "audio",
    "video",
    "file",
    "barcode"
  )
  known_structural <- c("begin group", "end group", "begin repeat", "end repeat")
  known_select_pattern <- "^select_(one|multiple)\\s+\\S+$"
  unknown_type <- type_norm != "" &
    !type_norm %in% known_simple_types &
    !type_norm %in% known_structural &
    !grepl(known_select_pattern, type_norm)
  if (any(unknown_type)) {
    add_warning(sprintf(
      "Unknown or unsupported survey question types in rows: %s",
      paste(which(unknown_type), collapse = ", ")
    ))
  }

  if ("label" %in% names(survey)) {
    label <- as_clean_chr(survey$label)
    no_label_types <- c(
      "begin group",
      "end group",
      "begin repeat",
      "end repeat",
      "calculate",
      "start",
      "end",
      "today",
      "deviceid",
      "subscriberid",
      "simserial",
      "phonenumber",
      "username"
    )
    missing_label <- label == "" &
      type_norm != "" &
      !type_norm %in% no_label_types
    if (any(missing_label)) {
      add_warning(sprintf(
        "Missing labels in survey sheet rows: %s",
        paste(which(missing_label), collapse = ", ")
      ))
    }
  } else {
    add_warning("Column 'label' not found in survey sheet.")
  }

  stack <- data.frame(
    type = character(),
    row = integer(),
    stringsAsFactors = FALSE
  )
  for (i in seq_along(type_norm)) {
    t <- type_norm[i]
    if (t %in% c("begin group", "begin repeat")) {
      stack <- rbind(
        stack,
        data.frame(type = t, row = i, stringsAsFactors = FALSE)
      )
      next
    }
    if (t %in% c("end group", "end repeat")) {
      if (nrow(stack) == 0) {
        add_error(sprintf(
          "Unmatched '%s' at survey row %s.",
          t,
          i
        ))
        next
      }

      top <- stack[nrow(stack), , drop = FALSE]
      expected <- if (top$type == "begin group") "end group" else "end repeat"
      if (!identical(t, expected)) {
        add_error(sprintf(
          "Mismatched '%s' at survey row %s; expected '%s' for opener at row %s.",
          t,
          i,
          expected,
          top$row
        ))
        next
      }

      stack <- stack[-nrow(stack), , drop = FALSE]
    }
  }
  if (nrow(stack) > 0) {
    for (j in seq_len(nrow(stack))) {
      add_error(sprintf(
        "Unclosed '%s' opened at survey row %s.",
        stack$type[j],
        stack$row[j]
      ))
    }
  }

  expr_cols <- intersect(
    c("relevance", "constraint", "calculation", "required", "choice_filter", "repeat_count"),
    names(survey)
  )
  defined_names <- unique(name_raw[name_raw != ""])
  for (col in expr_cols) {
    vals <- as_clean_chr(survey[[col]])
    non_blank <- which(vals != "")
    for (i in non_blank) {
      refs <- xls_extract_refs(vals[i])
      if (length(refs) == 0) {
        next
      }
      unknown_refs <- setdiff(refs, defined_names)
      if (length(unknown_refs) > 0) {
        add_error(sprintf(
          "Unknown ${} reference(s) in survey row %s column '%s': %s",
          i,
          col,
          paste(unknown_refs, collapse = ", ")
        ))
      }
    }
  }

  select_rows <- grepl("^select_(one|multiple)\\b", type_norm)
  referenced_lists <- character()
  if (any(select_rows)) {
    for (i in which(select_rows)) {
      parts <- strsplit(type_norm[i], " +")[[1]]
      if (length(parts) >= 2) {
        referenced_lists <- c(referenced_lists, parts[2])
      }
    }
  }
  referenced_lists <- unique(referenced_lists)

  if (nrow(choices) > 0 || length(referenced_lists) > 0) {
    choice_required <- c("list_name", "label")
    if (!all(choice_required %in% names(choices))) {
      missing_cols <- setdiff(choice_required, names(choices))
      add_error(sprintf(
        "Missing required columns in choices sheet: %s",
        paste(missing_cols, collapse = ", ")
      ))
    } else {
      value_col <- if ("name" %in% names(choices)) {
        "name"
      } else if ("value" %in% names(choices)) {
        "value"
      } else {
        NULL
      }

      if (is.null(value_col)) {
        add_error("Choices sheet must include a 'name' or 'value' column.")
      } else {
        list_name <- as_clean_chr(choices$list_name)
        value <- as_clean_chr(choices[[value_col]])

        missing_list_name <- which(list_name == "")
        if (length(missing_list_name) > 0) {
          add_warning(sprintf(
            "Blank list_name values found in choices rows: %s",
            paste(missing_list_name, collapse = ", ")
          ))
        }

        missing_value <- which(value == "")
        if (length(missing_value) > 0) {
          add_warning(sprintf(
            "Blank %s values found in choices rows: %s",
            value_col,
            paste(missing_value, collapse = ", ")
          ))
        }

        key <- paste(list_name, value, sep = "\r")
        dup_key <- duplicated(key) & list_name != "" & value != ""
        if (any(dup_key)) {
          add_error(sprintf(
            "Duplicate list_name/%s pairs found in choices rows: %s",
            value_col,
            paste(which(dup_key), collapse = ", ")
          ))
        }

        available_lists <- unique(list_name[list_name != ""])
        missing_lists <- setdiff(referenced_lists, available_lists)
        if (length(missing_lists) > 0) {
          add_error(sprintf(
            "Missing choices list(s) referenced in survey: %s",
            paste(missing_lists, collapse = ", ")
          ))
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

#' @keywords internal
#' @noRd
normalize_xlsform_type <- function(x) {
  x <- tolower(trimws(as.character(x)))
  x[is.na(x)] <- ""
  gsub("\\s+", " ", x)
}

#' @keywords internal
#' @noRd
xls_extract_refs <- function(expr) {
  matches <- gregexpr("\\$\\{[^}]+\\}", expr, perl = TRUE)[[1]]
  if (identical(matches, -1L)) {
    return(character())
  }
  refs <- regmatches(expr, list(matches))[[1]]
  refs <- gsub("^\\$\\{|\\}$", "", refs)
  unique(trimws(refs[nzchar(refs)]))
}

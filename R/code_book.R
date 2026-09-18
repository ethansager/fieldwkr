#' Export a codebook to Excel
#'
#' @param data Data frame to document.
#' @param path Output path for the Excel codebook.
#' @param survey Survey name to tag in the codebook.
#' @details
#' Exports two sheets:
#' - `survey`: variable metadata (`name`, `label`, `type`, `choices`) plus
#'   survey-specific columns `name_<survey>` and `recode_<survey>`.
#' - `choices`: generated choice lists for factors and labelled vectors.
#'
#' This structure is designed for survey operations where one master codebook
#' maps variable names across rounds/instruments.
#' @return Invisibly returns the output path.
#' @export
cb_export <- function(data, path, survey = "current") {
  stopifnot(is.data.frame(data))

  vars <- names(data)
  labels <- vapply(
    vars,
    function(v) {
      # exact = TRUE: otherwise attr() partially matches "labels" (the value
      # labels) on a labelled column that has no variable label.
      label <- attr(data[[v]], "label", exact = TRUE)
      if (is.null(label)) "" else as.character(label)
    },
    character(1)
  )

  types <- vapply(
    vars,
    function(v) {
      x <- data[[v]]
      if (inherits(x, "labelled")) {
        if (is.numeric(x)) "numeric" else "string"
      } else if (is.factor(x)) {
        "string"
      } else if (is.character(x)) {
        "string"
      } else if (is.integer(x)) {
        "integer"
      } else if (is.numeric(x)) {
        "numeric"
      } else {
        class(x)[1]
      }
    },
    character(1)
  )

  choices <- character(length(vars))
  choices_sheet <- data.frame(
    list_name = character(),
    value = character(),
    label = character(),
    stringsAsFactors = FALSE
  )

  for (i in seq_along(vars)) {
    v <- vars[i]
    x <- data[[v]]
    list_name <- ""
    choice_map <- NULL

    if (is.factor(x)) {
      list_name <- paste0("choices_", v)
      choice_map <- setNames(seq_along(levels(x)), levels(x))
    } else if (!is.null(attr(x, "labels"))) {
      labels_attr <- attr(x, "labels")
      if (is.numeric(labels_attr)) {
        list_name <- paste0("choices_", v)
        choice_map <- labels_attr
      }
    }

    if (nzchar(list_name)) {
      choices[i] <- list_name
      if (!is.null(choice_map)) {
        new_rows <- data.frame(
          list_name = list_name,
          value = as.character(unname(choice_map)),
          label = names(choice_map),
          stringsAsFactors = FALSE
        )
        choices_sheet <- rbind(choices_sheet, new_rows)
      }
    }
  }

  survey_sheet <- data.frame(
    name = vars,
    label = sanitize_text(labels),
    type = types,
    choices = choices,
    stringsAsFactors = FALSE
  )

  name_col <- paste0("name_", survey)
  recode_col <- paste0("recode_", survey)
  survey_sheet[[name_col]] <- vars
  survey_sheet[[recode_col]] <- ""

  sheets <- list(
    survey = survey_sheet,
    choices = choices_sheet
  )

  write_xlsx_sheets(path, sheets)
  invisible(path)
}

#' Apply a codebook to a data frame
#'
#' @param data Data frame to update.
#' @param path Path to the Excel codebook.
#' @param survey Survey name to apply.
#' @param drop Drop variables not listed in the codebook.
#' @param missing_values Named vector of missing-value labels to add.
#' @param strict When `TRUE` (the default), stop if a variable named in the
#'   codebook is not in the data. When `FALSE`, skip it with a message. Useful
#'   with a codebook from [cb_from_form()], which lists every field in the form
#'   whether or not a given export contains it.
#' @details
#' `cb_apply()` performs, in order:
#' 1) variable-level recodes (`recode_<survey>`),
#' 2) label updates,
#' 3) value label / factor reconstruction from `choices`,
#' 4) optional drops and renames from `name_<survey>` -> `name`.
#'
#' For numeric choice values, `haven::labelled` is used when available.
#' @return Updated data frame.
#' @export
cb_apply <- function(
  data,
  path,
  survey = "current",
  drop = FALSE,
  missing_values = NULL,
  strict = TRUE
) {
  stopifnot(is.data.frame(data))

  survey_sheet <- read_xlsx_sheet(path, "survey")
  choices_sheet <- read_xlsx_sheet(path, "choices")

  name_col <- paste0("name_", survey)
  recode_col <- paste0("recode_", survey)

  if (!name_col %in% names(survey_sheet)) {
    stop(
      sprintf("Survey column '%s' not found in codebook.", name_col),
      call. = FALSE
    )
  }

  survey_sheet <- as.data.frame(
    lapply(survey_sheet, sanitize_text),
    stringsAsFactors = FALSE
  )
  choices_sheet <- as.data.frame(
    lapply(choices_sheet, sanitize_text),
    stringsAsFactors = FALSE
  )

  old_names <- survey_sheet[[name_col]]
  new_names <- survey_sheet$name
  labels <- survey_sheet$label
  choices <- survey_sheet$choices
  recodes <- if (recode_col %in% names(survey_sheet)) {
    survey_sheet[[recode_col]]
  } else {
    rep("", nrow(survey_sheet))
  }

  to_drop <- character()
  not_found <- character()
  rename_from <- character()
  rename_to <- character()

  for (i in seq_len(nrow(survey_sheet))) {
    old <- old_names[i]
    new <- new_names[i]

    if (is_blank(old)) {
      next
    }

    if (!old %in% names(data)) {
      if (strict) {
        stop(sprintf("Variable '%s' not found in data.", old), call. = FALSE)
      }
      not_found <- c(not_found, old)
      next
    }

    if ((drop && is_blank(new)) || identical(new, ".")) {
      to_drop <- c(to_drop, old)
      next
    }

    if (!is_blank(new) && !identical(new, old)) {
      rename_from <- c(rename_from, old)
      rename_to <- c(rename_to, new)
    }

    if (!is_blank(labels[i])) {
      attr(data[[old]], "label") <- labels[i]
    }

    if (!is_blank(recodes[i])) {
      data[[old]] <- apply_recode(data[[old]], recodes[i])
    }

    if (!is_blank(choices[i])) {
      list_name <- choices[i]
      rows <- choices_sheet[
        choices_sheet$list_name == list_name,
        ,
        drop = FALSE
      ]
      if (nrow(rows) == 0) {
        stop(
          sprintf("Choice list '%s' not found in choices sheet.", list_name),
          call. = FALSE
        )
      }

      if (!is.null(missing_values)) {
        mv <- data.frame(
          list_name = list_name,
          value = names(missing_values),
          label = unname(missing_values),
          stringsAsFactors = FALSE
        )
        rows <- rbind(rows, mv)
      }

      vals <- suppressWarnings(as.numeric(rows$value))
      labs <- rows$label
      if (all(!is.na(vals))) {
        x <- data[[old]]
        if (is.character(x)) {
          # SurveyCTO exports often read select_one answers as character while
          # the codebook holds numeric codes. Promote the column when every
          # value parses; otherwise label it with character codes. Neither
          # branch turns a value into NA.
          parsed <- suppressWarnings(as.numeric(x))
          if (all(is.na(x) | !is.na(parsed))) {
            attr(parsed, "label") <- attr(x, "label", exact = TRUE)
            x <- parsed
          } else {
            vals <- as.character(rows$value)
          }
        }
        labels_map <- setNames(vals, labs)
        fmt <- attr(data[[old]], "format.stata", exact = TRUE)
        same_type <- identical(typeof(x), typeof(data[[old]]))
        data[[old]] <- make_labelled(x, labels = labels_map)
        if (!is.null(fmt) && same_type) {
          attr(data[[old]], "format.stata") <- fmt
        }
      } else {
        # factor() turns any value outside `levels` into NA, so check first.
        x <- as.character(data[[old]])
        present <- !is.na(x) & nzchar(x)
        unmatched <- unique(x[present & !x %in% rows$value])
        if (length(unmatched) > 0) {
          stop(
            sprintf(
              paste0(
                "Variable '%s' has value(s) not in choice list '%s': %s. ",
                "Converting it would make them missing; fix the codebook's ",
                "choices or the data first."
              ),
              old,
              list_name,
              paste(utils::head(unmatched, 5), collapse = ", ")
            ),
            call. = FALSE
          )
        }
        data[[old]] <- factor(data[[old]], levels = rows$value, labels = labs)
      }
    }
  }

  if (length(not_found) > 0) {
    message(sprintf(
      "%d codebook variable(s) not in the data were skipped: %s%s",
      length(not_found),
      paste(utils::head(not_found, 10), collapse = ", "),
      if (length(not_found) > 10) ", ..." else ""
    ))
  }

  if (length(to_drop) == ncol(data)) {
    stop("Dropping all variables is not allowed.", call. = FALSE)
  }

  if (length(to_drop) > 0) {
    data <- data[setdiff(names(data), to_drop)]
  }

  if (length(rename_from) > 0) {
    if (any(duplicated(rename_to))) {
      stop("Rename conflict detected in codebook.", call. = FALSE)
    }
    clash <- intersect(rename_to, setdiff(names(data), rename_from))
    if (length(clash) > 0) {
      stop(
        sprintf(
          "Rename target(s) already present in the data: %s",
          paste(clash, collapse = ", ")
        ),
        call. = FALSE
      )
    }
    names(data)[match(rename_from, names(data))] <- rename_to
  }

  data
}

#' Create a codebook template
#'
#' @param data Data frame to template.
#' @param path Output path for the Excel codebook.
#' @param survey Survey name to tag in the codebook.
#' @details
#' Convenience wrapper around [cb_export()] for initializing a codebook from a
#' single data frame.
#' @return Invisibly returns the output path.
#' @export
cb_template <- function(data, path, survey = "current") {
  stopifnot(is.data.frame(data))
  cb_export(data, path, survey = survey)
}

#' Append multiple surveys to a single codebook
#'
#' @param data_list List of data frames.
#' @param path Output path for the Excel codebook.
#' @param surveys Vector of survey names.
#' @details
#' Starts from the first data frame, then appends survey-specific name/recode
#' columns for each additional survey. Variables absent in some surveys are
#' kept with missing mappings in that survey's `name_<survey>` column.
#' @return Invisibly returns the output path.
#' @export
cb_append <- function(data_list, path, surveys) {
  if (length(data_list) != length(surveys)) {
    stop("data_list and surveys must be the same length.", call. = FALSE)
  }

  first <- data_list[[1]]
  cb_export(first, path, survey = surveys[1])

  survey_sheet <- read_xlsx_sheet(path, "survey")
  for (i in seq_along(data_list)[-1]) {
    data <- data_list[[i]]
    vars <- names(data)
    name_col <- paste0("name_", surveys[i])
    recode_col <- paste0("recode_", surveys[i])

    if (!name_col %in% names(survey_sheet)) {
      survey_sheet[[name_col]] <- NA_character_
      survey_sheet[[recode_col]] <- ""
    }

    for (v in vars) {
      if (!v %in% survey_sheet$name) {
        new_row <- survey_sheet[1, , drop = FALSE]
        new_row[,] <- NA
        new_row$name <- v
        new_row$label <- ""
        new_row$type <- class(data[[v]])[1]
        new_row$choices <- ""
        new_row[[name_col]] <- v
        new_row[[recode_col]] <- ""
        survey_sheet <- rbind(survey_sheet, new_row)
      } else {
        row_idx <- match(v, survey_sheet$name)
        survey_sheet[[name_col]][row_idx] <- v
      }
    }
  }

  sheets <- list(
    survey = survey_sheet,
    choices = read_xlsx_sheet(path, "choices")
  )
  write_xlsx_sheets(path, sheets)
  invisible(path)
}

#' Build a codebook from a SurveyCTO XLSForm
#'
#' @param form Path to the XLSForm (Excel file).
#' @param path Output path for the Excel codebook.
#' @param survey Survey name to tag in the codebook.
#' @param repeat_group Which level of the form to map to data columns: `""`
#'   for fields outside any repeat (the main export), or the name of a repeat
#'   group for that repeat's long-format export (one row per instance). Give
#'   several to map several levels.
#' @param language Label language, matching a `label:<language>` column.
#'   By default the form's `label` column is used.
#' @details
#' Writes the same two sheets as [cb_export()], so the result can be edited
#' and passed to [cb_apply()], but takes variable and value labels from the
#' form rather than from data: each field's label (with HTML tags removed) and
#' the static choice list of each `select_one` field. `select_multiple` fields
#' store space-separated values and get no value labels (their list is in
#' `type`); choices loaded with `search()` are not in the form and get none
#' either.
#'
#' Every field that stores data is listed, with its innermost enclosing repeat
#' in `repeat_group`. `name_<survey>` is filled only for fields at the levels
#' in `repeat_group`; other rows are left blank, and [cb_apply()] skips rows
#' with a blank `name_<survey>`. Geopoints are listed as the four columns
#' SurveyCTO exports: `<name>-Latitude`, `-Longitude`, `-Altitude` and
#' `-Accuracy`.
#' @return Invisibly returns the output path.
#' @export
cb_from_form <- function(
  form,
  path,
  survey = "current",
  repeat_group = "",
  language = NULL
) {
  require_pkg("openxlsx")
  sheets <- openxlsx::getSheetNames(form)
  survey_sheet <- match_sheet_name(sheets, "survey")
  if (is.null(survey_sheet)) {
    stop("XLSForm has no 'survey' sheet.", call. = FALSE)
  }
  sv <- read_xlsform_sheet(form, survey_sheet)
  choices_sheet <- match_sheet_name(sheets, "choices")
  ch <- if (is.null(choices_sheet)) {
    data.frame(list_name = character(), value = character())
  } else {
    read_xlsform_sheet(form, choices_sheet)
  }

  col <- function(df, nm) {
    v <- if (!is.na(nm) && nm %in% names(df)) as.character(df[[nm]]) else rep("", nrow(df))
    v[is.na(v)] <- ""
    trimws(v)
  }
  type <- normalize_xlsform_type(sv$type)
  kind <- xls_kind(type)
  name <- col(sv, "name")
  label <- xlsform_clean_label(col(sv, xlsform_label_col(names(sv), language)))
  disabled <- tolower(col(sv, "disabled")) %in% c("yes", "true", "1")
  dynamic <- grepl("search(", tolower(col(sv, "appearance")), fixed = TRUE)

  innermost <- character(nrow(sv))
  open <- character()
  for (i in seq_len(nrow(sv))) {
    if (kind[i] == "end repeat") {
      open <- open[-length(open)]
      next
    }
    innermost[i] <- if (length(open) > 0) open[length(open)] else ""
    if (kind[i] == "begin repeat") {
      open <- c(open, name[i])
    }
  }

  keep <- xls_stores_data(kind) & nzchar(name) & !disabled
  list_name <- ifelse(
    kind == "select_one" & !dynamic & grepl("^\\S+\\s+\\S+", type),
    sub("^\\S+\\s+(\\S+).*$", "\\1", type),
    ""
  )

  rows <- lapply(which(keep), function(i) {
    if (kind[i] == "geopoint") {
      parts <- c("Latitude", "Longitude", "Altitude", "Accuracy")
      return(data.frame(
        name = paste0(name[i], "-", parts),
        label = paste0(label[i], " (", tolower(parts), ")"),
        type = type[i],
        choices = "",
        repeat_group = innermost[i],
        stringsAsFactors = FALSE
      ))
    }
    data.frame(
      name = name[i],
      label = label[i],
      type = type[i],
      choices = list_name[i],
      repeat_group = innermost[i],
      stringsAsFactors = FALSE
    )
  })
  survey_out <- do.call(rbind, rows)
  if (is.null(survey_out)) {
    stop("The form has no fields that store data.", call. = FALSE)
  }
  survey_out$label <- sanitize_text(survey_out$label)
  survey_out[[paste0("name_", survey)]] <- ifelse(
    survey_out$repeat_group %in% repeat_group,
    survey_out$name,
    ""
  )
  survey_out[[paste0("recode_", survey)]] <- ""

  value_col <- intersect(c("value", "name"), names(ch))[1]
  used <- unique(survey_out$choices[nzchar(survey_out$choices)])
  choices_out <- data.frame(
    list_name = col(ch, "list_name"),
    value = col(ch, value_col),
    label = xlsform_clean_label(col(ch, xlsform_label_col(names(ch), language))),
    stringsAsFactors = FALSE
  )
  choices_out <- choices_out[
    choices_out$list_name %in% used & nzchar(choices_out$value),
    ,
    drop = FALSE
  ]
  choices_out$label <- sanitize_text(choices_out$label)
  rownames(choices_out) <- NULL

  write_xlsx_sheets(path, list(survey = survey_out, choices = choices_out))
  invisible(path)
}

#' Label column for a language: `label`, or `label:<language>`
#' @keywords internal
#' @noRd
xlsform_label_col <- function(cols, language = NULL) {
  if (is.null(language)) {
    hit <- intersect(c("label", grep("^label:", cols, value = TRUE)), cols)
    return(if (length(hit) > 0) hit[1] else NA_character_)
  }
  hit <- cols[tolower(cols) == tolower(paste0("label:", language))]
  if (length(hit) == 0) {
    stop(sprintf("No 'label:%s' column in the form.", language), call. = FALSE)
  }
  hit[1]
}

#' Strip HTML tags and collapse whitespace in form labels
#' @keywords internal
#' @noRd
xlsform_clean_label <- function(x) {
  x <- gsub("<[^>]+>", "", x)
  trimws(gsub("\\s+", " ", x))
}

#' @keywords internal
#' @noRd
apply_recode <- function(x, recode_str) {
  if (!is.numeric(x)) {
    warning("Recode applied to non-numeric variable; skipping.")
    return(x)
  }

  tokens <- unlist(strsplit(recode_str, "[; ]+"))
  tokens <- tokens[nzchar(tokens)]

  out <- x
  changed <- rep(FALSE, length(x))
  else_value <- NULL
  for (token in tokens) {
    if (!grepl("=", token, fixed = TRUE)) {
      next
    }
    parts <- strsplit(token, "=", fixed = TRUE)[[1]]
    lhs <- parts[1]
    rhs <- parts[2]

    if (tolower(lhs) == "else") {
      else_value <- rhs
      next
    }

    if (grepl("/", lhs, fixed = TRUE)) {
      bounds <- strsplit(lhs, "/", fixed = TRUE)[[1]]
      lo <- suppressWarnings(as.numeric(bounds[1]))
      hi <- suppressWarnings(as.numeric(bounds[2]))
      if (is.na(lo) || is.na(hi)) {
        next
      }
      idx <- out >= lo & out <= hi
    } else {
      val <- suppressWarnings(as.numeric(lhs))
      if (is.na(val)) {
        next
      }
      idx <- out == val
    }

    if (rhs == ".") {
      out[idx] <- NA
    } else {
      out[idx] <- as.numeric(rhs)
    }
    changed[idx] <- TRUE
  }

  if (!is.null(else_value)) {
    idx <- !changed
    if (else_value == ".") {
      out[idx] <- NA
    } else {
      out[idx] <- as.numeric(else_value)
    }
  }

  out
}

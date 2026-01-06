#' Export a codebook to Excel
#'
#' @param data Data frame to document.
#' @param path Output path for the Excel codebook.
#' @param survey Survey name to tag in the codebook.
#' @return Invisibly returns the output path.
#' @export
iecodebook_export <- function(data, path, survey = "current") {
  stopifnot(is.data.frame(data))

  vars <- names(data)
  labels <- vapply(vars, function(v) {
    label <- attr(data[[v]], "label")
    if (is.null(label)) "" else as.character(label)
  }, character(1))

  types <- vapply(vars, function(v) {
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
  }, character(1))

  choices <- character(length(vars))
  choices_sheet <- data.frame(list_name = character(), value = character(), label = character(),
                              stringsAsFactors = FALSE)

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

    if (length(list_name) > 0 && nzchar(list_name)) {
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
#' @return Updated data frame.
#' @export
iecodebook_apply <- function(data, path, survey = "current", drop = FALSE,
                             missing_values = NULL) {
  stopifnot(is.data.frame(data))

  survey_sheet <- read_xlsx_sheet(path, "survey")
  choices_sheet <- read_xlsx_sheet(path, "choices")

  name_col <- paste0("name_", survey)
  recode_col <- paste0("recode_", survey)

  if (!name_col %in% names(survey_sheet)) {
    stop(sprintf("Survey column '%s' not found in codebook.", name_col), call. = FALSE)
  }

  survey_sheet <- as.data.frame(lapply(survey_sheet, sanitize_text), stringsAsFactors = FALSE)
  choices_sheet <- as.data.frame(lapply(choices_sheet, sanitize_text), stringsAsFactors = FALSE)

  old_names <- survey_sheet[[name_col]]
  new_names <- survey_sheet$name
  labels <- survey_sheet$label
  choices <- survey_sheet$choices
  recodes <- if (recode_col %in% names(survey_sheet)) survey_sheet[[recode_col]] else rep("", nrow(survey_sheet))

  to_drop <- character()
  rename_from <- character()
  rename_to <- character()

  for (i in seq_len(nrow(survey_sheet))) {
    old <- old_names[i]
    new <- new_names[i]

    if (is_blank(old)) next

    if (!old %in% names(data)) {
      stop(sprintf("Variable '%s' not found in data.", old), call. = FALSE)
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
      rows <- choices_sheet[choices_sheet$list_name == list_name, , drop = FALSE]
      if (nrow(rows) == 0) {
        stop(sprintf("Choice list '%s' not found in choices sheet.", list_name), call. = FALSE)
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
        labels_map <- setNames(vals, labs)
        data[[old]] <- make_labelled(data[[old]], labels = labels_map)
      } else {
        data[[old]] <- factor(data[[old]], levels = rows$value, labels = labs)
      }
    }
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
    names(data)[match(rename_from, names(data))] <- rename_to
  }

  data
}

#' Create a codebook template
#'
#' @param data Data frame to template.
#' @param path Output path for the Excel codebook.
#' @param survey Survey name to tag in the codebook.
#' @return Invisibly returns the output path.
#' @export
iecodebook_template <- function(data, path, survey = "current") {
  stopifnot(is.data.frame(data))
  iecodebook_export(data, path, survey = survey)
}

#' Append multiple surveys to a single codebook
#'
#' @param data_list List of data frames.
#' @param path Output path for the Excel codebook.
#' @param surveys Vector of survey names.
#' @return Invisibly returns the output path.
#' @export
iecodebook_append <- function(data_list, path, surveys) {
  if (length(data_list) != length(surveys)) {
    stop("data_list and surveys must be the same length.", call. = FALSE)
  }

  first <- data_list[[1]]
  iecodebook_export(first, path, survey = surveys[1])

  survey_sheet <- read_xlsx_sheet(path, "survey")
  for (i in 2:length(data_list)) {
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
    if (!grepl("=", token, fixed = TRUE)) next
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
      if (is.na(lo) || is.na(hi)) next
      idx <- out >= lo & out <= hi
    } else {
      val <- suppressWarnings(as.numeric(lhs))
      if (is.na(val)) next
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

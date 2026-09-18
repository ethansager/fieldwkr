#' Generate and apply a duplicates report
#'
#' @param data Data frame to inspect.
#' @param idvar Column name for the ID variable. This is the project-level
#'   identifier that can legitimately appear more than once (a household ID
#'   carried in from the sample frame, for example) and is the thing being
#'   adjudicated.
#' @param uniquevars Columns that uniquely identify a single submission. In
#'   SurveyCTO data this is normally `KEY`, the submission uuid the server
#'   stamps on every record. It is deliberately not the same as `idvar`: rows
#'   to drop and replacement IDs are matched on `uniquevars`, so marking one
#'   submission affects only that submission and leaves its duplicate in place.
#' @param report_path Output path for the Excel duplicates report.
#' @param keepvars Additional columns to include in the report.
#' @param apply Apply corrections from the report if it exists.
#' @param overwrite Regenerate `report_path` even if it already exists.
#' @details
#' This workflow mirrors SurveyCTO field-cleaning practice:
#' 1) create a duplicates workbook for manual adjudication,
#' 2) mark rows to drop and/or assign a replacement ID,
#' 3) re-apply the workbook to produce cleaned data.
#'
#' Steps 1 and 3 are the same call, so an existing `report_path` is never
#' overwritten by default: that would discard the adjudication made in step 2.
#' Set `overwrite = TRUE` to rebuild a stale report after the input data has
#' changed.
#'
#' The `drop` column accepts `"drop"` or `"yes"` (case-sensitive).
#' The `newid` column replaces matching `idvar` values.
#'
#' `uniquevars` must uniquely identify rows before running duplicate checks, and
#' must be present in the report for the workbook to be applied back.
#' @return Invisibly returns a list with data and report.
#' @export
duplicates <- function(
  data,
  idvar,
  uniquevars,
  report_path,
  keepvars = NULL,
  apply = TRUE,
  overwrite = FALSE
) {
  stopifnot(is.data.frame(data))

  if (!idvar %in% names(data)) {
    stop(sprintf("ID variable '%s' not found.", idvar), call. = FALSE)
  }

  id_blank <- is_blank(as.character(data[[idvar]]))
  if (any(id_blank)) {
    stop(
      sprintf("%d row(s) have a missing or blank '%s'.", sum(id_blank), idvar),
      call. = FALSE
    )
  }

  if (!all(uniquevars %in% names(data))) {
    stop("Some uniquevars are not in the data.", call. = FALSE)
  }

  key_blank <- Reduce(
    `|`,
    lapply(uniquevars, function(v) is_blank(as.character(data[[v]])))
  )
  if (any(key_blank)) {
    stop(
      sprintf(
        paste0(
          "%d row(s) have a missing or blank value in uniquevars (%s). ",
          "In SurveyCTO data these are usually records with no submission; ",
          "drop them before checking duplicates."
        ),
        sum(key_blank),
        paste(uniquevars, collapse = ", ")
      ),
      call. = FALSE
    )
  }

  if (any(duplicated(data[uniquevars]))) {
    stop("uniquevars do not uniquely identify observations.", call. = FALSE)
  }

  report <- build_duplicates_report(data, idvar, uniquevars, keepvars)
  if (nrow(report) > 0) {
    if (overwrite || !file.exists(report_path)) {
      write_xlsx_sheets(report_path, list(duplicates = report))
    } else {
      message(sprintf(
        "Report '%s' already exists; applying it as-is. Use overwrite = TRUE to rebuild it.",
        report_path
      ))
    }
  }

  if (!apply) {
    return(invisible(list(data = data, report = report)))
  }

  if (!file.exists(report_path)) {
    return(invisible(list(data = data, report = report)))
  }

  updated <- apply_duplicates_report(data, idvar, uniquevars, report_path)
  invisible(list(data = updated, report = report))
}

#' @keywords internal
#' @noRd
build_duplicates_report <- function(data, idvar, uniquevars, keepvars) {
  dup_idx <- duplicated(data[[idvar]]) |
    duplicated(data[[idvar]], fromLast = TRUE)
  dup_data <- data[dup_idx, , drop = FALSE]
  if (nrow(dup_data) == 0) {
    return(data.frame())
  }

  ids <- unique(dup_data[[idvar]])
  report_rows <- list()
  duplistid <- 1
  today <- format(Sys.Date(), "%Y%m%d")

  for (id in ids) {
    group <- dup_data[dup_data[[idvar]] == id, , drop = FALSE]
    if (nrow(group) < 2) {
      next
    }

    listofdiffs <- ""
    if (nrow(group) == 2) {
      diffvars <- comp_dup(
        group,
        idvar = idvar,
        id = id,
        more2ok = TRUE
      )$diffvars
      if (length(diffvars) > 0) {
        listofdiffs <- paste(diffvars, collapse = " ")
        if (nchar(listofdiffs) > 250) {
          listofdiffs <- paste0(
            substr(listofdiffs, 1, 200),
            " ||| List truncated, use comp_dup for full list"
          )
        }
      }
    } else {
      listofdiffs <- "Cannot list differences for groups with 3+ duplicates."
    }

    rows <- group
    rows$duplistid <- duplistid
    rows$datelisted <- today
    rows$datefixed <- ""
    rows$correct <- ""
    rows$drop <- ""
    rows$newid <- ""
    rows$initials <- ""
    rows$notes <- ""
    rows$listofdiffs <- listofdiffs

    keep <- c(
      idvar,
      uniquevars,
      keepvars,
      "duplistid",
      "datelisted",
      "datefixed",
      "correct",
      "drop",
      "newid",
      "initials",
      "notes",
      "listofdiffs"
    )
    keep <- unique(keep[keep %in% names(rows)])

    report_rows[[length(report_rows) + 1]] <- rows[, keep, drop = FALSE]
    duplistid <- duplistid + 1
  }

  do.call(rbind, report_rows)
}

#' @keywords internal
#' @noRd
apply_duplicates_report <- function(data, idvar, uniquevars, report_path) {
  report <- read_xlsx_sheet(report_path, "duplicates")
  if (!idvar %in% names(report)) {
    stop("Report is missing ID variable column.", call. = FALSE)
  }
  if (!all(uniquevars %in% names(report))) {
    stop("Report is missing uniquevars column(s).", call. = FALSE)
  }

  row_key <- function(df) {
    do.call(
      paste,
      c(lapply(uniquevars, function(v) as.character(df[[v]])), sep = "\r")
    )
  }
  data_key <- row_key(data)

  if ("drop" %in% names(report)) {
    drops <- report[report$drop %in% c("drop", "yes"), , drop = FALSE]
    if (nrow(drops) > 0) {
      data <- data[!data_key %in% row_key(drops), , drop = FALSE]
      data_key <- row_key(data)
    }
  }

  if ("newid" %in% names(report)) {
    updates <- report[!is_blank(report$newid), , drop = FALSE]
    if (nrow(updates) > 0) {
      idx <- match(data_key, row_key(updates))
      has_new <- !is.na(idx)
      if (any(has_new)) {
        new_ids <- as.character(updates$newid[idx[has_new]])
        if (is.numeric(data[[idvar]])) {
          converted <- suppressWarnings(as.numeric(new_ids))
          if (any(is.na(converted))) {
            stop(
              "Non-numeric newid value for a numeric ID variable.",
              call. = FALSE
            )
          }
          new_ids <- converted
        }
        data[[idvar]][has_new] <- new_ids
      }
    }
  }

  data
}

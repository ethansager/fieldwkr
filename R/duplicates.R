#' Generate and apply a duplicates report
#'
#' @param data Data frame to inspect.
#' @param idvar Column name for the ID variable.
#' @param uniquevars Columns that uniquely identify observations.
#' @param report_path Output path for the Excel duplicates report.
#' @param keepvars Additional columns to include in the report.
#' @param apply Apply corrections from the report if it exists.
#' @details
#' This workflow mirrors SurveyCTO field-cleaning practice:
#' 1) create a duplicates workbook for manual adjudication,
#' 2) mark rows to drop and/or assign a replacement ID,
#' 3) re-apply the workbook to produce cleaned data.
#'
#' The `drop` column accepts `"drop"` or `"yes"` (case-sensitive).
#' The `newid` column replaces matching `idvar` values.
#'
#' `uniquevars` must uniquely identify rows before running duplicate checks.
#' @return Invisibly returns a list with data and report.
#' @export
duplicates <- function(
  data,
  idvar,
  uniquevars,
  report_path,
  keepvars = NULL,
  apply = TRUE
) {
  stopifnot(is.data.frame(data))

  if (!idvar %in% names(data)) {
    stop(sprintf("ID variable '%s' not found.", idvar), call. = FALSE)
  }

  if (any(is.na(data[[idvar]]))) {
    stop("Missing values in ID variable.", call. = FALSE)
  }

  if (!all(uniquevars %in% names(data))) {
    stop("Some uniquevars are not in the data.", call. = FALSE)
  }

  if (any(duplicated(data[uniquevars]))) {
    stop("uniquevars do not uniquely identify observations.", call. = FALSE)
  }

  report <- build_duplicates_report(data, idvar, uniquevars, keepvars)
  if (nrow(report) > 0) {
    write_xlsx_sheets(report_path, list(duplicates = report))
  }

  if (!apply) {
    return(invisible(list(data = data, report = report)))
  }

  if (!file.exists(report_path)) {
    return(invisible(list(data = data, report = report)))
  }

  updated <- apply_duplicates_report(data, idvar, report_path)
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
apply_duplicates_report <- function(data, idvar, report_path) {
  report <- read_xlsx_sheet(report_path, "duplicates")
  if (!idvar %in% names(report)) {
    stop("Report is missing ID variable column.", call. = FALSE)
  }

  if ("drop" %in% names(report)) {
    drops <- report[report$drop %in% c("drop", "yes"), , drop = FALSE]
    if (nrow(drops) > 0) {
      ids_to_drop <- unique(drops[[idvar]])
      data <- data[!data[[idvar]] %in% ids_to_drop, , drop = FALSE]
    }
  }

  if ("newid" %in% names(report)) {
    updates <- report[!is_blank(report$newid), , drop = FALSE]
    if (nrow(updates) > 0) {
      for (i in seq_len(nrow(updates))) {
        old_id <- updates[[idvar]][i]
        new_id <- updates$newid[i]
        data[[idvar]][data[[idvar]] == old_id] <- new_id
      }
    }
  }

  data
}

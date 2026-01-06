#' Create a corrections template workbook
#'
#' @param path Output path for the Excel template.
#' @param idvars ID columns to include in the template.
#' @return Invisibly returns the output path.
#' @export
iecorrect_template <- function(path, idvars) {
  if (length(idvars) == 0) {
    stop("idvars must be provided.", call. = FALSE)
  }

  base_cols <- c(idvars, "varname", "value", "valuecurrent", "initials", "notes")
  drop_cols <- c(idvars, "n_obs", "initials", "notes")

  empty_sheet <- function(cols) {
    as.data.frame(setNames(replicate(length(cols), character(0), simplify = FALSE), cols))
  }

  sheets <- list(
    string = empty_sheet(base_cols),
    numeric = empty_sheet(base_cols),
    drop = empty_sheet(drop_cols)
  )

  write_xlsx_sheets(path, sheets)
  invisible(path)
}

#' Apply corrections from a template workbook
#'
#' @param data Data frame to update.
#' @param path Path to the Excel corrections workbook.
#' @param idvars ID columns used to match records.
#' @param sheets Sheets to apply.
#' @return Updated data frame.
#' @export
iecorrect_apply <- function(data, path, idvars, sheets = c("numeric", "string", "drop")) {
  stopifnot(is.data.frame(data))

  if (!all(idvars %in% names(data))) {
    stop("Some idvars are not in the data.", call. = FALSE)
  }

  if ("numeric" %in% sheets) {
    data <- apply_corrections_sheet(data, path, idvars, "numeric", numeric = TRUE)
  }
  if ("string" %in% sheets) {
    data <- apply_corrections_sheet(data, path, idvars, "string", numeric = FALSE)
  }
  if ("drop" %in% sheets) {
    drop_sheet <- read_xlsx_sheet(path, "drop")
    if (nrow(drop_sheet) > 0) {
      data <- drop_rows_by_sheet(data, drop_sheet, idvars)
    }
  }

  data
}

#' @keywords internal
#' @noRd
apply_corrections_sheet <- function(data, path, idvars, sheet, numeric = FALSE) {
  sheet_data <- read_xlsx_sheet(path, sheet)
  if (nrow(sheet_data) == 0) return(data)

  required <- c(idvars, "varname", "value")
  if (!all(required %in% names(sheet_data))) {
    stop(sprintf("Sheet '%s' missing required columns.", sheet), call. = FALSE)
  }

  for (i in seq_len(nrow(sheet_data))) {
    row <- sheet_data[i, , drop = FALSE]
    varname <- row$varname
    if (is_blank(varname) || !varname %in% names(data)) next

    value <- row$value
    if (numeric) value <- suppressWarnings(as.numeric(value))

    idx <- rep(TRUE, nrow(data))
    for (id in idvars) {
      val <- row[[id]]
      if (is_blank(val) || val == "*") next
      idx <- idx & as.character(data[[id]]) == as.character(val)
    }

    if ("valuecurrent" %in% names(row) && !is_blank(row$valuecurrent)) {
      cur <- row$valuecurrent
      if (numeric) cur <- suppressWarnings(as.numeric(cur))
      idx <- idx & data[[varname]] == cur
    }

    data[[varname]][idx] <- value
  }

  data
}

#' @keywords internal
#' @noRd
drop_rows_by_sheet <- function(data, sheet_data, idvars) {
  for (i in seq_len(nrow(sheet_data))) {
    row <- sheet_data[i, , drop = FALSE]
    idx <- rep(TRUE, nrow(data))
    for (id in idvars) {
      val <- row[[id]]
      if (is_blank(val) || val == "*") next
      idx <- idx & as.character(data[[id]]) == as.character(val)
    }
    data <- data[!idx, , drop = FALSE]
  }
  data
}

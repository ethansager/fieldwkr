#' Create a corrections template workbook
#'
#' @param path Output path for the Excel template.
#' @param idvars ID columns to include in the template.
#' @details
#' The template contains three sheets:
#' - `numeric`: numeric edits keyed by `idvars`, with `varname`, `value`, and
#'   optional `valuecurrent` guards.
#' - `string`: string edits with the same structure.
#' - `drop`: row-level drops keyed by `idvars`.
#'
#' Every `idvars` cell must be filled in on every row. There is no wildcard.
#'
#' Use [correct_apply()] to apply completed templates back to data.
#' @return Invisibly returns the output path.
#' @export
correct_temp <- function(path, idvars) {
  if (length(idvars) == 0) {
    stop("idvars must be provided.", call. = FALSE)
  }

  base_cols <- c(
    idvars,
    "varname",
    "value",
    "valuecurrent",
    "initials",
    "notes"
  )
  drop_cols <- c(idvars, "n_obs", "initials", "notes")

  empty_sheet <- function(cols) {
    as.data.frame(setNames(
      replicate(length(cols), character(0), simplify = FALSE),
      cols
    ))
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
#' @details
#' Row matching is conjunctive across all `idvars`: every ID column must be
#' filled in on every correction row, and all of them must match. There is no
#' wildcard, because a cell that matches every record is far more often a
#' data-entry slip than an intention. A blank cell is an error.
#'
#' If a `valuecurrent` column is present in `numeric`/`string` sheets, edits are
#' applied only when the current value also matches.
#' @return Updated data frame.
#' @export
correct_apply <- function(
  data,
  path,
  idvars,
  sheets = c("numeric", "string", "drop")
) {
  stopifnot(is.data.frame(data))

  if (!all(idvars %in% names(data))) {
    stop("Some idvars are not in the data.", call. = FALSE)
  }

  if ("numeric" %in% sheets) {
    data <- apply_corrections_sheet(
      data,
      path,
      idvars,
      "numeric",
      numeric = TRUE
    )
  }
  if ("string" %in% sheets) {
    data <- apply_corrections_sheet(
      data,
      path,
      idvars,
      "string",
      numeric = FALSE
    )
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
apply_corrections_sheet <- function(
  data,
  path,
  idvars,
  sheet,
  numeric = FALSE
) {
  sheet_data <- read_xlsx_sheet(path, sheet)
  if (nrow(sheet_data) == 0) {
    return(data)
  }

  required <- c(idvars, "varname", "value")
  if (!all(required %in% names(sheet_data))) {
    stop(sprintf("Sheet '%s' missing required columns.", sheet), call. = FALSE)
  }

  for (i in seq_len(nrow(sheet_data))) {
    row <- sheet_data[i, , drop = FALSE]
    varname <- row$varname
    if (is_blank(varname)) {
      stop(
        sprintf("Sheet '%s' row %s: blank varname.", sheet, i),
        call. = FALSE
      )
    }
    if (!varname %in% names(data)) {
      stop(
        sprintf(
          "Sheet '%s' row %s: variable '%s' is not in the data.",
          sheet,
          i,
          varname
        ),
        call. = FALSE
      )
    }

    value <- row$value
    if (numeric) {
      if (is_blank(value)) {
        value <- NA_real_
      } else {
        parsed <- suppressWarnings(as.numeric(value))
        if (is.na(parsed)) {
          stop(
            sprintf(
              "Sheet '%s' row %s: value '%s' for variable '%s' is not numeric.",
              sheet,
              i,
              value,
              varname
            ),
            call. = FALSE
          )
        }
        value <- parsed
      }
    }

    idx <- corrections_match(data, row, idvars)

    if ("valuecurrent" %in% names(row) && !is_blank(row$valuecurrent)) {
      cur <- row$valuecurrent
      if (numeric) {
        cur <- suppressWarnings(as.numeric(cur))
      }
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
    idx <- corrections_match(data, row, idvars)
    data <- data[!idx, , drop = FALSE]
  }
  data
}

#' @keywords internal
#' @noRd
#' Rows of `data` matched by one correction row, conjunctively across `idvars`.
#' A comparison against a missing value in the data yields FALSE, not NA: an
#' NA index would fabricate all-NA rows when used to subset.
corrections_match <- function(data, row, idvars) {
  idx <- rep(TRUE, nrow(data))
  for (id in idvars) {
    val <- row[[id]]
    if (is_blank(val)) {
      stop(
        sprintf(
          "Blank '%s' value in a corrections row; every ID column must be filled in.",
          id
        ),
        call. = FALSE
      )
    }
    cmp <- as.character(data[[id]]) == as.character(val)
    idx <- idx & !is.na(cmp) & cmp
  }
  idx
}

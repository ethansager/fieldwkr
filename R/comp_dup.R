#' Compare two duplicate records
#'
#' @param data Data frame containing duplicates.
#' @param idvar Column name for the ID variable.
#' @param id The ID value to compare.
#' @param didifference Print differing variables when TRUE.
#' @param keepdifference Return a data frame of differing variables.
#' @param keepother Additional variables to keep in the returned data.
#' @param more2ok Allow more than two duplicates by taking the first two.
#' @param filter Optional logical vector to filter rows for comparison.
#' @details
#' This function is intended for adjudicating duplicate IDs during survey
#' cleaning. It compares exactly two records (or the first two when
#' `more2ok = TRUE`) and reports variables with matching versus differing
#' values.
#' @return A list with match/diff variables and counts. If keepdifference is
#'   TRUE, includes a data element with the subset.
#' @export
comp_dup <- function(
  data,
  idvar,
  id,
  didifference = FALSE,
  keepdifference = FALSE,
  keepother = NULL,
  more2ok = FALSE,
  filter = NULL
) {
  stopifnot(is.data.frame(data))
  if (!idvar %in% names(data)) {
    stop(sprintf("ID variable '%s' not found in data.", idvar), call. = FALSE)
  }

  id_values <- data[[idvar]]
  if (any(is.na(id_values))) {
    stop("The ID variable contains missing values.", call. = FALSE)
  }

  if (is.numeric(id_values)) {
    if (any(id_values %% 1 != 0)) {
      stop("The ID variable must be integer-like or a string.", call. = FALSE)
    }
  }

  data$id__tmp <- as.character(id_values)
  target <- as.character(id)

  subset_data <- data[data$id__tmp == target, , drop = FALSE]
  if (!is.null(filter)) {
    subset_data <- subset_data[filter, , drop = FALSE]
  }

  if (nrow(subset_data) == 0) {
    stop(sprintf("No rows found for %s == %s.", idvar, id), call. = FALSE)
  }

  if (nrow(subset_data) > 2) {
    if (!more2ok && is.null(filter)) {
      stop(
        "More than two duplicates found; use more2ok or filter to select two.",
        call. = FALSE
      )
    }
    subset_data <- subset_data[seq_len(2), , drop = FALSE]
  }

  if (nrow(subset_data) < 2) {
    stop("Need at least two observations to compare.", call. = FALSE)
  }

  match_vars <- character()
  diff_vars <- character()

  for (var in setdiff(names(subset_data), "id__tmp")) {
    v1 <- subset_data[[var]][1]
    v2 <- subset_data[[var]][2]

    if (all(is.na(c(v1, v2)))) {
      next
    }

    if (identical(v1, v2)) {
      match_vars <- c(match_vars, var)
    } else {
      diff_vars <- c(diff_vars, var)
    }
  }

  match_vars <- setdiff(match_vars, idvar)

  if (didifference) {
    message(
      "The following variables have different values across the duplicates:"
    )
    message(paste(diff_vars, collapse = " "))
  }

  num_non_missing <- length(match_vars) + length(diff_vars)
  message(sprintf(
    "The duplicate observations with ID = %s have non-missing values in %s variables.",
    id,
    num_non_missing
  ))
  message(sprintf(
    "%s variable(s) are identical across the duplicates",
    length(match_vars)
  ))
  message(sprintf(
    "%s variable(s) have different values across the duplicates",
    length(diff_vars)
  ))

  result <- list(
    matchvars = match_vars,
    diffvars = diff_vars,
    nummatch = length(match_vars),
    numdiff = length(diff_vars),
    numnomiss = num_non_missing
  )

  if (keepdifference && length(diff_vars) > 0) {
    keep <- c(idvar, diff_vars, keepother)
    keep <- unique(keep[keep %in% names(data)])
    result$data <- data[data$id__tmp == target, keep, drop = FALSE]
  }

  data$id__tmp <- NULL
  result
}

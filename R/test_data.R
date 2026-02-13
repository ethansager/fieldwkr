#' Validate a dataset for unit-test checks
#'
#' @param data Data frame to validate.
#' @param required_cols Columns that must exist in `data`.
#' @param unique_keys Columns that must uniquely identify each row.
#' @param non_missing_cols Columns that cannot contain missing or blank values.
#' @param value_ranges Named list of inclusive numeric ranges, e.g.
#'   `list(age = c(0, 120))`.
#' @param allowed_values Named list of allowed values per column, e.g.
#'   `list(gender = c("female", "male"))`.
#' @param min_rows Minimum allowed row count.
#' @param max_rows Maximum allowed row count.
#' @param verbose Print summary output.
#'
#' @details
#' Returns validation results as `errors` and `warnings` vectors for use in
#' `testthat` checks and CI.
#'
#' @return A list with errors and warnings.
#' @export
test_data <- function(
  data,
  required_cols = NULL,
  unique_keys = NULL,
  non_missing_cols = NULL,
  value_ranges = NULL,
  allowed_values = NULL,
  min_rows = NULL,
  max_rows = NULL,
  verbose = TRUE
) {
  stopifnot(is.data.frame(data))

  errors <- character()
  warnings <- character()

  add_error <- function(msg) {
    errors <<- c(errors, msg)
  }
  add_warning <- function(msg) {
    warnings <<- c(warnings, msg)
  }

  required_cols <- normalize_test_data_cols(required_cols, "required_cols")
  unique_keys <- normalize_test_data_cols(unique_keys, "unique_keys")
  non_missing_cols <- normalize_test_data_cols(
    non_missing_cols,
    "non_missing_cols"
  )

  value_ranges <- normalize_test_data_named_list(value_ranges, "value_ranges")
  allowed_values <- normalize_test_data_named_list(
    allowed_values,
    "allowed_values"
  )

  min_rows <- normalize_test_data_count(min_rows, "min_rows")
  max_rows <- normalize_test_data_count(max_rows, "max_rows")
  if (!is.null(min_rows) && !is.null(max_rows) && min_rows > max_rows) {
    stop("'min_rows' cannot be greater than 'max_rows'.", call. = FALSE)
  }

  n <- nrow(data)
  if (n == 0) {
    add_warning("Data has 0 rows.")
  }

  if (!is.null(min_rows) && n < min_rows) {
    add_error(sprintf("Data has %s rows; minimum required is %s.", n, min_rows))
  }
  if (!is.null(max_rows) && n > max_rows) {
    add_error(sprintf("Data has %s rows; maximum allowed is %s.", n, max_rows))
  }

  missing_required <- setdiff(required_cols, names(data))
  if (length(missing_required) > 0) {
    add_error(sprintf(
      "Missing required column(s): %s",
      paste(missing_required, collapse = ", ")
    ))
  }

  if (length(unique_keys) > 0) {
    missing_keys <- setdiff(unique_keys, names(data))
    if (length(missing_keys) > 0) {
      add_error(sprintf(
        "Column(s) in unique_keys not found in data: %s",
        paste(missing_keys, collapse = ", ")
      ))
    } else {
      key_missing <- rep(FALSE, n)
      for (col in unique_keys) {
        key_missing <- key_missing | test_data_is_missing(data[[col]])
      }
      if (any(key_missing)) {
        add_error(sprintf(
          "Missing values in unique key column(s) '%s' at rows: %s",
          paste(unique_keys, collapse = ", "),
          format_test_data_rows(which(key_missing))
        ))
      }

      dup <- duplicated(data[unique_keys]) | duplicated(
        data[unique_keys],
        fromLast = TRUE
      )
      if (any(dup)) {
        add_error(sprintf(
          "Duplicate key combinations found for unique_keys '%s' at rows: %s",
          paste(unique_keys, collapse = ", "),
          format_test_data_rows(which(dup))
        ))
      }
    }
  }

  if (length(non_missing_cols) > 0) {
    missing_non_missing <- setdiff(non_missing_cols, names(data))
    if (length(missing_non_missing) > 0) {
      add_error(sprintf(
        "Column(s) in non_missing_cols not found in data: %s",
        paste(missing_non_missing, collapse = ", ")
      ))
    }

    for (col in intersect(non_missing_cols, names(data))) {
      missing_rows <- which(test_data_is_missing(data[[col]]))
      if (length(missing_rows) > 0) {
        add_error(sprintf(
          "Missing values in '%s' at rows: %s",
          col,
          format_test_data_rows(missing_rows)
        ))
      }
    }
  }

  if (length(value_ranges) > 0) {
    missing_range_cols <- setdiff(names(value_ranges), names(data))
    if (length(missing_range_cols) > 0) {
      add_error(sprintf(
        "Column(s) in value_ranges not found in data: %s",
        paste(missing_range_cols, collapse = ", ")
      ))
    }

    for (col in intersect(names(value_ranges), names(data))) {
      bounds <- value_ranges[[col]]
      if (length(bounds) != 2) {
        add_error(sprintf(
          "Range for column '%s' must have exactly two values.",
          col
        ))
        next
      }

      bounds_num <- suppressWarnings(as.numeric(bounds))
      if (any(is.na(bounds_num))) {
        add_error(sprintf(
          "Range for column '%s' must be numeric.",
          col
        ))
        next
      }

      min_val <- bounds_num[1]
      max_val <- bounds_num[2]
      if (min_val > max_val) {
        add_error(sprintf(
          "Range for column '%s' is invalid: min is greater than max.",
          col
        ))
        next
      }

      x <- data[[col]]
      if (is.factor(x)) {
        x <- as.character(x)
      }
      x_num <- suppressWarnings(as.numeric(x))

      cast_fail <- !is.na(x) & is.na(x_num)
      if (is.character(x)) {
        cast_fail <- cast_fail & trimws(x) != ""
      }
      if (any(cast_fail)) {
        add_error(sprintf(
          "Non-numeric value(s) found in '%s' for range check at rows: %s",
          col,
          format_test_data_rows(which(cast_fail))
        ))
        next
      }

      out_of_range <- !is.na(x_num) & (x_num < min_val | x_num > max_val)
      if (any(out_of_range)) {
        add_error(sprintf(
          "Values outside [%s, %s] found in '%s' at rows: %s",
          min_val,
          max_val,
          col,
          format_test_data_rows(which(out_of_range))
        ))
      }
    }
  }

  if (length(allowed_values) > 0) {
    missing_allowed_cols <- setdiff(names(allowed_values), names(data))
    if (length(missing_allowed_cols) > 0) {
      add_error(sprintf(
        "Column(s) in allowed_values not found in data: %s",
        paste(missing_allowed_cols, collapse = ", ")
      ))
    }

    for (col in intersect(names(allowed_values), names(data))) {
      allowed <- as.character(allowed_values[[col]])
      if (length(allowed) == 0) {
        add_warning(sprintf(
          "No allowed values supplied for '%s'; skipping check.",
          col
        ))
        next
      }

      x <- data[[col]]
      x_chr <- as.character(x)
      missing_rows <- test_data_is_missing(x)
      invalid <- !missing_rows & !(x_chr %in% allowed)

      if (any(invalid)) {
        bad_values <- unique(x_chr[invalid])
        bad_values <- bad_values[seq_len(min(length(bad_values), 5))]
        add_error(sprintf(
          "Unexpected value(s) in '%s': %s (rows: %s)",
          col,
          paste(shQuote(bad_values), collapse = ", "),
          format_test_data_rows(which(invalid))
        ))
      }
    }
  }

  if (verbose) {
    message(sprintf("Errors: %s", length(errors)))
    if (length(errors) > 0) {
      message(paste(errors, collapse = "\n"))
    }
    message(sprintf("Warnings: %s", length(warnings)))
    if (length(warnings) > 0) {
      message(paste(warnings, collapse = "\n"))
    }
  }

  list(errors = errors, warnings = warnings)
}

#' @keywords internal
#' @noRd
normalize_test_data_cols <- function(x, arg_name) {
  if (is.null(x)) {
    return(character())
  }
  if (!is.character(x)) {
    stop(sprintf("'%s' must be a character vector.", arg_name), call. = FALSE)
  }
  x <- trimws(x)
  unique(x[nzchar(x)])
}

#' @keywords internal
#' @noRd
normalize_test_data_named_list <- function(x, arg_name) {
  if (is.null(x)) {
    return(list())
  }
  if (!is.list(x)) {
    stop(sprintf("'%s' must be a named list.", arg_name), call. = FALSE)
  }
  nm <- names(x)
  if (is.null(nm) || any(trimws(nm) == "")) {
    stop(sprintf("'%s' must be a named list.", arg_name), call. = FALSE)
  }
  x
}

#' @keywords internal
#' @noRd
normalize_test_data_count <- function(x, arg_name) {
  if (is.null(x)) {
    return(NULL)
  }
  if (length(x) != 1 || !is.numeric(x) || is.na(x)) {
    stop(sprintf("'%s' must be a single non-missing number.", arg_name), call. = FALSE)
  }
  if (x < 0 || x %% 1 != 0) {
    stop(sprintf("'%s' must be a non-negative integer.", arg_name), call. = FALSE)
  }
  as.integer(x)
}

#' @keywords internal
#' @noRd
test_data_is_missing <- function(x) {
  missing <- is.na(x)
  if (is.factor(x)) {
    x <- as.character(x)
  }
  if (is.character(x)) {
    missing <- missing | trimws(x) == ""
  }
  missing
}

#' @keywords internal
#' @noRd
format_test_data_rows <- function(rows, max_rows = 10) {
  rows <- unique(as.integer(rows))
  rows <- rows[!is.na(rows)]
  rows <- sort(rows)
  if (length(rows) == 0) {
    return("")
  }
  if (length(rows) <= max_rows) {
    return(paste(rows, collapse = ", "))
  }
  shown <- paste(rows[seq_len(max_rows)], collapse = ", ")
  sprintf("%s ... (+%s more)", shown, length(rows) - max_rows)
}

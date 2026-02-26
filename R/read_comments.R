#' Read SurveyCTO comment CSVs into a wide data frame
#'
#' @param path Character. Folder containing `Comments-[UUID].csv` files.
#' @return Wide data frame with one row per UUID and one column per field name
#'   (using the final path segment). A `uuid` column is always present and first.
#' @export
read_comments <- function(path) {
  files <- list.files(path, pattern = "^Comments-.+\\.csv$", full.names = TRUE)
  if (length(files) == 0) {
    stop("No Comments-*.csv files found in: ", path, call. = FALSE)
  }

  rows <- lapply(files, function(f) {
    uuid <- sub("^Comments-(.+)\\.csv$", "\\1", basename(f))
    dat <- read.csv(f, stringsAsFactors = FALSE, check.names = FALSE)

    if (nrow(dat) == 0) {
      return(data.frame(uuid = uuid, stringsAsFactors = FALSE))
    }

    field_names <- sub("^.*/", "", dat[["Field name"]])
    comments <- dat[["Comment"]]

    # Number duplicate field names: suffix all occurrences with _1, _2, ...
    dups <- names(which(table(field_names) > 1))
    if (length(dups) > 0) {
      idx <- field_names %in% dups
      field_names[idx] <- paste0(
        field_names[idx], "_",
        ave(seq_along(field_names)[idx], field_names[idx], FUN = seq_along)
      )
    }

    row <- as.data.frame(as.list(setNames(comments, field_names)),
                         stringsAsFactors = FALSE, check.names = FALSE)
    row$uuid <- uuid
    row
  })

  # Union of all columns, filling missing with NA_character_
  all_cols <- unique(c("uuid", unlist(lapply(rows, names))))
  rows <- lapply(rows, function(r) {
    missing <- setdiff(all_cols, names(r))
    r[missing] <- NA_character_
    r[all_cols]
  })

  result <- do.call(rbind, rows)
  rownames(result) <- NULL
  result
}

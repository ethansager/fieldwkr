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

    required <- c("Field name", "Comment")
    if (!all(required %in% names(dat))) {
      stop(
        sprintf(
          "'%s' is missing column(s): %s",
          basename(f),
          paste(setdiff(required, names(dat)), collapse = ", ")
        ),
        call. = FALSE
      )
    }

    if (nrow(dat) == 0) {
      return(data.frame(uuid = uuid, stringsAsFactors = FALSE))
    }

    field_names <- sub("^.*/", "", dat[["Field name"]])
    comments <- dat[["Comment"]]

    # Repeat comments on one field: the first keeps the bare field name so a
    # field's column is the same whether or not it was commented on twice.
    # Later ones become _2, _3, ...
    occurrence <- ave(seq_along(field_names), field_names, FUN = seq_along)
    field_names <- ifelse(
      occurrence == 1,
      field_names,
      paste0(field_names, "_", occurrence)
    )

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

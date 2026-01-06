#' Generate mock data from a SurveyCTO/ODK XLSForm
#'
#' @param path Path to the XLSForm (Excel file).
#' @param n Number of mock rows to generate.
#' @param seed Optional random seed.
#' @param max_tries Maximum attempts to satisfy constraints.
#' @param max_repeat Maximum repeat loops when repeat_count is missing.
#' @return A data frame of mock responses.
#' @export
iemockdata_xlsform <- function(path, n = 1, seed = NULL, max_tries = 50, max_repeat = 3) {
  if (!is.null(seed)) set.seed(seed)

  survey <- read_xlsx_sheet(path, "survey")
  choices <- read_xlsx_sheet(path, "choices")

  results <- vector("list", n)
  all_names <- character()
  for (i in seq_len(n)) {
    row_list <- simulate_xlsform_once(survey, choices, max_tries, max_repeat)
    results[[i]] <- row_list
    all_names <- union(all_names, names(row_list))
  }

  rows <- lapply(results, function(x) {
    missing <- setdiff(all_names, names(x))
    if (length(missing) > 0) {
      x[missing] <- NA
    }
    as.data.frame(x[all_names], stringsAsFactors = FALSE)
  })

  do.call(rbind, rows)
}

simulate_xlsform_once <- function(survey, choices, max_tries, max_repeat) {
  survey <- as.data.frame(survey, stringsAsFactors = FALSE)
  choices <- as.data.frame(choices, stringsAsFactors = FALSE)

  answers <- list()
  end_row <- nrow(survey)
  process_rows(survey, choices, 1, end_row, answers, integer(0), max_tries, max_repeat)$answers
}

process_rows <- function(survey, choices, start_row, end_row, answers, repeat_stack,
                         max_tries, max_repeat) {
  i <- start_row
  while (i <= end_row) {
    type <- survey$type[i]
    type <- if (is.na(type)) "" else as.character(type)
    type_parts <- strsplit(type, " ")[[1]]
    base_type <- type_parts[1]

    if (base_type == "begin" && length(type_parts) >= 2 && type_parts[2] == "repeat") {
      end_repeat <- find_end_repeat(survey, i, end_row)
      if (end_repeat < i) {
        i <- i + 1
        next
      }

      if (should_skip_row(survey, i, answers, repeat_stack)) {
        i <- end_repeat + 1
        next
      }

      loops <- get_repeat_loops(survey, i, answers, repeat_stack, max_repeat)
      for (loop in seq_len(loops)) {
        answers <- process_rows(
          survey,
          choices,
          i + 1,
          end_repeat - 1,
          answers,
          c(repeat_stack, loop),
          max_tries,
          max_repeat
        )$answers
      }
      i <- end_repeat + 1
      next
    }

    if (base_type == "end" && length(type_parts) >= 2 && type_parts[2] == "repeat") {
      i <- i + 1
      next
    }

    if (type %in% c("begin group", "end group")) {
      i <- i + 1
      next
    }

    name <- survey$name[i]
    if (is.na(name) || trimws(name) == "") {
      i <- i + 1
      next
    }
    name <- as.character(name)
    full_name <- make_repeat_name(name, repeat_stack)

    if (base_type == "calculate") {
      calc <- get_col_value(survey, "calculation", i)
      if (!is.na(calc) && nzchar(calc)) {
        answers[[full_name]] <- eval_xlsform_expr(calc, answers, repeat_stack)
      }
      i <- i + 1
      next
    }

    if (should_skip_row(survey, i, answers, repeat_stack)) {
      answers[[full_name]] <- NA
      i <- i + 1
      next
    }

    value <- sample_value(survey, choices, i, repeat_stack)
    constraint <- get_col_value(survey, "constraint", i)
    if (!is.na(constraint) && nzchar(constraint)) {
      tries <- 1
      ok <- FALSE
      while (tries <= max_tries) {
        tmp <- answers
        tmp[[full_name]] <- value
        ok <- isTRUE(eval_xlsform_expr(constraint, tmp, repeat_stack))
        if (ok) break
        value <- sample_value(survey, choices, i, repeat_stack)
        tries <- tries + 1
      }
      if (!ok) value <- NA
    }

    answers[[full_name]] <- value
    i <- i + 1
  }

  list(answers = answers, row = i)
}

find_end_repeat <- function(survey, start_row, end_row) {
  depth <- 0
  for (i in seq(start_row, end_row)) {
    type <- as.character(survey$type[i])
    if (type == "begin repeat") depth <- depth + 1
    if (type == "end repeat") {
      depth <- depth - 1
      if (depth == 0) return(i)
    }
  }
  -1
}

get_repeat_loops <- function(survey, row, answers, repeat_stack, max_repeat) {
  repeat_count <- get_col_value(survey, "repeat_count", row)
  if (!is.na(repeat_count) && nzchar(repeat_count)) {
    count <- eval_xlsform_expr(repeat_count, answers, repeat_stack)
    count <- suppressWarnings(as.integer(count))
    if (!is.na(count) && count > 0) {
      return(min(count, max_repeat))
    }
  }
  sample(seq_len(max_repeat), 1)
}

should_skip_row <- function(survey, row, answers, repeat_stack) {
  rel <- get_col_value(survey, "relevance", row)
  if (is.na(rel) || !nzchar(rel)) return(FALSE)
  isTRUE(eval_xlsform_expr(rel, answers, repeat_stack)) == FALSE
}

sample_value <- function(survey, choices, row, repeat_stack) {
  type <- as.character(survey$type[row])
  type_parts <- strsplit(type, " ")[[1]]
  base_type <- type_parts[1]

  if (base_type == "select_one" && length(type_parts) >= 2) {
    list_name <- type_parts[2]
    opts <- get_choice_values(choices, list_name)
    return(if (length(opts) == 0) NA else sample(opts, 1))
  }

  if (base_type == "select_multiple" && length(type_parts) >= 2) {
    list_name <- type_parts[2]
    opts <- get_choice_values(choices, list_name)
    if (length(opts) == 0) return(NA)
    n_select <- sample(seq_len(length(opts)), 1)
    return(paste(sample(opts, n_select), collapse = " "))
  }

  if (base_type == "integer") {
    return(sample(0:100, 1))
  }

  if (base_type == "decimal") {
    return(runif(1, min = 0, max = 100))
  }

  if (base_type == "date") {
    return(as.character(Sys.Date() - sample(0:365, 1)))
  }

  if (base_type == "datetime") {
    return(as.character(Sys.time() - sample(0:86400, 1)))
  }

  if (base_type == "time") {
    return(format(Sys.time(), "%H:%M:%S"))
  }

  if (base_type == "text") {
    n_chars <- sample(5:20, 1)
    letters_pool <- c(letters, " ")
    return(paste(sample(letters_pool, n_chars, replace = TRUE), collapse = ""))
  }

  NA
}

get_choice_values <- function(choices, list_name) {
  if (!"list_name" %in% names(choices)) return(character())
  value_col <- if ("name" %in% names(choices)) "name" else if ("value" %in% names(choices)) "value" else NULL
  if (is.null(value_col)) return(character())
  rows <- choices$list_name == list_name
  vals <- choices[[value_col]][rows]
  vals <- vals[!is.na(vals)]
  as.character(vals)
}

get_col_value <- function(survey, col, row) {
  if (!col %in% names(survey)) return(NA_character_)
  survey[[col]][row]
}

make_repeat_name <- function(name, repeat_stack) {
  if (length(repeat_stack) == 0) return(name)
  suffix <- paste0("r", repeat_stack, collapse = "_")
  paste0(name, "__", suffix)
}

eval_xlsform_expr <- function(expr, answers, repeat_stack) {
  expr <- trimws(expr)
  if (!nzchar(expr)) return(TRUE)

  translated <- translate_xlsform_expr(expr)
  get_answer <- function(name) {
    if (name %in% names(answers)) return(answers[[name]])
    alt <- make_repeat_name(name, repeat_stack)
    if (alt %in% names(answers)) return(answers[[alt]])
    NA
  }

  env <- new.env(parent = baseenv())
  env$get_answer <- get_answer
  env$xls_selected <- xls_selected
  env$xls_count_selected <- xls_count_selected
  env$xls_regex <- xls_regex
  env$xls_contains <- xls_contains

  tryCatch(
    eval(parse(text = translated), env),
    error = function(e) {
      warning(sprintf("Failed to evaluate expression '%s': %s", expr, e$message))
      NA
    }
  )
}

translate_xlsform_expr <- function(expr) {
  out <- expr
  out <- gsub("\\$\\{([^}]+)\\}", "get_answer(\"\\1\")", out)
  out <- gsub("(?i)\\band\\b", "&", out, perl = TRUE)
  out <- gsub("(?i)\\bor\\b", "|", out, perl = TRUE)
  out <- gsub("(?i)\\bnot\\b", "!", out, perl = TRUE)
  out <- gsub("(?<![<>=!])=(?!=)", "==", out, perl = TRUE)
  out <- gsub("(?i)\\bselected\\s*\\(", "xls_selected(", out, perl = TRUE)
  out <- gsub("(?i)\\bcount-selected\\s*\\(", "xls_count_selected(", out, perl = TRUE)
  out <- gsub("(?i)\\bregex\\s*\\(", "xls_regex(", out, perl = TRUE)
  out <- gsub("(?i)\\bcontains\\s*\\(", "xls_contains(", out, perl = TRUE)
  out <- gsub("(?i)\\btoday\\s*\\(\\s*\\)", "Sys.Date()", out, perl = TRUE)
  out <- gsub("(?i)\\bstring-length\\s*\\(", "nchar(", out, perl = TRUE)
  out <- gsub("(?i)\\bint\\s*\\(", "as.integer(", out, perl = TRUE)
  out <- gsub("(?i)\\bdecimal\\s*\\(", "as.numeric(", out, perl = TRUE)
  out
}

xls_selected <- function(value, choice) {
  if (length(value) == 0 || is.na(value)) return(FALSE)
  if (length(value) > 1) {
    vals <- as.character(value)
  } else {
    vals <- unlist(strsplit(as.character(value), " "))
  }
  choice %in% vals
}

xls_count_selected <- function(value) {
  if (length(value) == 0 || is.na(value)) return(0L)
  if (length(value) > 1) return(length(value))
  length(unlist(strsplit(as.character(value), " ")))
}

xls_regex <- function(value, pattern) {
  if (length(value) == 0 || is.na(value)) return(FALSE)
  grepl(pattern, as.character(value))
}

xls_contains <- function(value, pattern) {
  if (length(value) == 0 || is.na(value)) return(FALSE)
  grepl(pattern, as.character(value), fixed = TRUE)
}

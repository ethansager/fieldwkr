#' Generate mock data from a SurveyCTO/ODK XLSForm
#'
#' @param path Path to the XLSForm (Excel file).
#' @param n Number of mock submissions to generate.
#' @param seed Optional random seed. With `seed` and `today` both fixed the
#'   output is reproducible.
#' @param max_tries Maximum draws per field when trying to satisfy its
#'   constraint.
#' @param max_repeat Cap on repeat instances. A repeat with a `repeat_count`
#'   uses it, up to this cap; a repeat without one gets 1 to `max_repeat`
#'   instances at random.
#' @param today Date used for `today()`, date fields and timestamps.
#' @param metadata Add SurveyCTO's `SubmissionDate`, `formdef_version` and
#'   `KEY` columns, as in a server export.
#' @details
#' Walks the `survey` sheet the way the form would be filled in. Relevance on
#' fields, groups and repeats is evaluated against earlier answers,
#' constraints are met by redrawing, `calculate` and `calculate_here` fields
#' are computed, and select fields draw only from choices that pass their
#' `choice_filter`. Integer and decimal fields draw within the bounds stated in
#' their constraint (for example `. >= 18 and . <= 99`) when it has any.
#'
#' Expressions use SurveyCTO's expression language, including `if()`,
#' `selected()`, `count-selected()`, `selected-at()`, `choice-label()`,
#' `concat()`, `coalesce()`, date functions such as `format-date-time()`, and
#' repeat functions such as `index()`, `count()`, `sum()`, `join()` and their
#' `-if` forms.
#'
#' Some expressions cannot be simulated, most commonly `pulldata()`, which
#' needs the form's attached data. Relevance that cannot be evaluated is
#' treated as true (the question is asked), a constraint that cannot be
#' evaluated accepts the drawn value, and a calculation that cannot be
#' evaluated is left missing. Each such expression is reported once, in a
#' single warning and in `attr(result, "expression_issues")`.
#'
#' Notes are not output, since they store no data. Repeats are flattened into
#' columns with suffixes like `name__r1` or `name__r1_r2`.
#'
#' This is intended for testing pipelines, not for generating realistic survey
#' distributions.
#' @return A data frame with one row per mock submission.
#' @export
dummy_dat <- function(
  path,
  n = 1,
  seed = NULL,
  max_tries = 50,
  max_repeat = 3,
  today = Sys.Date(),
  metadata = TRUE
) {
  if (!is.null(seed)) {
    set.seed(seed)
  }

  ctx <- xls_prepare_form(
    path,
    max_tries = max_tries,
    max_repeat = max_repeat,
    today = as.Date(today)
  )
  records <- lapply(seq_len(n), function(i) xls_simulate_record(ctx))

  out <- xls_records_to_df(records)
  if (metadata) {
    out <- xls_add_metadata(out, records, ctx)
  }

  issues <- xls_issue_table(ctx)
  if (nrow(issues) > 0) {
    attr(out, "expression_issues") <- issues
    warning(xls_issue_summary(issues), call. = FALSE)
  }
  out
}

# ---------------------------------------------------------------------------
# Reading the form

#' Read the form once and precompute everything the simulation looks up
#' @keywords internal
#' @noRd
xls_prepare_form <- function(path, max_tries, max_repeat, today) {
  require_pkg("openxlsx")
  sheets <- openxlsx::getSheetNames(path)

  survey_sheet <- match_sheet_name(sheets, "survey")
  if (is.null(survey_sheet)) {
    stop("XLSForm has no 'survey' sheet.", call. = FALSE)
  }
  # Read positionally so row numbers in reported issues are spreadsheet rows.
  survey <- read_xlsform_sheet(path, survey_sheet, skip_empty = FALSE)
  if (!all(c("type", "name") %in% names(survey))) {
    stop("The survey sheet needs 'type' and 'name' columns.", call. = FALSE)
  }

  col <- function(...) {
    for (nm in c(...)) {
      if (nm %in% names(survey)) {
        v <- as.character(survey[[nm]])
        v[is.na(v)] <- ""
        return(trimws(v))
      }
    }
    rep("", nrow(survey))
  }
  yes <- function(v) tolower(v) %in% c("yes", "true", "1")

  ctx <- new.env()
  ctx$n <- nrow(survey)
  ctx$type <- normalize_xlsform_type(survey$type)
  ctx$kind <- xls_kind(ctx$type)
  ctx$name <- col("name")
  ctx$relevance <- col("relevance", "relevant")
  ctx$constraint <- col("constraint")
  ctx$calculation <- col("calculation")
  ctx$repeat_count <- col("repeat_count")
  ctx$choice_filter <- col("choice_filter")
  ctx$appearance <- tolower(col("appearance"))
  ctx$read_only <- yes(col("read only", "read.only", "read_only"))
  ctx$disabled <- yes(col("disabled"))
  ctx$list_of <- ifelse(
    ctx$kind %in% c("select_one", "select_multiple", "rank") &
      grepl("^\\S+\\s+\\S+", ctx$type),
    sub("^\\S+\\s+(\\S+).*$", "\\1", ctx$type),
    ""
  )
  ctx$dynamic <- grepl("search(", ctx$appearance, fixed = TRUE)
  # Re-entry checks such as `. = ${phone}` cannot be met by random draws, so
  # the field copies the referenced value instead.
  ctx$equals <- ifelse(
    grepl("^\\.\\s*=\\s*\\$\\{[^}]+\\}$", ctx$constraint),
    sub("^\\.\\s*=\\s*", "", ctx$constraint),
    ""
  )
  ctx$bounds <- lapply(seq_len(ctx$n), function(i) {
    if (ctx$kind[i] %in% c("integer", "decimal")) {
      xls_constraint_bounds(ctx$constraint[i])
    } else {
      list()
    }
  })

  xls_index_structure(ctx)
  ctx$lists <- xls_index_choices(path, sheets)

  settings_sheet <- match_sheet_name(sheets, "settings")
  settings <- if (is.null(settings_sheet)) NULL else read_xlsx_sheet(path, settings_sheet)
  ctx$version <- if (!is.null(settings) && "version" %in% names(settings) && nrow(settings) > 0) {
    as.character(settings$version[1])
  } else {
    NA_character_
  }

  ctx$max_tries <- max_tries
  ctx$max_repeat <- max_repeat
  ctx$today <- today
  ctx$lib <- xls_new_lib()
  ctx$parsed <- new.env(hash = TRUE)
  ctx$issues <- new.env(hash = TRUE)
  ctx
}

#' Row kind: the type keyword, or the whole type for multi-word types
#' @keywords internal
#' @noRd
xls_kind <- function(type) {
  multi <- c(
    "begin group",
    "end group",
    "begin repeat",
    "end repeat",
    "text audit",
    "audio audit",
    "speed violations count",
    "speed violations list",
    "speed violations audit"
  )
  ifelse(type %in% multi, type, sub(" .*$", "", type))
}

#' Match begin/end rows and record which repeats enclose each field
#' @keywords internal
#' @noRd
xls_index_structure <- function(ctx) {
  ctx$block_end <- rep(NA_integer_, ctx$n)
  ctx$field_repeats <- new.env(hash = TRUE)
  ctx$repeat_paths <- new.env(hash = TRUE)
  ctx$field_list <- new.env(hash = TRUE)

  open <- integer()
  repeats <- character()
  for (i in seq_len(ctx$n)) {
    kind <- ctx$kind[i]
    nm <- ctx$name[i]
    if (kind %in% c("begin group", "begin repeat")) {
      open <- c(open, i)
      if (kind == "begin repeat") {
        repeats <- c(repeats, nm)
        assign(nm, repeats, envir = ctx$repeat_paths)
      }
      next
    }
    if (kind %in% c("end group", "end repeat")) {
      if (length(open) > 0) {
        opener <- open[length(open)]
        ctx$block_end[opener] <- i
        open <- open[-length(open)]
        if (ctx$kind[opener] == "begin repeat") {
          repeats <- repeats[-length(repeats)]
        }
      }
      next
    }
    if (nzchar(nm)) {
      assign(nm, repeats, envir = ctx$field_repeats)
      if (nzchar(ctx$list_of[i])) {
        assign(nm, ctx$list_of[i], envir = ctx$field_list)
      }
    }
  }
  invisible(ctx)
}

#' Choice lists keyed by list name, spacer rows dropped
#' @keywords internal
#' @noRd
xls_index_choices <- function(path, sheets) {
  lists <- new.env(hash = TRUE)
  choices_sheet <- match_sheet_name(sheets, "choices")
  if (is.null(choices_sheet)) {
    return(lists)
  }
  choices <- read_xlsform_sheet(path, choices_sheet)
  value_col <- intersect(c("value", "name"), names(choices))[1]
  if (!"list_name" %in% names(choices) || is.na(value_col)) {
    return(lists)
  }
  label_col <- intersect(
    c("label", grep("^label:", names(choices), value = TRUE)),
    names(choices)
  )[1]

  clean <- function(v) {
    v <- as.character(v)
    v[is.na(v)] <- ""
    trimws(v)
  }
  choices[] <- lapply(choices, clean)
  choices$value <- choices[[value_col]]
  choices$label <- if (!is.na(label_col)) choices[[label_col]] else ""
  choices <- choices[choices$list_name != "" & choices$value != "", , drop = FALSE]

  for (ln in unique(choices$list_name)) {
    rows <- choices[choices$list_name == ln, , drop = FALSE]
    rownames(rows) <- NULL
    assign(ln, rows, envir = lists)
  }
  lists
}

#' Bounds such as `. >= 18` or `. <= ${hhsize}` read from a constraint
#' @keywords internal
#' @noRd
xls_constraint_bounds <- function(constraint) {
  if (!nzchar(constraint)) {
    return(list())
  }
  rx <- paste0(
    "(?<![A-Za-z0-9_.}])\\.(?![0-9A-Za-z_.])\\s*(>=|<=|>|<)\\s*",
    "(-?[0-9]+(?:\\.[0-9]+)?|\\$\\{[^}]+\\})"
  )
  hits <- regmatches(constraint, gregexpr(rx, constraint, perl = TRUE))[[1]]
  lapply(hits, function(h) {
    parts <- regmatches(h, regexec(rx, h, perl = TRUE))[[1]]
    list(op = parts[2], rhs = parts[3])
  })
}

# ---------------------------------------------------------------------------
# Simulating one submission

#' @keywords internal
#' @noRd
xls_simulate_record <- function(ctx) {
  state <- new.env()
  state$ans <- new.env(hash = TRUE)
  state$counts <- new.env(hash = TRUE)
  state$keys <- character()
  state$start <- as.POSIXct(format(ctx$today), tz = "UTC") +
    stats::runif(1, 7, 17) * 3600
  state$now <- state$start
  state$duration <- 0
  xls_run_block(ctx, state, 1L, ctx$n, character(), integer())
  state
}

#' Fill rows `from`..`to` in the repeat context given by `names`/`idx`
#' @keywords internal
#' @noRd
xls_run_block <- function(ctx, state, from, to, names, idx) {
  i <- from
  while (i <= to) {
    kind <- ctx$kind[i]
    if (kind %in% c("begin group", "begin repeat")) {
      end <- ctx$block_end[i]
      if (is.na(end)) {
        end <- to + 1L
      }
      if (ctx$disabled[i]) {
        i <- end + 1L
        next
      }
      relevant <- xls_row_relevant(ctx, state, i, names, idx)
      if (kind == "begin group") {
        if (relevant) {
          xls_run_block(ctx, state, i + 1L, end - 1L, names, idx)
        } else {
          xls_blank_block(ctx, state, i + 1L, end - 1L, idx)
        }
      } else {
        loops <- if (relevant) xls_repeat_loops(ctx, state, i, names, idx) else 0L
        assign(make_repeat_name(ctx$name[i], idx), loops, envir = state$counts)
        for (loop in seq_len(loops)) {
          xls_run_block(
            ctx,
            state,
            i + 1L,
            end - 1L,
            c(names, ctx$name[i]),
            c(idx, loop)
          )
        }
      }
      i <- end + 1L
      next
    }
    if (!ctx$disabled[i]) {
      xls_run_field(ctx, state, i, names, idx)
    }
    i <- i + 1L
  }
  invisible(state)
}

#' Set every field of a non-relevant group to missing
#' @keywords internal
#' @noRd
xls_blank_block <- function(ctx, state, from, to, idx) {
  i <- from
  while (i <= to) {
    kind <- ctx$kind[i]
    if (kind == "begin repeat") {
      assign(make_repeat_name(ctx$name[i], idx), 0L, envir = state$counts)
      end <- ctx$block_end[i]
      i <- (if (is.na(end)) to else end) + 1L
      next
    }
    if (xls_stores_data(kind) && nzchar(ctx$name[i]) && !ctx$disabled[i]) {
      xls_store(state, make_repeat_name(ctx$name[i], idx), NA)
    }
    i <- i + 1L
  }
}

#' @keywords internal
#' @noRd
xls_stores_data <- function(kind) {
  !kind %in% c("", "note", "begin group", "end group", "begin repeat", "end repeat")
}

#' @keywords internal
#' @noRd
xls_run_field <- function(ctx, state, i, names, idx) {
  kind <- ctx$kind[i]
  if (!xls_stores_data(kind) || !nzchar(ctx$name[i])) {
    return(invisible())
  }
  key <- make_repeat_name(ctx$name[i], idx)

  step <- stats::runif(1, 3, 30)
  state$duration <- state$duration + step
  state$now <- state$now + step

  if (!xls_row_relevant(ctx, state, i, names, idx)) {
    xls_store(state, key, NA)
    return(invisible())
  }

  computed <- kind %in% c("calculate", "calculate_here") ||
    (ctx$read_only[i] && nzchar(ctx$calculation[i]))
  if (computed) {
    xls_store(state, key, xls_calculate(ctx, state, i, names, idx))
    return(invisible())
  }

  if (nzchar(ctx$equals[i])) {
    value <- xls_quiet_eval(ctx, state, ctx$equals[i], names, idx)
    if (!xls_is_empty(value)) {
      xls_store(state, key, value)
      return(invisible())
    }
  }

  value <- xls_draw(ctx, state, i, names, idx)
  if (nzchar(ctx$constraint[i]) && !(length(value) == 1 && is.na(value))) {
    value <- xls_satisfy(ctx, state, i, names, idx, key, value)
  }
  xls_store(state, key, value)
}

#' @keywords internal
#' @noRd
xls_store <- function(state, key, value) {
  if (!exists(key, envir = state$ans, inherits = FALSE)) {
    state$keys <- c(state$keys, key)
  }
  assign(key, value, envir = state$ans)
}

#' @keywords internal
#' @noRd
xls_where <- function(ctx, i, column) {
  list(row = i, field = ctx$name[i], column = column)
}

#' Relevance, with an unevaluable expression treated as relevant
#' @keywords internal
#' @noRd
xls_row_relevant <- function(ctx, state, i, names, idx) {
  rel <- ctx$relevance[i]
  if (!nzchar(rel)) {
    return(TRUE)
  }
  r <- xls_eval(ctx, state, rel, names, idx, xls_where(ctx, i, "relevance"))
  if (inherits(r, "xls_unevaluable")) {
    return(TRUE)
  }
  isTRUE(xls_bool(r)[1])
}

#' @keywords internal
#' @noRd
xls_repeat_loops <- function(ctx, state, i, names, idx) {
  rc <- ctx$repeat_count[i]
  if (nzchar(rc)) {
    r <- xls_eval(ctx, state, rc, names, idx, xls_where(ctx, i, "repeat_count"))
    if (!inherits(r, "xls_unevaluable")) {
      k <- suppressWarnings(as.integer(xls_num(r)[1]))
      if (!is.na(k) && k >= 0) {
        return(min(k, ctx$max_repeat))
      }
    }
  }
  if (ctx$max_repeat < 1) 0L else sample.int(ctx$max_repeat, 1)
}

#' @keywords internal
#' @noRd
xls_calculate <- function(ctx, state, i, names, idx) {
  expr <- ctx$calculation[i]
  if (!nzchar(expr)) {
    return(NA)
  }
  where <- xls_where(ctx, i, "calculation")
  v <- xls_eval(ctx, state, expr, names, idx, where)
  if (inherits(v, "xls_unevaluable")) {
    return(NA)
  }
  if (length(v) != 1) {
    xls_note_issue(ctx, where, expr, "evaluated to more than one value; left missing")
    return(NA)
  }
  v
}

#' Redraw until the constraint holds, or give up after max_tries
#' @keywords internal
#' @noRd
xls_satisfy <- function(ctx, state, i, names, idx, key, value) {
  where <- xls_where(ctx, i, "constraint")
  for (try in seq_len(ctx$max_tries)) {
    # Stored so the constraint can also refer to this field by name.
    xls_store(state, key, value)
    r <- xls_eval(
      ctx,
      state,
      ctx$constraint[i],
      names,
      idx,
      where,
      current = value
    )
    if (inherits(r, "xls_unevaluable") || isTRUE(xls_bool(r)[1])) {
      return(value)
    }
    value <- xls_draw(ctx, state, i, names, idx)
  }
  xls_note_issue(
    ctx,
    where,
    ctx$constraint[i],
    sprintf("no draw met the constraint in %d tries; left missing", ctx$max_tries)
  )
  NA
}

# ---------------------------------------------------------------------------
# Drawing values

#' @keywords internal
#' @noRd
xls_draw <- function(ctx, state, i, names, idx) {
  kind <- ctx$kind[i]
  switch(
    kind,
    select_one = ,
    select_multiple = ,
    rank = xls_draw_select(ctx, state, i, names, idx),
    integer = xls_draw_number(ctx, state, i, names, idx, integer = TRUE),
    decimal = xls_draw_number(ctx, state, i, names, idx, integer = FALSE),
    text = if (grepl("numbers", ctx$appearance[i], fixed = TRUE)) {
      paste(sample(0:9, 10, replace = TRUE), collapse = "")
    } else {
      xls_draw_text()
    },
    date = ctx$today - sample.int(366, 1) + 1L,
    datetime = state$now - stats::runif(1, 0, 30 * 86400),
    time = sprintf("%02d:%02d:00", sample(6:19, 1), sample(0:59, 1)),
    start = state$start,
    end = state$start + stats::runif(1, 20, 90) * 60,
    today = ctx$today,
    deviceid = paste(sample(c(0:9, letters[1:6]), 16, replace = TRUE), collapse = ""),
    username = "enumerator",
    enumerator = "1",
    acknowledge = "OK",
    barcode = paste(sample(0:9, 12, replace = TRUE), collapse = ""),
    geopoint = xls_draw_geopoint(),
    geotrace = ,
    geoshape = paste(replicate(4, xls_draw_geopoint()), collapse = "; "),
    image = paste0(ctx$name[i], "_", sample.int(1e9, 1), ".jpg"),
    audio = paste0(ctx$name[i], "_", sample.int(1e9, 1), ".m4a"),
    video = paste0(ctx$name[i], "_", sample.int(1e9, 1), ".mp4"),
    file = paste0(ctx$name[i], "_", sample.int(1e9, 1), ".bin"),
    comments = paste0("media/Comments-", xls_uuid(), ".csv"),
    `text audit` = paste0("media/TA_", xls_uuid(), ".csv"),
    `audio audit` = paste0("media/AA_", xls_uuid(), ".m4a"),
    sensor_stream = paste0("media/SS_", xls_uuid(), ".csv"),
    sensor_statistic = round(stats::runif(1, 0, 100), 2),
    `speed violations count` = 0L,
    # subscriberid, simserial, phonenumber, caseid, hidden and the
    # speed-violation list/audit are blank on current devices
    NA
  )
}

#' @keywords internal
#' @noRd
xls_draw_text <- function() {
  n_chars <- sample(5:20, 1)
  paste(sample(c(letters, " "), n_chars, replace = TRUE), collapse = "")
}

#' @keywords internal
#' @noRd
xls_draw_geopoint <- function() {
  sprintf(
    "%.6f %.6f %.1f %.1f",
    stats::runif(1, -60, 60),
    stats::runif(1, -180, 180),
    stats::runif(1, 0, 2000),
    stats::runif(1, 3, 20)
  )
}

#' @keywords internal
#' @noRd
xls_draw_select <- function(ctx, state, i, names, idx) {
  if (ctx$dynamic[i]) {
    xls_note_issue(
      ctx,
      xls_where(ctx, i, "appearance"),
      ctx$appearance[i],
      "choices loaded with search() need the form's attached data"
    )
    return(NA)
  }
  rows <- get0(ctx$list_of[i], envir = ctx$lists, inherits = FALSE)
  if (is.null(rows) || nrow(rows) == 0) {
    return(NA)
  }
  vals <- rows$value

  if (nzchar(ctx$choice_filter[i])) {
    where <- xls_where(ctx, i, "choice_filter")
    keep <- xls_eval(
      ctx,
      state,
      ctx$choice_filter[i],
      names,
      idx,
      where,
      bind = rows
    )
    if (!inherits(keep, "xls_unevaluable")) {
      keep <- xls_bool(keep)
      if (length(keep) == 1) {
        keep <- rep(keep, length(vals))
      }
      if (length(keep) == length(vals)) {
        vals <- vals[keep]
      } else {
        xls_note_issue(
          ctx,
          where,
          ctx$choice_filter[i],
          "did not give one result per choice; filter ignored"
        )
      }
    }
  }

  if (length(vals) == 0) {
    return(NA)
  }
  switch(
    ctx$kind[i],
    select_one = vals[sample.int(length(vals), 1)],
    select_multiple = {
      k <- sample.int(length(vals), 1)
      paste(vals[sort(sample.int(length(vals), k))], collapse = " ")
    },
    rank = paste(vals[sample.int(length(vals))], collapse = " ")
  )
}

#' Draw a number within the constraint's stated bounds, if any
#' @keywords internal
#' @noRd
xls_draw_number <- function(ctx, state, i, names, idx, integer) {
  lo <- NA_real_
  hi <- NA_real_
  step <- if (integer) 1 else 0.01
  for (b in ctx$bounds[[i]]) {
    v <- xls_bound_value(ctx, state, b$rhs, names, idx)
    if (is.na(v)) next
    if (b$op %in% c(">=", ">")) {
      lo <- max(lo, if (b$op == ">") v + step else v, na.rm = TRUE)
    } else {
      hi <- min(hi, if (b$op == "<") v - step else v, na.rm = TRUE)
    }
  }
  if (is.na(lo) && is.na(hi)) {
    lo <- 0
    hi <- 100
  } else if (is.na(lo)) {
    lo <- min(0, hi - 100)
  } else if (is.na(hi)) {
    hi <- lo + 100
  }
  if (lo > hi) {
    lo <- 0
    hi <- 100
  }
  if (integer) {
    lo <- ceiling(lo)
    hi <- floor(hi)
    if (lo > hi) {
      return(NA)
    }
    return(as.integer(lo + floor(stats::runif(1) * (hi - lo + 1))))
  }
  round(stats::runif(1, lo, hi), 2)
}

#' Value of a bound: a literal number, or a field reference
#' @keywords internal
#' @noRd
xls_bound_value <- function(ctx, state, rhs, names, idx) {
  if (!startsWith(rhs, "${")) {
    return(as.numeric(rhs))
  }
  suppressWarnings(xls_num(xls_quiet_eval(ctx, state, rhs, names, idx))[1])
}

#' Evaluate a helper expression without recording issues; NA on failure
#' @keywords internal
#' @noRd
xls_quiet_eval <- function(ctx, state, expr, names, idx) {
  parsed <- xls_parse(ctx, expr)
  if (inherits(parsed, "condition")) {
    return(NA)
  }
  tryCatch(
    eval(parsed, xls_frame(ctx, state, names, idx)),
    error = function(e) NA
  )
}

# ---------------------------------------------------------------------------
# Output

#' Repeat instance key, e.g. name__r1_r2 for instance 2 inside instance 1
#' @keywords internal
#' @noRd
make_repeat_name <- function(name, repeat_stack) {
  if (length(repeat_stack) == 0) {
    return(name)
  }
  suffix <- paste0("r", repeat_stack, collapse = "_")
  paste0(name, "__", suffix)
}

#' @keywords internal
#' @noRd
xls_records_to_df <- function(records) {
  keys <- unique(unlist(lapply(records, function(s) s$keys), use.names = FALSE))
  if (length(keys) == 0) {
    return(data.frame(row.names = seq_along(records)))
  }
  cols <- lapply(keys, function(k) {
    xls_collapse(lapply(records, function(s) {
      xls_output_value(get0(k, envir = s$ans, inherits = FALSE, ifnotfound = NA))
    }))
  })
  names(cols) <- keys
  as.data.frame(cols, stringsAsFactors = FALSE, check.names = FALSE)
}

#' @keywords internal
#' @noRd
xls_output_value <- function(v) {
  if (length(v) != 1) {
    return(NA)
  }
  if (inherits(v, c("Date", "POSIXct"))) {
    return(xls_chr(v))
  }
  if (is.character(v) && !nzchar(v)) {
    return(NA)
  }
  v
}

#' One column from per-record values: numeric, logical, or character
#' @keywords internal
#' @noRd
xls_collapse <- function(vals) {
  is_na <- vapply(vals, function(v) is.na(v), logical(1))
  present <- vals[!is_na]
  if (all(vapply(present, is.numeric, logical(1))) ||
      all(vapply(present, is.logical, logical(1)))) {
    return(unlist(vals, use.names = FALSE))
  }
  vapply(
    vals,
    function(v) if (is.na(v)) NA_character_ else xls_chr(v),
    character(1)
  )
}

#' @keywords internal
#' @noRd
xls_add_metadata <- function(out, records, ctx) {
  submitted <- vapply(
    records,
    function(s) xls_chr(s$now + stats::runif(1, 60, 3 * 86400)),
    character(1)
  )
  keys <- vapply(records, function(s) paste0("uuid:", xls_uuid()), character(1))
  meta <- data.frame(SubmissionDate = submitted, stringsAsFactors = FALSE)
  out <- if (ncol(out) > 0) cbind(meta, out) else meta
  out$formdef_version <- rep(ctx$version, nrow(out))
  out$KEY <- keys
  out
}

#' @keywords internal
#' @noRd
xls_issue_summary <- function(issues) {
  counts <- sort(table(issues$problem), decreasing = TRUE)
  lines <- sprintf("  %d x %s", as.integer(counts), names(counts))
  paste0(
    nrow(issues),
    " expression(s) in the form could not be fully simulated. ",
    "Rows and details are in attr(result, \"expression_issues\").\n",
    paste(utils::head(lines, 8), collapse = "\n")
  )
}

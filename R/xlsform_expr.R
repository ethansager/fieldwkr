# Evaluating SurveyCTO / XLSForm expressions in R.
#
# An expression is translated to R code once (and cached), then evaluated in a
# fresh environment whose parent is a function library. The library supplies
# get_answer(), one xls_fn_<name>() per supported SurveyCTO function, and
# XPath-style comparison, arithmetic and boolean operators. Semantics follow
# SurveyCTO's expression reference: empty answers compare as '' (never NA),
# `=` compares numerically when either side is a number, `and`/`or`
# short-circuit, and repeated fields referenced from outside their repeat
# resolve to all instances so that sum(), join() and friends work.

#' Translate a SurveyCTO expression into R code
#'
#' The expression is split into string literals and code, and only the code is
#' rewritten: `${field}` becomes `get_answer("field")`, every function call
#' `name(` becomes `xls_fn_<name>(` with `-` and `:` mapped to `_`, and the
#' operators `and`, `or`, `div`, `mod` and `=` become `&`, `|`, `/`, `%%` and
#' `==`. Literals are re-emitted with deparse() so regex patterns keep their
#' backslashes.
#' @keywords internal
#' @noRd
translate_xlsform_expr <- function(expr) {
  pieces <- regmatches(
    expr,
    gregexpr("'[^']*'|\"[^\"]*\"", expr),
    invert = NA
  )[[1]]
  out <- vapply(
    seq_along(pieces),
    function(k) {
      p <- pieces[k]
      if (k %% 2 == 0) {
        paste(deparse(substr(p, 2, nchar(p) - 1)), collapse = "")
      } else {
        xls_translate_code(p)
      }
    },
    character(1)
  )
  paste(out, collapse = "")
}

#' @keywords internal
#' @noRd
xls_translate_code <- function(code) {
  # Park field references behind placeholders so no later rewrite can touch a
  # field name (a field called "order" must not become "|der").
  refs <- regmatches(code, gregexpr("\\$\\{[^}]*\\}", code))[[1]]
  ref_names <- trimws(substr(refs, 3, nchar(refs) - 1))
  for (k in seq_along(refs)) {
    code <- sub(refs[k], sprintf(" XLSREF%dX ", k), code, fixed = TRUE)
  }

  call_rx <- "(?<![A-Za-z0-9_.])[A-Za-z][A-Za-z0-9_:-]*\\s*\\("
  m <- gregexpr(call_rx, code, perl = TRUE)
  calls <- regmatches(code, m)[[1]]
  if (length(calls) > 0) {
    fn <- tolower(sub("\\s*\\($", "", calls, perl = TRUE))
    regmatches(code, m) <- list(ifelse(
      fn %in% c("and", "or", "div", "mod"),
      calls,
      paste0("xls_fn_", gsub("[-:]", "_", fn), "(")
    ))
  }

  code <- gsub("\\band\\b", "&", code, perl = TRUE)
  code <- gsub("\\bor\\b", "|", code, perl = TRUE)
  code <- gsub("\\bdiv\\b", "/", code, perl = TRUE)
  code <- gsub("\\bmod\\b", "%%", code, perl = TRUE)
  code <- gsub("(?<![<>=!])=(?!=)", "==", code, perl = TRUE)

  for (k in seq_along(refs)) {
    code <- sub(
      sprintf(" XLSREF%dX ", k),
      sprintf("get_answer(%s)", deparse(ref_names[k])),
      code,
      fixed = TRUE
    )
  }
  code
}

# ---------------------------------------------------------------------------
# Evaluation

#' Evaluate one expression for one record
#'
#' `names` and `idx` are the repeat context: the enclosing repeat names and the
#' current instance number of each. `current` binds `.` (constraints), and
#' `bind` binds extra symbols, used for choice-sheet columns in choice_filter.
#' Returns the value, or an object of class `xls_unevaluable` when the
#' expression cannot be evaluated; the reason is recorded in `ctx$issues`.
#' @keywords internal
#' @noRd
xls_eval <- function(
  ctx,
  state,
  expr,
  names,
  idx,
  where,
  current = NULL,
  bind = NULL
) {
  parsed <- xls_parse(ctx, expr)
  if (inherits(parsed, "condition")) {
    xls_note_issue(ctx, where, expr, "could not be parsed")
    return(xls_unevaluable())
  }

  env <- xls_frame(ctx, state, names, idx)
  if (!is.null(current)) {
    assign(".", current, envir = env)
  }
  for (nm in names(bind)) {
    assign(nm, bind[[nm]], envir = env)
  }

  tryCatch(
    eval(parsed, env),
    xls_unresolvable = function(e) {
      xls_note_issue(ctx, where, expr, conditionMessage(e))
      xls_unevaluable()
    },
    error = function(e) {
      xls_note_issue(ctx, where, expr, xls_describe_error(e))
      xls_unevaluable()
    }
  )
}

#' @keywords internal
#' @noRd
xls_parse <- function(ctx, expr) {
  hit <- get0(expr, envir = ctx$parsed, inherits = FALSE)
  if (!is.null(hit)) {
    return(hit)
  }
  parsed <- tryCatch(
    {
      p <- parse(text = translate_xlsform_expr(expr), keep.source = FALSE)
      if (length(p) != 1) stop("expected a single expression")
      p[[1]]
    },
    error = function(e) e
  )
  assign(expr, parsed, envir = ctx$parsed)
  parsed
}

#' @keywords internal
#' @noRd
xls_frame <- function(ctx, state, names, idx) {
  env <- new.env(parent = ctx$lib)
  env$.xls <- list(ctx = ctx, state = state, names = names, idx = idx)
  env
}

#' @keywords internal
#' @noRd
xls_unevaluable <- function() {
  structure(list(), class = "xls_unevaluable")
}

#' Signal that a function cannot be simulated (pulldata() and the like)
#' @keywords internal
#' @noRd
xls_unresolvable <- function(message) {
  stop(structure(
    class = c("xls_unresolvable", "error", "condition"),
    list(message = message, call = NULL)
  ))
}

#' @keywords internal
#' @noRd
xls_describe_error <- function(e) {
  msg <- conditionMessage(e)
  fn <- regmatches(msg, regexpr("xls_fn_[A-Za-z0-9_]+", msg))
  if (length(fn) == 1 && grepl("could not find function", msg, fixed = TRUE)) {
    return(sprintf(
      "unsupported function %s()",
      gsub("_", "-", sub("^xls_fn_", "", fn))
    ))
  }
  msg
}

#' Record a problem once per (row, column), counting repeats
#' @keywords internal
#' @noRd
xls_note_issue <- function(ctx, where, expr, problem) {
  key <- paste(where$row, where$column, sep = "|")
  hit <- get0(key, envir = ctx$issues, inherits = FALSE)
  if (is.null(hit)) {
    hit <- list(
      row = where$row + 1L,
      field = where$field,
      column = where$column,
      expression = expr,
      problem = problem,
      n = 0L
    )
  }
  hit$n <- hit$n + 1L
  assign(key, hit, envir = ctx$issues)
}

#' @keywords internal
#' @noRd
xls_issue_table <- function(ctx) {
  keys <- ls(ctx$issues)
  if (length(keys) == 0) {
    return(data.frame(
      row = integer(),
      field = character(),
      column = character(),
      expression = character(),
      problem = character(),
      n = integer(),
      stringsAsFactors = FALSE
    ))
  }
  rows <- lapply(keys, function(k) {
    as.data.frame(get(k, envir = ctx$issues), stringsAsFactors = FALSE)
  })
  out <- do.call(rbind, rows)
  out <- out[order(out$row, out$column), , drop = FALSE]
  rownames(out) <- NULL
  out
}

# ---------------------------------------------------------------------------
# Answer lookup across repeats

#' Resolve a field reference in the current repeat context
#'
#' A field in the current repeat instance (or outside any repeat) resolves to a
#' single value. A field inside a repeat the expression is not in resolves to
#' all of its instances, which is what SurveyCTO's aggregate functions expect.
#' A repeat group's own name resolves to one element per instance, so count()
#' works on it.
#' @keywords internal
#' @noRd
xls_lookup <- function(x, name) {
  ctx <- x$ctx
  rpath <- get0(name, envir = ctx$repeat_paths, inherits = FALSE)
  if (!is.null(rpath)) {
    return(seq_along(xls_instances(x, rpath, own_level = TRUE)))
  }
  path <- get0(name, envir = ctx$field_repeats, inherits = FALSE)
  if (is.null(path)) {
    return(NA)
  }
  k <- xls_common_prefix(path, x$names)
  if (length(path) == k) {
    return(xls_answer(x$state, name, x$idx[seq_len(k)]))
  }
  vals <- lapply(xls_instances(x, path), function(i) {
    xls_answer(x$state, name, i)
  })
  if (length(vals) == 0) {
    return(character())
  }
  do.call(c, vals)
}

#' @keywords internal
#' @noRd
xls_answer <- function(state, name, idx) {
  get0(
    make_repeat_name(name, idx),
    envir = state$ans,
    inherits = FALSE,
    ifnotfound = NA
  )
}

#' @keywords internal
#' @noRd
xls_common_prefix <- function(a, b) {
  n <- min(length(a), length(b))
  k <- 0L
  while (k < n && identical(a[k + 1L], b[k + 1L])) {
    k <- k + 1L
  }
  k
}

#' Index vectors of every instance of the innermost repeat in `path`
#'
#' Enumerates below the part of `path` shared with the current context. With
#' `own_level = TRUE` (a repeat group referenced by name) the enumeration
#' happens at the group's own level even from inside one of its instances.
#' @keywords internal
#' @noRd
xls_instances <- function(x, path, own_level = FALSE) {
  k <- xls_common_prefix(path, x$names)
  if (own_level) {
    k <- min(k, length(path) - 1L)
  }
  prefixes <- list(x$idx[seq_len(k)])
  if (length(path) > k) {
    for (d in seq(k + 1L, length(path))) {
      nxt <- list()
      for (p in prefixes) {
        n <- get0(
          make_repeat_name(path[d], p),
          envir = x$state$counts,
          inherits = FALSE,
          ifnotfound = 0L
        )
        for (i in seq_len(n)) {
          nxt[[length(nxt) + 1L]] <- c(p, i)
        }
      }
      prefixes <- nxt
    }
  }
  prefixes
}

#' Frames for evaluating an expression once per instance of a repeat
#' @keywords internal
#' @noRd
xls_instance_frames <- function(x, name) {
  rpath <- get0(name, envir = x$ctx$repeat_paths, inherits = FALSE)
  path <- if (!is.null(rpath)) {
    rpath
  } else {
    get0(name, envir = x$ctx$field_repeats, inherits = FALSE)
  }
  if (is.null(path) || length(path) == 0) {
    stop(sprintf("'%s' is not in a repeat group", name), call. = FALSE)
  }
  lapply(xls_instances(x, path, own_level = !is.null(rpath)), function(i) {
    xls_frame(x$ctx, x$state, path, i)
  })
}

#' Field name from an unevaluated argument such as get_answer("age")
#' @keywords internal
#' @noRd
xls_ref_name <- function(arg) {
  if (is.call(arg) && identical(arg[[1]], as.name("get_answer"))) {
    return(as.character(arg[[2]]))
  }
  if (is.character(arg) && length(arg) == 1) {
    return(gsub("^\\$\\{|\\}$", "", arg))
  }
  stop("expected a field reference like ${field}", call. = FALSE)
}

# ---------------------------------------------------------------------------
# Value coercion (XPath semantics)

#' @keywords internal
#' @noRd
xls_is_empty <- function(x) {
  length(x) == 0 || (length(x) == 1 && (is.na(x) || identical(as.character(x), "")))
}

#' @keywords internal
#' @noRd
xls_num <- function(x) {
  if (inherits(x, "POSIXct")) {
    return(as.numeric(x) / 86400)
  }
  if (inherits(x, "Date") || is.numeric(x) || is.logical(x)) {
    return(as.numeric(x))
  }
  suppressWarnings(as.numeric(as.character(x)))
}

#' @keywords internal
#' @noRd
xls_is_num <- function(x) {
  is.numeric(x) || is.logical(x) || inherits(x, c("Date", "POSIXct"))
}

#' @keywords internal
#' @noRd
xls_bool <- function(x) {
  if (is.logical(x)) {
    x[is.na(x)] <- FALSE
    return(x)
  }
  if (xls_is_num(x)) {
    r <- xls_num(x) != 0
    r[is.na(r)] <- FALSE
    return(r)
  }
  r <- nzchar(as.character(x))
  r[is.na(x)] <- FALSE
  r
}

#' @keywords internal
#' @noRd
xls_chr <- function(x) {
  if (length(x) == 0) {
    return("")
  }
  if (inherits(x, "Date")) {
    return(format(x, "%Y-%m-%d"))
  }
  if (inherits(x, "POSIXct")) {
    return(format(x, "%Y-%m-%dT%H:%M:%S"))
  }
  if (is.double(x)) {
    return(vapply(
      x,
      function(v) {
        if (is.na(v)) "" else format(v, scientific = FALSE, digits = 15, trim = TRUE)
      },
      character(1)
    ))
  }
  out <- as.character(x)
  out[is.na(x)] <- ""
  out
}

#' @keywords internal
#' @noRd
xls_date <- function(x) {
  if (inherits(x, "Date")) {
    return(x)
  }
  if (inherits(x, "POSIXct")) {
    return(as.Date(x))
  }
  if (is.numeric(x)) {
    return(as.Date(x, origin = "1970-01-01"))
  }
  s <- xls_chr(x)
  if (!nzchar(s)) {
    return(as.Date(NA))
  }
  out <- as.Date(substr(s, 1, 10), format = "%Y-%m-%d")
  if (is.na(out)) {
    stop(sprintf("cannot convert '%s' to a date", s), call. = FALSE)
  }
  out
}

#' @keywords internal
#' @noRd
xls_datetime <- function(x) {
  if (inherits(x, "POSIXct")) {
    return(x)
  }
  if (inherits(x, "Date")) {
    return(as.POSIXct(format(x), tz = "UTC"))
  }
  s <- sub("T", " ", xls_chr(x), fixed = TRUE)
  out <- as.POSIXct(substr(s, 1, 19), tz = "UTC", format = "%Y-%m-%d %H:%M:%S")
  if (is.na(out)) {
    out <- as.POSIXct(substr(s, 1, 10), tz = "UTC", format = "%Y-%m-%d")
  }
  if (is.na(out)) {
    stop(sprintf("cannot convert '%s' to a date-time", s), call. = FALSE)
  }
  out
}

#' @keywords internal
#' @noRd
xls_split_list <- function(sep, s) {
  s <- xls_chr(s)
  if (!nzchar(s)) {
    return(character())
  }
  # Appending the separator keeps a trailing empty item, which SurveyCTO counts.
  strsplit(paste0(s, sep), sep, fixed = TRUE)[[1]]
}

#' @keywords internal
#' @noRd
xls_selections <- function(value) {
  s <- xls_chr(value)
  if (length(s) != 1 || !nzchar(s)) {
    return(character())
  }
  strsplit(trimws(s), "\\s+")[[1]]
}

# ---------------------------------------------------------------------------
# Operators installed in the evaluation library

#' @keywords internal
#' @noRd
xls_op_eq <- function(e1, e2) {
  r <- if (xls_is_num(e1) || xls_is_num(e2)) {
    xls_num(e1) == xls_num(e2)
  } else {
    xls_chr(e1) == xls_chr(e2)
  }
  r[is.na(r)] <- FALSE
  r
}

#' @keywords internal
#' @noRd
xls_op_ne <- function(e1, e2) {
  !xls_op_eq(e1, e2)
}

#' @keywords internal
#' @noRd
xls_op_rel <- function(op) {
  force(op)
  function(e1, e2) {
    r <- op(xls_num(e1), xls_num(e2))
    r[is.na(r)] <- FALSE
    r
  }
}

#' @keywords internal
#' @noRd
xls_op_arith <- function(op, date_aware = FALSE) {
  force(op)
  function(e1, e2) {
    if (missing(e2)) {
      return(op(xls_num(e1)))
    }
    r <- op(xls_num(e1), xls_num(e2))
    if (date_aware) {
      d1 <- inherits(e1, "Date")
      d2 <- inherits(e2, "Date")
      # date +/- days stays a date; date - date is a number of days
      if (xor(d1, d2)) {
        r <- as.Date(r, origin = "1970-01-01")
      }
    }
    r
  }
}

#' @keywords internal
#' @noRd
xls_op_and <- function(e1, e2) {
  a <- xls_bool(e1)
  if (length(a) == 1 && !a) {
    return(FALSE)
  }
  a & xls_bool(e2)
}

#' @keywords internal
#' @noRd
xls_op_or <- function(e1, e2) {
  a <- xls_bool(e1)
  if (length(a) == 1 && a) {
    return(TRUE)
  }
  a | xls_bool(e2)
}

#' @keywords internal
#' @noRd
xls_get_answer <- function(name) {
  x <- get(".xls", envir = parent.frame())
  v <- xls_lookup(x, name)
  if (length(v) == 1 && is.na(v)) "" else v
}

#' Build the function library that expressions are evaluated against
#' @keywords internal
#' @noRd
xls_new_lib <- function() {
  lib <- new.env(parent = baseenv())
  home <- environment(xls_new_lib)
  for (f in ls(home, pattern = "^xls_fn_")) {
    assign(f, get(f, envir = home), envir = lib)
  }
  lib$get_answer <- xls_get_answer
  lib$`==` <- xls_op_eq
  lib$`!=` <- xls_op_ne
  lib$`<` <- xls_op_rel(base::`<`)
  lib$`>` <- xls_op_rel(base::`>`)
  lib$`<=` <- xls_op_rel(base::`<=`)
  lib$`>=` <- xls_op_rel(base::`>=`)
  lib$`+` <- xls_op_arith(base::`+`, date_aware = TRUE)
  lib$`-` <- xls_op_arith(base::`-`, date_aware = TRUE)
  lib$`*` <- xls_op_arith(base::`*`)
  lib$`/` <- xls_op_arith(base::`/`)
  lib$`%%` <- xls_op_arith(base::`%%`)
  lib$`&` <- xls_op_and
  lib$`|` <- xls_op_or
  lib
}

#' Context of the expression currently being evaluated
#'
#' SurveyCTO functions are called from the evaluation frame built by
#' xls_frame(), so `.xls` is found two frames up from here.
#' @keywords internal
#' @noRd
xls_here <- function() {
  get(".xls", envir = parent.frame(2))
}

# ---------------------------------------------------------------------------
# SurveyCTO functions. Each is called as xls_fn_<name with - and : as _>.
# Semantics follow SurveyCTO's expression reference: selected-at(), item-at()
# and substr() count from 0, index() from 1.

# Logic
xls_fn_if <- function(condition, if_true, if_false) {
  cond <- xls_bool(condition)
  if (length(cond) == 1) {
    if (cond) if_true else if_false
  } else {
    ifelse(cond, if_true, if_false)
  }
}
xls_fn_not <- function(x) !xls_bool(x)
xls_fn_true <- function() TRUE
xls_fn_false <- function() FALSE
xls_fn_boolean <- function(x) xls_bool(x)
xls_fn_once <- function(x) x
xls_fn_empty <- function(x) xls_is_empty(x)
xls_fn_coalesce <- function(...) {
  for (v in list(...)) {
    if (!xls_is_empty(v)) {
      return(v)
    }
  }
  ""
}
xls_fn_relevant <- function(field) {
  x <- xls_here()
  !is.na(xls_lookup(x, xls_ref_name(substitute(field))))[1]
}

# Select fields and choices
xls_fn_selected <- function(field, value) {
  xls_chr(value) %in% xls_selections(field)
}
xls_fn_count_selected <- function(field) length(xls_selections(field))
xls_fn_selected_at <- function(field, n) {
  items <- xls_selections(field)
  i <- as.integer(xls_num(n)) + 1L
  if (is.na(i) || i < 1L || i > length(items)) "" else items[i]
}
xls_fn_choice_label <- function(field, value) {
  x <- xls_here()
  xls_choice_labels(x$ctx, xls_ref_name(substitute(field)), value)
}
xls_fn_jr_choice_name <- function(value, field) {
  x <- xls_here()
  xls_choice_labels(x$ctx, xls_ref_name(xls_chr(field)), value)
}

# Strings
xls_fn_concat <- function(...) {
  paste0(vapply(list(...), function(v) paste(xls_chr(v), collapse = " "), ""), collapse = "")
}
xls_fn_string <- function(x) xls_chr(x)
xls_fn_string_length <- function(x) nchar(xls_chr(x))
xls_fn_substr <- function(x, start, end) {
  s <- xls_chr(x)
  from <- as.integer(xls_num(start)) + 1L
  to <- if (missing(end)) nchar(s) else as.integer(xls_num(end))
  substr(s, from, to)
}
xls_fn_lower <- function(x) tolower(xls_chr(x))
xls_fn_upper <- function(x) toupper(xls_chr(x))
xls_fn_linebreak <- function() "\n"
xls_fn_contains <- function(x, s) grepl(xls_chr(s), xls_chr(x), fixed = TRUE)
xls_fn_starts_with <- function(x, s) startsWith(xls_chr(x), xls_chr(s))
xls_fn_ends_with <- function(x, s) endsWith(xls_chr(x), xls_chr(s))
xls_fn_regex <- function(x, pattern) {
  tryCatch(
    grepl(xls_chr(pattern), xls_chr(x), perl = TRUE),
    error = function(e) grepl(xls_chr(pattern), xls_chr(x))
  )
}

# Numbers
xls_fn_number <- function(x) xls_num(x)
xls_fn_int <- function(x) trunc(xls_num(x))
xls_fn_round <- function(x, digits = 0) round(xls_num(x), xls_num(digits))
xls_fn_abs <- function(x) abs(xls_num(x))
xls_fn_pow <- function(base, exponent) xls_num(base)^xls_num(exponent)
xls_fn_sqrt <- function(x) sqrt(xls_num(x))
xls_fn_exp <- function(x) exp(xls_num(x))
xls_fn_log10 <- function(x) log10(xls_num(x))
xls_fn_pi <- function() pi
xls_fn_sin <- function(x) sin(xls_num(x))
xls_fn_cos <- function(x) cos(xls_num(x))
xls_fn_tan <- function(x) tan(xls_num(x))
xls_fn_asin <- function(x) asin(xls_num(x))
xls_fn_acos <- function(x) acos(xls_num(x))
xls_fn_atan <- function(x) atan(xls_num(x))
xls_fn_atan2 <- function(x, y) atan2(xls_num(x), xls_num(y))
xls_fn_format_number <- function(x) {
  format(xls_num(x), big.mark = ",", scientific = FALSE, trim = TRUE)
}
xls_fn_random <- function() stats::runif(1)

# min() and max() take either several values or one repeated field.
xls_fn_min <- function(...) xls_extreme(c(...), min)
xls_fn_max <- function(...) xls_extreme(c(...), max)
xls_extreme <- function(v, f) {
  n <- xls_num(v)
  if (all(is.na(n))) "" else f(n, na.rm = TRUE)
}

# Dates and times
xls_fn_today <- function() xls_here()$ctx$today
xls_fn_now <- function() xls_here()$state$now
xls_fn_date <- function(x) xls_date(x)
xls_fn_date_time <- function(x) xls_datetime(x)
xls_fn_decimal_date_time <- function(x) as.numeric(xls_datetime(x)) / 86400
xls_fn_decimal_time <- function(x) {
  s <- xls_chr(x)
  t <- regmatches(s, regexpr("[0-9]{1,2}:[0-9]{2}(:[0-9]{2})?", s))
  if (length(t) == 0) {
    return("")
  }
  p <- as.numeric(strsplit(t, ":", fixed = TRUE)[[1]])
  sum(p * c(3600, 60, 1)[seq_along(p)]) / 86400
}
xls_fn_duration <- function() xls_here()$state$duration
xls_fn_format_date_time <- function(x, format) {
  if (xls_is_empty(x)) {
    return("")
  }
  dt <- xls_datetime(x)
  fmt <- xls_chr(format)
  # SurveyCTO tokens that differ from strftime: %n month, %e day and %h hour
  # without padding, %3 milliseconds.
  fmt <- gsub("%n", as.integer(format(dt, "%m")), fmt, fixed = TRUE)
  fmt <- gsub("%e", as.integer(format(dt, "%d")), fmt, fixed = TRUE)
  fmt <- gsub("%h", as.integer(format(dt, "%H")), fmt, fixed = TRUE)
  fmt <- gsub("%3", "000", fmt, fixed = TRUE)
  format(dt, fmt)
}

# Repeats and aggregates
xls_fn_index <- function() {
  idx <- xls_here()$idx
  if (length(idx) == 0) "" else idx[length(idx)]
}
xls_fn_position <- function(...) {
  idx <- xls_here()$idx
  if (length(idx) == 0) "" else idx[length(idx)]
}
xls_fn_count <- function(group) length(group)
xls_fn_sum <- function(field) sum(xls_num(field), na.rm = TRUE)
xls_fn_join <- function(sep, field) {
  v <- xls_chr(field)
  paste(v[nzchar(v)], collapse = xls_chr(sep))
}
xls_fn_count_if <- function(group, expression) {
  x <- xls_here()
  e <- substitute(expression)
  frames <- xls_instance_frames(x, xls_ref_name(substitute(group)))
  sum(vapply(frames, function(f) isTRUE(xls_bool(eval(e, f))[1]), logical(1)))
}
xls_fn_sum_if <- function(field, expression) {
  v <- xls_if_values(xls_here(), substitute(field), substitute(expression))
  sum(xls_num(v), na.rm = TRUE)
}
xls_fn_min_if <- function(field, expression) {
  xls_extreme(xls_if_values(xls_here(), substitute(field), substitute(expression)), min)
}
xls_fn_max_if <- function(field, expression) {
  xls_extreme(xls_if_values(xls_here(), substitute(field), substitute(expression)), max)
}
xls_fn_join_if <- function(sep, field, expression) {
  v <- xls_chr(xls_if_values(xls_here(), substitute(field), substitute(expression)))
  paste(v[nzchar(v)], collapse = xls_chr(sep))
}
xls_if_values <- function(x, field_arg, expr) {
  name <- xls_ref_name(field_arg)
  frames <- xls_instance_frames(x, name)
  keep <- vapply(frames, function(f) isTRUE(xls_bool(eval(expr, f))[1]), logical(1))
  vals <- lapply(frames[keep], function(f) xls_lookup(get(".xls", envir = f), name))
  if (length(vals) == 0) character() else do.call(c, vals)
}
xls_fn_indexed_repeat <- function(field, ...) {
  x <- xls_here()
  name <- xls_ref_name(substitute(field))
  args <- list(...)
  idx <- as.integer(xls_num(unlist(args[seq(2, length(args), by = 2)])))
  v <- xls_answer(x$state, name, idx)
  if (is.na(v)[1]) "" else v
}
xls_fn_rank_index <- function(i, field) {
  v <- xls_num(field)
  i <- as.integer(xls_num(i))
  if (is.na(i) || i < 1 || i > length(v) || is.na(v[i])) {
    return(999)
  }
  rank(-v, ties.method = "first", na.last = "keep")[i]
}

# Lists of items
xls_fn_count_items <- function(sep, x) length(xls_split_list(xls_chr(sep), x))
xls_fn_item_at <- function(sep, x, n) {
  items <- xls_split_list(xls_chr(sep), x)
  i <- as.integer(xls_num(n)) + 1L
  if (is.na(i) || i < 1L || i > length(items)) "" else items[i]
}
xls_fn_item_index <- function(sep, x, value) {
  hit <- match(xls_chr(value), xls_split_list(xls_chr(sep), x))
  if (is.na(hit)) -1 else hit - 1
}
xls_fn_item_present <- function(sep, x, value) {
  xls_chr(value) %in% xls_split_list(xls_chr(sep), x)
}
xls_fn_de_duplicate <- function(sep, x) {
  paste(unique(xls_split_list(xls_chr(sep), x)), collapse = xls_chr(sep))
}
xls_fn_rank_value <- function(value, list) {
  v <- xls_num(strsplit(trimws(xls_chr(list)), "\\s+")[[1]])
  sum(v > xls_num(value), na.rm = TRUE) + 1
}

# Geography
xls_fn_short_geopoint <- function(x) {
  p <- strsplit(trimws(xls_chr(x)), "\\s+")[[1]]
  paste(p[seq_len(min(2, length(p)))], collapse = " ")
}
xls_fn_distance_between <- function(p1, p2) {
  a <- xls_num(strsplit(trimws(xls_chr(p1)), "\\s+")[[1]][1:2]) * pi / 180
  b <- xls_num(strsplit(trimws(xls_chr(p2)), "\\s+")[[1]][1:2]) * pi / 180
  h <- sin((b[1] - a[1]) / 2)^2 + cos(a[1]) * cos(b[1]) * sin((b[2] - a[2]) / 2)^2
  2 * 6371008.8 * asin(sqrt(h))
}

# Identity and session metadata
xls_fn_uuid <- function() xls_uuid()
xls_fn_hash <- function(...) {
  s <- paste(vapply(list(...), function(v) paste(xls_chr(v), collapse = ""), ""), collapse = "|")
  h <- 0
  for (b in utf8ToInt(s)) {
    h <- (h * 31 + b) %% 2147483647
  }
  sprintf("%08x", as.integer(h))
}
xls_fn_username <- function() "enumerator"
xls_fn_enumerator_name <- function() "Enumerator"
xls_fn_enumerator_id <- function() "1"
xls_fn_version <- function() xls_here()$ctx$version
xls_fn_device_info <- function() "fieldwkr dummy_dat"
xls_fn_plug_in_metadata <- function(field) ""
xls_fn_phone_call_log <- function() ""
xls_fn_phone_call_duration <- function() 0
xls_fn_collect_is_phone_app <- function() FALSE

# Functions that need data the form does not contain.
xls_fn_pulldata <- function(...) {
  xls_unresolvable("pulldata() needs the form's attached data")
}
xls_fn_search <- function(...) {
  xls_unresolvable("search() needs the form's attached data")
}
xls_fn_area <- function(...) xls_unresolvable("area() is not simulated")
xls_fn_geo_scatter <- function(...) {
  xls_unresolvable("geo-scatter() is not simulated")
}

#' Labels for one or more space-separated choice values of a select field
#' @keywords internal
#' @noRd
xls_choice_labels <- function(ctx, field, value) {
  list_name <- get0(field, envir = ctx$field_list, inherits = FALSE)
  if (is.null(list_name)) {
    return("")
  }
  rows <- get0(list_name, envir = ctx$lists, inherits = FALSE)
  if (is.null(rows)) {
    return("")
  }
  vals <- xls_selections(value)
  labs <- rows$label[match(vals, rows$value)]
  paste(labs[!is.na(labs)], collapse = " ")
}

#' Random version 4 UUID from R's random number stream (so seeds reproduce it)
#' @keywords internal
#' @noRd
xls_uuid <- function() {
  h <- sample(c(0:9, letters[1:6]), 32, replace = TRUE)
  h[13] <- "4"
  h[17] <- sample(c("8", "9", "a", "b"), 1)
  paste0(
    paste(h[1:8], collapse = ""), "-",
    paste(h[9:12], collapse = ""), "-",
    paste(h[13:16], collapse = ""), "-",
    paste(h[17:20], collapse = ""), "-",
    paste(h[21:32], collapse = "")
  )
}

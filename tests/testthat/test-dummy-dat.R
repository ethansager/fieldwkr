skip_if_not_installed("openxlsx")

today <- as.Date("2024-06-01")

test_that("SurveyCTO expressions translate to R", {
  tr <- fieldwkr:::translate_xlsform_expr
  expect_equal(
    tr("count-selected(${x}) > 1"),
    "xls_fn_count_selected(get_answer(\"x\")) > 1"
  )
  expect_equal(
    tr("${order} mod 2 = 1 or ${a} div 2 >= 3"),
    "get_answer(\"order\") %% 2 == 1 | get_answer(\"a\") / 2 >= 3"
  )
  # operators inside string literals are left alone
  expect_equal(
    tr("concat(${f}, ' and ', ${l})"),
    "xls_fn_concat(get_answer(\"f\"), \" and \", get_answer(\"l\"))"
  )
  # regex backslashes survive
  expect_equal(eval(parse(text = tr("'a\\.b'"))), "a\\.b")
})

test_that("calculations, constraints and choice filters are simulated", {
  survey <- survey_rows(
    type = c(
      "select_one region", "select_one district", "select_multiple crops",
      "calculate", "integer", "integer", "calculate", "integer"
    ),
    name = c("region", "district", "crops", "n_crops", "age", "hhsize", "half", "age_check"),
    constraint = c("", "", "", "", ". >= 18 and . <= 99", ". > 0 and . < 8", "", ". = ${age}"),
    calculation = c("", "", "", "count-selected(${crops})", "", "", "if(${hhsize} > 2, ${hhsize} div 2, 0)", ""),
    choice_filter = c("", "filter = ${region}", "", "", "", "", "", "")
  )
  choices <- data.frame(
    list_name = c("region", "region", "district", "district", "district", "crops", "crops", "crops"),
    value = c("1", "2", "11", "12", "21", "1", "2", "3"),
    label = c("North", "South", "N-a", "N-b", "S-a", "Maize", "Beans", "Rice"),
    filter = c("", "", "1", "1", "2", "", "", "")
  )
  out <- dummy_dat(make_form(survey, choices), n = 30, seed = 1, today = today)

  expect_true(all(substr(out$district, 1, 1) == out$region))
  expect_equal(out$n_crops, lengths(strsplit(out$crops, " ")))
  expect_true(all(out$age >= 18 & out$age <= 99))
  expect_true(all(out$hhsize >= 1 & out$hhsize <= 7))
  expect_equal(out$half, ifelse(out$hhsize > 2, out$hhsize / 2, 0))
  # a re-entry check copies the value rather than drawing until it matches
  expect_equal(out$age_check, out$age)
})

test_that("repeats honour repeat_count and support index() and aggregates", {
  survey <- survey_rows(
    type = c(
      "integer", "begin repeat", "text", "integer", "calculate", "end repeat",
      "calculate", "calculate", "calculate", "calculate"
    ),
    name = c("hhsize", "member", "mname", "mage", "pos", "member", "total_age", "names", "n_adults", "adults"),
    constraint = c(". >= 1 and . <= 5", "", "", ". >= 0 and . <= 90", "", "", "", "", "", ""),
    repeat_count = c("", "${hhsize}", "", "", "", "", "", "", "", ""),
    calculation = c(
      "", "", "", "", "index()", "",
      "sum(${mage})", "join(', ', ${mname})",
      "count-if(${member}, ${mage} >= 18)", "join-if('|', ${mname}, ${mage} >= 18)"
    )
  )
  out <- dummy_dat(make_form(survey), n = 25, seed = 2, today = today, max_repeat = 5)

  ages <- out[instance_cols(out, "mage")]
  names_ <- out[instance_cols(out, "mname")]
  expect_equal(rowSums(!is.na(names_)), out$hhsize, ignore_attr = TRUE)
  for (k in seq_along(instance_cols(out, "pos"))) {
    pos <- out[[instance_cols(out, "pos")[k]]]
    expect_true(all(pos[!is.na(pos)] == k))
  }
  expect_equal(out$total_age, rowSums(ages, na.rm = TRUE), ignore_attr = TRUE)
  expect_equal(out$n_adults, rowSums(ages >= 18, na.rm = TRUE), ignore_attr = TRUE)
  expect_equal(
    out$names,
    apply(names_, 1, function(r) paste(r[!is.na(r)], collapse = ", ")),
    ignore_attr = TRUE
  )
})

test_that("group relevance, notes and unevaluable expressions", {
  survey <- survey_rows(
    type = c("integer", "note", "begin group", "text", "end group", "calculate", "text"),
    name = c("age", "intro", "grp", "inside", "grp", "pd", "asked"),
    relevance = c("", "", "${age} > 200", "", "", "", "pulldata('hh', 'x', 'id', ${age}) = 'y'"),
    constraint = c(". >= 0 and . <= 100", "", "", "", "", "", ""),
    calculation = c("", "", "", "", "", "pulldata('hh', 'x', 'id', ${age})", "")
  )
  expect_warning(
    out <- dummy_dat(make_form(survey), n = 10, seed = 3, today = today),
    "could not be fully simulated"
  )
  expect_false("intro" %in% names(out))
  expect_true(all(is.na(out$inside)))
  expect_true(all(is.na(out$pd)))
  # relevance that cannot be evaluated is treated as relevant
  expect_true(all(!is.na(out$asked)))
  issues <- attr(out, "expression_issues")
  expect_setequal(issues$field, c("pd", "asked"))
  expect_true(all(grepl("pulldata", issues$problem)))
})

test_that("output is reproducible and carries SurveyCTO metadata", {
  survey <- survey_rows(
    type = c("text", "date", "calculate"),
    name = c("t", "d", "stamp"),
    calculation = c("", "", "once(format-date-time(now(), '%Y-%m-%d'))")
  )
  form <- make_form(
    survey,
    settings = data.frame(form_title = "t", form_id = "t", version = "42")
  )
  a <- dummy_dat(form, n = 5, seed = 9, today = today)
  b <- dummy_dat(form, n = 5, seed = 9, today = today)
  expect_identical(a, b)
  expect_true(all(a$stamp == "2024-06-01"))
  expect_true(all(as.Date(a$d) <= today))
  expect_equal(unique(a$formdef_version), "42")
  expect_false(anyDuplicated(a$KEY) > 0)
  expect_true(all(grepl("^uuid:", a$KEY)))
})

test_that("choice values stored as floats (1.0) match expressions using '1'", {
  survey <- survey_rows(
    type = c("select_one yn", "text"),
    name = c("q", "follow"),
    relevance = c("", "selected(${q}, '1')")
  )
  choices <- data.frame(list_name = c("yn", "yn"), value = c("1.0", "0.0"), label = c("Yes", "No"))
  out <- dummy_dat(make_form(survey, choices), n = 40, seed = 4, today = today)
  expect_setequal(unique(out$q), c("1", "0"))
  expect_true(all(is.na(out$follow[out$q == "0"])))
  expect_true(all(!is.na(out$follow[out$q == "1"])))
})

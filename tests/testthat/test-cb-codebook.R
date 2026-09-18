skip_if_not_installed("openxlsx")

test_that("cb_export and cb_apply work", {
  df <- data.frame(a = 1:3, b = c("x", "y", "z"), stringsAsFactors = FALSE)
  attr(df$a, "label") <- "Old label"

  path <- tempfile(fileext = ".xlsx")
  fieldwkr::cb_export(df, path)
  expect_true(file.exists(path))

  survey <- openxlsx::read.xlsx(path, sheet = "survey")
  survey$name[survey$name == "a"] <- "a_new"
  survey$label[survey$name == "a_new"] <- "New label"

  wb <- openxlsx::createWorkbook()
  openxlsx::addWorksheet(wb, "survey")
  openxlsx::writeData(wb, "survey", survey)
  openxlsx::addWorksheet(wb, "choices")
  openxlsx::writeData(wb, "choices", openxlsx::read.xlsx(path, sheet = "choices"))
  openxlsx::saveWorkbook(wb, path, overwrite = TRUE)

  updated <- fieldwkr::cb_apply(df, path)
  expect_true("a_new" %in% names(updated))
  expect_equal(attr(updated$a_new, "label"), "New label")
})

rewrite_codebook <- function(path, survey, choices) {
  wb <- openxlsx::createWorkbook()
  openxlsx::addWorksheet(wb, "survey")
  openxlsx::writeData(wb, "survey", survey)
  openxlsx::addWorksheet(wb, "choices")
  openxlsx::writeData(wb, "choices", choices)
  openxlsx::saveWorkbook(wb, path, overwrite = TRUE)
}

test_that("cb_apply keeps variable labels and Stata formats on labelled columns", {
  skip_if_not_installed("haven")
  df <- data.frame(q = c(1, 2, 1))
  df$q <- haven::labelled(df$q, labels = c(Yes = 1, No = 2), label = "Question")
  attr(df$q, "format.stata") <- "%8.0g"
  path <- tempfile(fileext = ".xlsx")
  fieldwkr::cb_export(df, path)
  out <- fieldwkr::cb_apply(df, path)
  expect_equal(attr(out$q, "label", exact = TRUE), "Question")
  expect_equal(attr(out$q, "format.stata", exact = TRUE), "%8.0g")
  expect_equal(unclass(attr(out$q, "labels")), c(Yes = 1, No = 2))
})

test_that("cb_export handles labelled columns without a variable label", {
  skip_if_not_installed("haven")
  df <- data.frame(q = haven::labelled(c(1, 2), labels = c(Yes = 1, No = 2)))
  path <- tempfile(fileext = ".xlsx")
  expect_silent(fieldwkr::cb_export(df, path))
})

test_that("cb_apply labels character codes from a SurveyCTO export", {
  skip_if_not_installed("haven")
  df <- data.frame(q = c("1", "2", NA), stringsAsFactors = FALSE)
  path <- tempfile(fileext = ".xlsx")
  fieldwkr::cb_export(df, path)
  survey <- openxlsx::read.xlsx(path, "survey")
  survey$choices <- "yn"
  rewrite_codebook(path, survey, data.frame(list_name = "yn", value = c(1, 2), label = c("Yes", "No")))
  out <- fieldwkr::cb_apply(df, path)
  expect_equal(as.numeric(unclass(out$q)), c(1, 2, NA), ignore_attr = TRUE)
})

test_that("cb_apply refuses to drop values missing from a text choice list", {
  df <- data.frame(consent = c(1, 0, 1))
  path <- tempfile(fileext = ".xlsx")
  fieldwkr::cb_export(df, path)
  survey <- openxlsx::read.xlsx(path, "survey")
  survey$choices <- "yn"
  rewrite_codebook(path, survey, data.frame(list_name = "yn", value = c("yes", "no"), label = c("Yes", "No")))
  expect_error(fieldwkr::cb_apply(df, path), "not in choice list 'yn'")
})

test_that("cb_apply stops on a rename collision and cb_append takes one data frame", {
  df <- data.frame(a = 1:2, b = 3:4)
  path <- tempfile(fileext = ".xlsx")
  fieldwkr::cb_export(df, path)
  survey <- openxlsx::read.xlsx(path, "survey")[1, , drop = FALSE]
  survey$name <- "b"
  rewrite_codebook(path, survey, openxlsx::read.xlsx(path, "choices"))
  expect_error(fieldwkr::cb_apply(df, path), "already present")

  expect_silent(fieldwkr::cb_append(list(df), tempfile(fileext = ".xlsx"), "r1"))
})

test_that("cb_from_form builds a codebook that cb_apply can use", {
  skip_if_not_installed("haven")
  form <- make_form(
    data.frame(
      type = c("select_one yn", "select_multiple crops", "note", "geopoint",
               "begin repeat", "integer", "end repeat"),
      name = c("consent", "crops", "hello", "gps", "member", "age", "member"),
      label = c("<b>Consent</b>", "Crops", "Hi", "Location", "", "Age", ""),
      stringsAsFactors = FALSE
    ),
    data.frame(
      list_name = c("yn", "yn", "crops", "crops"),
      value = c("1.0", "0.0", "1", "2"),
      label = c("Yes", "No", "Maize", "Beans")
    )
  )
  path <- tempfile(fileext = ".xlsx")
  fieldwkr::cb_from_form(form, path, survey = "r1")
  survey <- openxlsx::read.xlsx(path, "survey", na.strings = c("", "NA"))
  choices <- openxlsx::read.xlsx(path, "choices")

  expect_false("hello" %in% survey$name)
  expect_true(all(c("gps-Latitude", "gps-Accuracy") %in% survey$name))
  expect_equal(survey$label[survey$name == "consent"], "Consent")
  expect_true(is.na(survey$choices[survey$name == "crops"]))
  expect_equal(sort(choices$value), c("0", "1"))
  # fields inside the repeat are listed but not mapped at the top level
  expect_true(is.na(survey$name_r1[survey$name == "age"]))

  export <- data.frame(consent = c("1", "0"), crops = c("1 2", "2"), stringsAsFactors = FALSE)
  expect_message(
    out <- fieldwkr::cb_apply(export, path, survey = "r1", strict = FALSE),
    "not in the data were skipped"
  )
  expect_equal(attr(out$consent, "label", exact = TRUE), "Consent")
  expect_equal(unclass(attr(out$consent, "labels")), c(Yes = 1, No = 0))
})

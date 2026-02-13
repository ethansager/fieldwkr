skip_if_not_installed("openxlsx")

write_form <- function(path, survey, choices = NULL) {
  wb <- openxlsx::createWorkbook()
  openxlsx::addWorksheet(wb, "survey")
  openxlsx::writeData(wb, "survey", survey)

  if (!is.null(choices)) {
    openxlsx::addWorksheet(wb, "choices")
    openxlsx::writeData(wb, "choices", choices)
  }

  openxlsx::saveWorkbook(wb, path, overwrite = TRUE)
}

test_that("test_form validates a minimal form", {
  path <- tempfile(fileext = ".xlsx")

  survey <- data.frame(
    type = "text",
    name = "respondent",
    label = "Respondent name",
    stringsAsFactors = FALSE
  )

  choices <- data.frame(
    list_name = character(),
    value = character(),
    label = character(),
    stringsAsFactors = FALSE
  )

  write_form(path, survey, choices)

  res <- fieldwkr::test_form(path, verbose = FALSE)
  expect_length(res$errors, 0)
})

test_that("test_form catches missing choices list references", {
  path <- tempfile(fileext = ".xlsx")

  survey <- data.frame(
    type = c("select_one missing_list", "text"),
    name = c("q1", "q2"),
    label = c("Question 1", "Question 2"),
    stringsAsFactors = FALSE
  )

  choices <- data.frame(
    list_name = "other_list",
    name = "a",
    label = "A",
    stringsAsFactors = FALSE
  )

  write_form(path, survey, choices)

  res <- fieldwkr::test_form(path, verbose = FALSE)
  expect_true(any(grepl("Missing choices list\\(s\\) referenced", res$errors)))
})

test_that("test_form catches unknown ${} references", {
  path <- tempfile(fileext = ".xlsx")

  survey <- data.frame(
    type = c("integer", "integer"),
    name = c("age", "income"),
    label = c("Age", "Income"),
    relevance = c("", "${missing_var} > 0"),
    stringsAsFactors = FALSE
  )

  choices <- data.frame(
    list_name = character(),
    name = character(),
    label = character(),
    stringsAsFactors = FALSE
  )

  write_form(path, survey, choices)

  res <- fieldwkr::test_form(path, verbose = FALSE)
  expect_true(any(grepl("Unknown \\$\\{\\} reference", res$errors)))
})

test_that("test_form catches mismatched group/repeat closures", {
  path <- tempfile(fileext = ".xlsx")

  survey <- data.frame(
    type = c("begin group", "text", "end repeat"),
    name = c("grp1", "q1", ""),
    label = c("Group", "Question", ""),
    stringsAsFactors = FALSE
  )

  choices <- data.frame(
    list_name = character(),
    name = character(),
    label = character(),
    stringsAsFactors = FALSE
  )

  write_form(path, survey, choices)

  res <- fieldwkr::test_form(path, verbose = FALSE)
  expect_true(any(grepl("Mismatched 'end repeat'", res$errors)))
})

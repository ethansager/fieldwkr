skip_if_not_installed("openxlsx")

test_that("ietestform validates a minimal form", {
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

  write_xlsx_sheets(path, list(survey = survey, choices = choices))

  res <- ietestform(path, verbose = FALSE)
  expect_length(res$errors, 0)
})

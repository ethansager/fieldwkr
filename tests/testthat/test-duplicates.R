skip_if_not_installed("openxlsx")

test_that("duplicates creates a report", {
  df <- data.frame(
    id = c(1, 1, 2),
    uid = c("a", "b", "c"),
    x = c(10, 11, 20),
    stringsAsFactors = FALSE
  )

  report_path <- tempfile(fileext = ".xlsx")
  res <- fieldwkr::duplicates(df, idvar = "id", uniquevars = "uid", report_path = report_path, apply = FALSE)

  expect_true(file.exists(report_path))
  expect_true(is.data.frame(res$report))
  expect_gt(nrow(res$report), 0)
})

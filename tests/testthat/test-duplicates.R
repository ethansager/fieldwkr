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

adjudicate <- function(path, edit) {
  report <- openxlsx::read.xlsx(path, "duplicates")
  report <- edit(report)
  wb <- openxlsx::createWorkbook()
  openxlsx::addWorksheet(wb, "duplicates")
  openxlsx::writeData(wb, "duplicates", report)
  openxlsx::saveWorkbook(wb, path, overwrite = TRUE)
}

dup_df <- data.frame(
  id = c(1, 1, 2, 3),
  KEY = c("uuid:a", "uuid:b", "uuid:c", "uuid:d"),
  x = c(10, 11, 20, 30),
  stringsAsFactors = FALSE
)

test_that("marking one record drops only that record, and reruns keep the report", {
  path <- tempfile(fileext = ".xlsx")
  suppressMessages(fieldwkr::duplicates(dup_df, "id", "KEY", path, apply = FALSE))
  adjudicate(path, function(r) { r$drop[r$KEY == "uuid:b"] <- "drop"; r })

  res <- suppressMessages(fieldwkr::duplicates(dup_df, "id", "KEY", path))
  expect_equal(res$data$KEY, c("uuid:a", "uuid:c", "uuid:d"))
  expect_equal(openxlsx::read.xlsx(path, "duplicates")$drop[2], "drop")
})

test_that("newid changes only the marked record and keeps a numeric id numeric", {
  path <- tempfile(fileext = ".xlsx")
  suppressMessages(fieldwkr::duplicates(dup_df, "id", "KEY", path, apply = FALSE))
  adjudicate(path, function(r) { r$newid[r$KEY == "uuid:b"] <- 99; r })

  res <- suppressMessages(fieldwkr::duplicates(dup_df, "id", "KEY", path))
  expect_equal(res$data$id, c(1, 99, 2, 3))
})

test_that("blank submission keys are reported as such", {
  df <- dup_df
  df$KEY[4] <- ""
  expect_error(
    fieldwkr::duplicates(df, "id", "KEY", tempfile(fileext = ".xlsx"), apply = FALSE),
    "1 row\\(s\\) have a missing or blank value in uniquevars"
  )
})

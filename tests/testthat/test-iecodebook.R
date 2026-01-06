skip_if_not_installed("openxlsx")

test_that("iecodebook export and apply work", {
  df <- data.frame(a = 1:3, b = c("x", "y", "z"), stringsAsFactors = FALSE)
  attr(df$a, "label") <- "Old label"

  path <- tempfile(fileext = ".xlsx")
  iecodebook_export(df, path)
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

  updated <- iecodebook_apply(df, path)
  expect_true("a_new" %in% names(updated))
  expect_equal(attr(updated$a_new, "label"), "New label")
})

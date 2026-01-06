skip_if_not_installed("openxlsx")

test_that("iecorrect applies numeric corrections", {
  df <- data.frame(id = c(1, 2), x = c(1, 2), stringsAsFactors = FALSE)

  path <- tempfile(fileext = ".xlsx")
  iecorrect_template(path, idvars = "id")

  numeric_sheet <- data.frame(
    id = 1,
    varname = "x",
    value = 99,
    valuecurrent = 1,
    initials = "",
    notes = "",
    stringsAsFactors = FALSE
  )

  wb <- openxlsx::loadWorkbook(path)
  openxlsx::writeData(wb, "numeric", numeric_sheet)
  openxlsx::saveWorkbook(wb, path, overwrite = TRUE)

  updated <- iecorrect_apply(df, path, idvars = "id", sheets = "numeric")
  expect_equal(updated$x[updated$id == 1], 99)
})

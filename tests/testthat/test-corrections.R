skip_if_not_installed("openxlsx")

test_that("correct_apply applies numeric corrections", {
  df <- data.frame(id = c(1, 2), x = c(1, 2), stringsAsFactors = FALSE)

  path <- tempfile(fileext = ".xlsx")
  fieldwkr::correct_temp(path, idvars = "id")

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

  updated <- fieldwkr::correct_apply(df, path, idvars = "id", sheets = "numeric")
  expect_equal(updated$x[updated$id == 1], 99)
})

write_sheet <- function(path, sheet, rows) {
  wb <- openxlsx::loadWorkbook(path)
  openxlsx::writeData(wb, sheet, rows)
  openxlsx::saveWorkbook(wb, path, overwrite = TRUE)
}

test_that("drops leave rows with a missing ID untouched", {
  df <- data.frame(hh = c("1", "1", "2"), mem = c("1", NA, "1"), x = 1:3, stringsAsFactors = FALSE)
  path <- tempfile(fileext = ".xlsx")
  fieldwkr::correct_temp(path, idvars = c("hh", "mem"))
  write_sheet(path, "drop", data.frame(hh = "1", mem = "1", n_obs = 1, initials = "", notes = ""))
  out <- fieldwkr::correct_apply(df, path, idvars = c("hh", "mem"), sheets = "drop")
  expect_equal(out$x, c(2L, 3L))
})

test_that("blank IDs, unparseable numbers and unknown variables stop the run", {
  df <- data.frame(id = c(1, 2), inc = c(100, 200))
  corrections <- function(rows) {
    path <- tempfile(fileext = ".xlsx")
    fieldwkr::correct_temp(path, idvars = "id")
    write_sheet(path, "numeric", rows)
    path
  }
  row <- function(id, varname, value) {
    data.frame(id = id, varname = varname, value = value, valuecurrent = NA,
               initials = "", notes = "", stringsAsFactors = FALSE)
  }
  expect_error(
    fieldwkr::correct_apply(df, corrections(row(NA, "inc", 5)), "id", "numeric"),
    "Blank 'id'"
  )
  expect_error(
    fieldwkr::correct_apply(df, corrections(row(1, "inc", "1,500")), "id", "numeric"),
    "is not numeric"
  )
  expect_error(
    fieldwkr::correct_apply(df, corrections(row(1, "incom", 5)), "id", "numeric"),
    "is not in the data"
  )
})

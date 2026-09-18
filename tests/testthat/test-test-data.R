test_that("test_data validates a clean dataset", {
  df <- data.frame(
    id = 1:3,
    age = c(21, 34, 55),
    gender = c("female", "male", "female"),
    stringsAsFactors = FALSE
  )

  res <- fieldwkr::test_data(
    df,
    required_cols = c("id", "age", "gender"),
    unique_keys = "id",
    non_missing_cols = c("id", "age", "gender"),
    value_ranges = list(age = c(0, 120)),
    allowed_values = list(gender = c("female", "male")),
    min_rows = 1,
    verbose = FALSE
  )

  expect_length(res$errors, 0)
  expect_length(res$warnings, 0)
})

test_that("test_data catches required-column and key issues", {
  df <- data.frame(
    id = c(1, 1, NA),
    name = c("Alice", "", "Cara"),
    stringsAsFactors = FALSE
  )

  res <- fieldwkr::test_data(
    df,
    required_cols = c("id", "name", "age"),
    unique_keys = "id",
    non_missing_cols = "name",
    verbose = FALSE
  )

  expect_true(any(grepl("Missing required column\\(s\\): age", res$errors)))
  expect_true(any(grepl("Missing values in unique key column\\(s\\)", res$errors)))
  expect_true(any(grepl("Duplicate key combinations found", res$errors)))
  expect_true(any(grepl("Missing values in 'name'", res$errors)))
})

test_that("test_data catches range and allowed-value failures", {
  df <- data.frame(
    age = c(25, 130),
    gender = c("female", "other"),
    stringsAsFactors = FALSE
  )

  res <- fieldwkr::test_data(
    df,
    value_ranges = list(age = c(0, 120)),
    allowed_values = list(gender = c("female", "male")),
    verbose = FALSE
  )

  expect_true(any(grepl("Values outside \\[0, 120\\] found in 'age'", res$errors)))
  expect_true(any(grepl("Unexpected value\\(s\\) in 'gender'", res$errors)))
})

test_that("test_data reports empty data and row limits", {
  df <- data.frame(id = integer(), stringsAsFactors = FALSE)

  res <- fieldwkr::test_data(
    df,
    min_rows = 1,
    verbose = FALSE
  )

  expect_true(any(grepl("Data has 0 rows\\.", res$warnings)))
  expect_true(any(grepl("minimum required is 1\\.", res$errors)))
})

test_that("happy path: two CSVs produce correct wide output", {
  tmp <- tempdir()
  uuid1 <- "aaaaaaaa-0000-0000-0000-000000000001"
  uuid2 <- "aaaaaaaa-0000-0000-0000-000000000002"

  write.csv(
    data.frame(
      `Field name` = c("Consented[1]/I[1]/customers[1]/name",
                       "Consented[1]/I[1]/customers[1]/age"),
      Comment      = c("wrong name", "too old"),
      check.names  = FALSE
    ),
    file.path(tmp, paste0("Comments-", uuid1, ".csv")),
    row.names = FALSE
  )

  write.csv(
    data.frame(
      `Field name` = c("Consented[1]/I[1]/customers[1]/name",
                       "Consented[1]/I[1]/customers[1]/city"),
      Comment      = c("misspelled", "unknown"),
      check.names  = FALSE
    ),
    file.path(tmp, paste0("Comments-", uuid2, ".csv")),
    row.names = FALSE
  )

  on.exit({
    file.remove(file.path(tmp, paste0("Comments-", uuid1, ".csv")))
    file.remove(file.path(tmp, paste0("Comments-", uuid2, ".csv")))
  })

  result <- fieldwkr::read_comments(tmp)

  expect_equal(nrow(result), 2)
  expect_true("uuid" %in% names(result))
  expect_true("name" %in% names(result))
  expect_true("age" %in% names(result))
  expect_true("city" %in% names(result))
  expect_equal(names(result)[1], "uuid")

  row1 <- result[result$uuid == uuid1, ]
  expect_equal(row1$name, "wrong name")
  expect_equal(row1$age, "too old")
  expect_true(is.na(row1$city))
})

test_that("duplicate field names are numbered", {
  tmp <- tempdir()
  uuid <- "bbbbbbbb-0000-0000-0000-000000000001"

  write.csv(
    data.frame(
      `Field name` = c("repeat[1]/group[1]/field",
                       "repeat[2]/group[1]/field"),
      Comment      = c("first", "second"),
      check.names  = FALSE
    ),
    file.path(tmp, paste0("Comments-", uuid, ".csv")),
    row.names = FALSE
  )

  on.exit(file.remove(file.path(tmp, paste0("Comments-", uuid, ".csv"))))

  result <- fieldwkr::read_comments(tmp)

  expect_equal(nrow(result), 1)
  expect_true("field_1" %in% names(result))
  expect_true("field_2" %in% names(result))
  expect_equal(result$field_1, "first")
  expect_equal(result$field_2, "second")
})

test_that("empty CSV (header only) produces a uuid-only row", {
  tmp <- tempdir()
  uuid <- "cccccccc-0000-0000-0000-000000000001"

  write.csv(
    data.frame(`Field name` = character(0), Comment = character(0),
               check.names = FALSE),
    file.path(tmp, paste0("Comments-", uuid, ".csv")),
    row.names = FALSE
  )

  on.exit(file.remove(file.path(tmp, paste0("Comments-", uuid, ".csv"))))

  result <- fieldwkr::read_comments(tmp)

  expect_equal(nrow(result), 1)
  expect_equal(result$uuid, uuid)
  expect_equal(ncol(result), 1L)
})

test_that("missing fields across files become NA", {
  tmp <- tempdir()
  uuid1 <- "dddddddd-0000-0000-0000-000000000001"
  uuid2 <- "dddddddd-0000-0000-0000-000000000002"

  write.csv(
    data.frame(`Field name` = "grp[1]/foo", Comment = "val1",
               check.names = FALSE),
    file.path(tmp, paste0("Comments-", uuid1, ".csv")),
    row.names = FALSE
  )
  write.csv(
    data.frame(`Field name` = "grp[1]/bar", Comment = "val2",
               check.names = FALSE),
    file.path(tmp, paste0("Comments-", uuid2, ".csv")),
    row.names = FALSE
  )

  on.exit({
    file.remove(file.path(tmp, paste0("Comments-", uuid1, ".csv")))
    file.remove(file.path(tmp, paste0("Comments-", uuid2, ".csv")))
  })

  result <- fieldwkr::read_comments(tmp)

  expect_equal(nrow(result), 2)
  row1 <- result[result$uuid == uuid1, ]
  row2 <- result[result$uuid == uuid2, ]
  expect_equal(row1$foo, "val1")
  expect_true(is.na(row1$bar))
  expect_equal(row2$bar, "val2")
  expect_true(is.na(row2$foo))
})

test_that("no matching files causes an error", {
  tmp <- tempdir()
  expect_error(fieldwkr::read_comments(tmp), "No Comments-.*\\.csv")
})

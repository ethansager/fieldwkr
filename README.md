# fieldwkr

R port of selected field survey workflows for primary data collection.

## Installation

```r
# devtools::install_github("ethansager/fieldwkr")
```

## Usage

```r
library(fieldwkr)

# Example: export and apply a codebook
cb_path <- tempfile(fileext = ".xlsx")
cb_export(mtcars, cb_path)
updated <- cb_apply(mtcars, cb_path)

# Validate an XLSForm before deployment
form_check <- test_form("my_form.xlsx", verbose = FALSE)
form_check$errors
form_check$warnings

# Build a codebook from the SurveyCTO form, then use it to label an export.
# strict = FALSE skips form fields that this export does not contain.
cb_from_form("my_form.xlsx", "codebook.xlsx", survey = "baseline")
labelled <- cb_apply(export, "codebook.xlsx", survey = "baseline", strict = FALSE)

# Generate mock submissions from the form to test a cleaning pipeline.
# Fixing seed and today makes the output reproducible.
mock <- dummy_dat("my_form.xlsx", n = 100, seed = 1, today = as.Date("2024-06-01"))
attr(mock, "expression_issues") # expressions that could not be simulated, e.g. pulldata()

# Unit-test a dataset against field rules
data_check <- test_data(
  mtcars,
  required_cols = c("mpg", "cyl"),
  non_missing_cols = c("mpg", "cyl"),
  value_ranges = list(mpg = c(0, 80)),
  min_rows = 1,
  verbose = FALSE
)
data_check$errors
```

## Notes

- Requires the `openxlsx` package; `haven` is optional for labelled data.
- This package mirrors the Stata workflows at a practical level but does not
  implement every validation rule in the original commands.
- Public API uses canonical `fieldwkr` function names (`cb_*`, `correct_*`,
  `comp_dup`, `duplicates`, `read_comments`, `test_form`, `test_data`,
  `dummy_dat`).
- In `duplicates()`, `idvar` is the project identifier being adjudicated (for
  example a household ID from the sample frame) and `uniquevars` identifies a
  single submission, which in SurveyCTO data is normally `KEY`.

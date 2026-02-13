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
```

## Notes

- Requires the `openxlsx` package; `haven` is optional for labelled data.
- This package mirrors the Stata workflows at a practical level but does not
  implement every validation rule in the original commands.
- Public API uses canonical `fieldwkr` function names (`cb_*`, `correct_*`,
  `comp_dup`, `duplicates`, `test_form`, `dummy_dat`).

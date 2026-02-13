#' Print package version information
#'
#' @details
#' Utility for quickly checking the installed package build in field workflow
#' environments where multiple versions may circulate.
#' @return A list with version and version_date.
#' @export
fieldwkr <- function() {
  version <- "0.1"
  version_date <- "06JAN2026"

  message("")
  message(sprintf(
    "This version of fieldwkr installed is version %s",
    version
  ))
  message(sprintf(
    "This version of fieldwkr was released on %s",
    version_date
  ))

  invisible(list(version = version, version_date = version_date))
}

#' Print package version information
#'
#' @return A list with version and version_date.
#' @export
iefieldkit <- function() {
  version <- "3.2"
  version_date <- "31JUL2023"

  message("")
  message(sprintf("This version of iefieldkit installed is version %s", version))
  message(sprintf("This version of iefieldkit was released on %s", version_date))
  message("")
  message("This package includes the following commands:")
  message("- iecodebook")
  message("- ieduplicates / iecompdup")
  message("- ietestform")

  invisible(list(version = version, version_date = version_date))
}

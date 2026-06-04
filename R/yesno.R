#' Scan for Yes/No Variables
#'
#' @description
#' `scan_yesno()` identifies all factor variables in a dataset whose first two
#' levels are "yes" and "no" (case-insensitive). Returns a tibble of variable
#' names and their labels, useful for identifying candidates for `relabel_yesno()`.
#'
#' @param data A data frame.
#'
#' @return A tibble with columns `variable` and `label`. If no yes/no variables
#'   are found, returns `NULL` invisibly with a message.
#'
#' @examples
#' \dontrun{
#' scan_yesno(data)
#' }
#' @importFrom tibble tibble
#' @importFrom labelled var_label
scan_yesno <- function(data) {
  hits <- names(data)[sapply(data, function(x) {
    is.factor(x) && all(tolower(levels(x)[1:2]) %in% c("yes", "no"))
  })]

  if (length(hits) == 0) {
    message("No yes/no variables found in data.")
    return(invisible(NULL))
  }

  tibble(
    variable = hits,
    label    = sapply(hits, function(v) {
      lbl <- var_label(data[[v]])
      if (is.null(lbl) || lbl == "") NA_character_ else as.character(lbl)
    })
  )
}


#' Relabel Yes/No Factor Variables
#'
#' @description
#' `relabel_yesno()` replaces the first two levels ("Yes" and "No") of one or
#' more factor variables with meaningful labels, leaving DK/Refused/Web Blank
#' levels untouched. Use `scan_yesno()` first to identify candidates.
#' @export
#' @param data A data frame.
#' @param ... Named arguments where each name is a variable in `data` and each
#'   value is a character vector of length 2 giving the new labels for the yes
#'   and no levels respectively.
#'
#' @return The original data frame with relabelled factor levels.
#'
#' @examples
#' \dontrun{
#' data <- relabel_yesno(data,
#'   MAHA1 = c("MAHA supporter", "Not MAHA supporter"),
#'   MAGA1 = c("MAGA Rep/leaner", "Non-MAGA Rep/leaner"),
#'   RVOTE = c("Reg. voter", "Not reg. voter"),
#'   CHILD = c("Parent", "Not parent")
#' )
#'}
#'
#' @importFrom labelled var_label
relabel_yesno <- function(data, ...) {
  relabels <- list(...)

  # Check all provided variables exist in data
  missing_vars <- names(relabels)[!names(relabels) %in% names(data)]
  if (length(missing_vars) > 0) {
    stop(paste0(
      "The following variables are not in the dataset: ",
      paste(missing_vars, collapse = ", ")
    ))
  }

  # Check all provided variables are actually yes/no
  not_yesno <- names(relabels)[!names(relabels) %in% names(data)[sapply(data, function(x) {
    is.factor(x) && all(tolower(levels(x)[1:2]) %in% c("yes", "no"))
  })]]
  if (length(not_yesno) > 0) {
    stop(paste0(
      "The following variables are not yes/no factors: ",
      paste(not_yesno, collapse = ", ")
    ))
  }

  # Check all provided replacements are length 2
  wrong_length <- names(relabels)[sapply(relabels, length) != 2]
  if (length(wrong_length) > 0) {
    stop(paste0(
      "Replacements must have exactly 2 labels (yes replacement, no replacement). Check: ",
      paste(wrong_length, collapse = ", ")
    ))
  }

  for (v in names(relabels)) {
    new_labels   <- relabels[[v]]
    old_levels   <- levels(data[[v]])
    old_levels[1:2] <- new_labels
    levels(data[[v]]) <- old_levels
  }

  data
}

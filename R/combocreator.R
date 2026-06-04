#' Combine a Parent and Child Factor Variable
#'
#' @description
#' `combocreator()` combines a parent factor variable and a child factor variable
#' that was only asked of a subset of respondents into a single combo variable.
#' For respondents who were asked the child question, the combo variable takes the
#' form `"[parent label] - [child label]"`. For respondents who were not asked the
#' child question, the combo variable retains the parent label. The new variable is
#' placed immediately after the child variable in the dataset.
#'
#' The function auto-detects which parent groups were asked the child question by
#' finding parent levels with any non-NA child responses. This can be overridden
#' by supplying `child_asked_values` manually.
#'
#' Level ordering follows the original factor ordering of `var1` and `var2` rather
#' than alphabetical sorting.
#' @export
#' @param data A data frame.
#' @param var1 The parent factor variable.
#' @param var2 The child factor variable (asked of a subset of `var1` levels).
#' @param child_asked_values Optional character vector of `var1` levels for which
#'   `var2` was asked. If `NULL` (default), these are auto-detected from the data.
#'
#' @return The original data frame with a new factor variable named
#'   `{var1}{var2}_combo` inserted immediately after `var2`.
#'
#' @examples
#' \dontrun{
#' # Auto-detect which groups were asked the child question
#' data <- combocreator(data, party3, MAGA1)
#'
#' # Manually specify which groups were asked
#' data <- combocreator(data, party3, maga1,
#'   child_asked_values = c("Republican", "Republican-leaning Independent"))
#' }
#' @importFrom rlang ensym as_string sym
#' @importFrom dplyr filter pull relocate case_when
#' @importFrom labelled var_label var_label<-

combocreator <- function(
    data,
    var1,
    var2,
    child_asked_values = NULL
) {
  var1_sym <- ensym(var1)
  var2_sym <- ensym(var2)
  var1_str <- as_string(var1_sym)
  var2_str <- as_string(var2_sym)
  combo_name <- paste0(var1_str, var2_str, "_combo")

  # Check var1 and var2 exist in data
  if (!var1_str %in% names(data)) stop(paste0("'", var1_str, "' is not a variable in the dataset."))
  if (!var2_str %in% names(data)) stop(paste0("'", var2_str, "' is not a variable in the dataset."))

  # Check both are factors
  if (!is.factor(data[[var1_str]])) stop(paste0("'", var1_str, "' is not a factor."))
  if (!is.factor(data[[var2_str]])) stop(paste0("'", var2_str, "' is not a factor."))

  # Check var1 and var2 aren't the same
  if (var1_str == var2_str) stop("'var1' and 'var2' cannot be the same variable.")

  # Check child has any non-NA values at all
  if (all(is.na(data[[var2_str]]))) stop(paste0("'", var2_str, "' is entirely NA."))

  parent_labels_chr <- as.character(data[[var1_str]])
  child_labels_chr  <- as.character(data[[var2_str]])
  valid_parent_levels <- as.character(levels(data[[var1_str]]))
  valid_child_levels  <- as.character(levels(data[[var2_str]]))

  # Validate and check manually provided child_asked_values all at once
  if (!is.null(child_asked_values)) {
    bad_values <- child_asked_values[!child_asked_values %in% valid_parent_levels]
    if (length(bad_values) > 0) {
      stop(paste0(
        paste0("'", bad_values, "'", collapse = ", "),
        " is not an answer option in ", var1_str, "."
      ))
    }
    empty_asked <- child_asked_values[sapply(child_asked_values, function(v) {
      all(is.na(data[[var2_str]][parent_labels_chr == v]))
    })]
    if (length(empty_asked) > 0) {
      stop(paste0(
        paste0(var2_str, " was not asked of ", var1_str, " = '", empty_asked, "'", collapse = "\n")
      ))
    }
  }

  # Auto-detect child_asked_values if not provided
  if (is.null(child_asked_values)) {
    child_asked_values <- data %>%
      filter(!is.na(.data[[var2_str]])) %>%
      pull(.data[[var1_str]]) %>%
      as.character() %>%
      unique()
  }

  # Check that var1 and var2 are actually filtered on one another
  not_asked <- as.character(valid_parent_levels[!valid_parent_levels %in% child_asked_values])
  if (length(not_asked) == 0) stop("These two questions are not filtered on one another.")

  # Warn if groups not in child_asked_values have non-NA child responses
  accidental_data <- not_asked[sapply(not_asked, function(v) {
    any(!is.na(data[[var2_str]][parent_labels_chr == v]))
  })]
  if (length(accidental_data) > 0) {
    warning(paste0(
      var2_str, " has non-NA responses among groups not identified as asked:\n",
      paste0("  ", var1_str, " = '", accidental_data, "'", collapse = "\n"),
      "\nConsider reviewing your child_asked_values."
    ))
  }

  # Warn if combo variable already exists
  if (combo_name %in% names(data)) {
    warning(paste0("'", combo_name, "' already exists in the dataset and will be overwritten."))
  }

  combo_var <- case_when(
    data[[var1_str]] %in% child_asked_values & !is.na(data[[var2_str]]) ~
      paste(parent_labels_chr, "-", child_labels_chr),
    !data[[var1_str]] %in% child_asked_values & !is.na(data[[var1_str]]) ~
      parent_labels_chr,
    TRUE ~ NA_character_
  )

  # Preserve original factor ordering from var1 and var2
  parent_combos <- valid_parent_levels[valid_parent_levels %in% child_asked_values]
  ordered_levels <- unlist(lapply(parent_combos, function(p) {
    paste(p, "-", valid_child_levels)
  }))
  not_asked_levels <- valid_parent_levels[!valid_parent_levels %in% child_asked_values]
  combo_levels <- c(
    ordered_levels[ordered_levels %in% combo_var],
    not_asked_levels[not_asked_levels %in% combo_var]
  )
  combo_var <- factor(combo_var, levels = combo_levels)

  data[[combo_name]] <- combo_var

  # Use var2 label if available, otherwise fall back to var name
  child_label <- var_label(data[[var2_str]])
  if (is.null(child_label) || child_label == "") {
    child_label <- var2_str
  }
  var_label(data[[combo_name]]) <- paste0(child_label, " (combined with ", var1_str, ")")

  data %>%
    relocate(!!sym(combo_name), .after = !!var2_sym)
}

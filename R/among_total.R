#' Create a Total-Sample Version of a Filtered Variable
#'
#' @description
#' `among_total()` takes a variable that was only asked of a subset of respondents
#' (and is therefore NA for others) and creates a new version of it that represents
#' the total sample. The NA values are given an explicit label (e.g., "(Costs have
#' not increased)") that describes why those respondents were not asked the question.
#' The new variable is placed immediately after the original in the dataset.
#'
#' Substantive levels appear first, followed by the explicit NA label, followed by
#' any DK/Refused/Web Blank levels. DK detection is case-insensitive.
#'
#' @param data A data frame.
#' @param var The variable to total. Must be a factor.
#' @param na_level A character string giving the label to apply to NA values
#'   (e.g., `"(Costs have not increased)"`). By convention, wrap in parentheses
#'   to indicate these respondents were not asked the question.
#'
#' @return The original data frame with a new factor variable named `{var}_Total`
#'   inserted immediately after `var`.
#'
#' @examples
#' data <- among_total(data, Q4_A, "(Costs have not increased)")
#' # Creates Q4_A_Total with NAs labeled and placed after substantive responses
#'
#' @importFrom forcats fct_explicit_na
#' @importFrom rlang ensym as_string sym
#' @importFrom dplyr relocate
among_total <- function(data, var, na_level = "(Missing)") {
  var_sym <- ensym(var)
  var_str <- as_string(var_sym)
  new_name <- paste0(var_str, "_tot")

  # Check var exists in data
  if (!var_str %in% names(data)) {
    stop(paste0("'", var_str, "' is not a variable in the dataset."))
  }

  # Check var is a factor
  if (!is.factor(data[[var_str]])) {
    stop(paste0("'", var_str, "' is not a factor."))
  }

  # Check na_level is a single character string
  if (!is.character(na_level) || length(na_level) != 1) {
    stop("'na_level' must be a single character string.")
  }

  # Check there are actually NAs to fill
  if (!anyNA(data[[var_str]])) {
    warning(paste0("'", var_str, "' has no missing values. The _tot variable will be identical to the original."))
  }

  # Warn if new variable already exists
  if (new_name %in% names(data)) {
    warning(paste0("'", new_name, "' already exists in the dataset and will be overwritten."))
  }

  # Pull label before any transformation strips it
  orig_label <- attr(data[[var_str]], "label")

  dk_levels <- c("not sure", "refused", "web blank", "don't know")
  x <- fct_explicit_na(data[[var_str]], na_level = na_level)
  substantive <- levels(x)[!tolower(levels(x)) %in% c(dk_levels, tolower(na_level))]
  dk_present  <- levels(x)[tolower(levels(x)) %in% dk_levels]
  data[[new_name]] <- factor(x, levels = c(substantive, na_level, dk_present))

  if (!is.null(orig_label)) {
    attr(data[[new_name]], "label") <- orig_label
  }

  data %>%
    relocate(!!sym(new_name), .after = !!var_sym)

  }

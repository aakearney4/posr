#' Collapse and Recode Factor Levels
#'
#' @description
#' `netter()` collapses one or more groups of factor levels into single levels,
#' optionally renaming them. Levels not mentioned in `groups` are left untouched
#' and appear after the collapsed levels in their original order. Useful for
#' creating nets (e.g., top-2-box) or collapsing response categories.
#' @export
#' @param data A data frame.
#' @param var The factor variable to recode.
#' @param groups A named or unnamed list of integer vectors specifying which
#'   levels (by position) to collapse. If named, the name is used as the new
#'   label. If unnamed, the original labels are collapsed and joined with "/".
#' @param label Optional character string for the variable label of the new
#'   variable. If `NULL`, inherits the label from `var`.
#' @param new_name Optional character string for the name of the new variable.
#'   Defaults to `{var}_rec`. If `{var}_rec` already exists, defaults to
#'   `{var}_rec2` with a warning.
#'
#' @return The original data frame with the new recoded factor variable inserted
#'   immediately after `var`.
#'
#' @examples
#' \dontrun{
#' # Collapse levels 1:3, keep original labels joined with "/"
#' data <- netter(data, Q10_A, groups = list(1:3))
#'
#' # Collapse and rename one group, leave rest as-is
#' data <- netter(data, Q11_A,
#'   groups = list("Uses AI" = 1:4),
#'   label  = "AI USE"
#' )
#'
#' # Collapse and rename multiple groups, DK/Refused left untouched
#' data <- netter(data, Q11_A,
#'   groups   = list("Uses AI" = 1:4, "Does not use AI" = 5),
#'   label    = "AI USE",
#'   new_name = "Q11A_REC"
#' )
#' }
#' @importFrom rlang ensym as_string sym
#' @importFrom dplyr relocate
#' @importFrom labelled var_label var_label<-

netter <- function(data, var, groups, label = NULL, new_name = NULL) {
  var_sym <- ensym(var)
  var_str <- as_string(var_sym)

  # Check var exists and is a factor
  if (!var_str %in% names(data)) stop(paste0("'", var_str, "' is not a variable in the dataset."))
  if (!is.factor(data[[var_str]])) stop(paste0("'", var_str, "' is not a factor."))

  # Handle new_name defaulting and conflicts
  if (is.null(new_name)) {
    new_name <- paste0(var_str, "_rec")
    if (new_name %in% names(data)) {
      warning(paste0("'", new_name, "' already exists, creating '", var_str, "_rec2' instead."))
      new_name <- paste0(var_str, "_rec2")
    }
  } else {
    if (new_name %in% names(data)) {
      warning(paste0("'", new_name, "' already exists and will be overwritten."))
    }
  }

  existing_levels <- levels(data[[var_str]])
  new_factor <- as.character(data[[var_str]])
  grouped_positions <- unlist(groups)

  # Process each group
  for (i in seq_along(groups)) {
    grp <- groups[[i]]
    nm  <- if (!is.null(names(groups))) names(groups)[i] else NULL
    grp_name <- if (is.null(nm) || nm == "") paste(existing_levels[grp], collapse = "/") else nm
    new_factor[new_factor %in% existing_levels[grp]] <- grp_name
  }

  # Build final level order preserving original order
  grouped_new_names <- sapply(seq_along(groups), function(i) {
    grp <- groups[[i]]
    nm  <- if (!is.null(names(groups))) names(groups)[i] else NULL
    if (is.null(nm) || nm == "") paste(existing_levels[grp], collapse = "/") else nm
  })

  ungrouped_levels <- existing_levels[!seq_along(existing_levels) %in% grouped_positions]
  final_levels <- unique(c(grouped_new_names, ungrouped_levels))
  final_levels <- final_levels[final_levels %in% new_factor]

  data[[new_name]] <- factor(new_factor, levels = final_levels)

  # Set variable label
  if (!is.null(label)) {
    var_label(data[[new_name]]) <- label
  } else if (!is.null(var_label(data[[var_str]]))) {
    var_label(data[[new_name]]) <- var_label(data[[var_str]])
  }

  data %>%
    relocate(!!sym(new_name), .after = !!var_sym)
}

#' Frequency Table with Optional Weighting
#'
#' @description
#' `freq()` produces a frequency table for a single variable, with optional
#' survey weighting. Returns a data frame with percentages and unweighted counts,
#' with a Total row at the bottom. Factor level ordering is preserved. NA values
#' are either excluded (default) or shown as "(Missing)".
#'
#' The percentage column is labeled `weighted_pct` when a weight is supplied and
#' `unweighted_pct` when not. Counts are always unweighted.
#'
#' @param data A data frame.
#' @param var A character string giving the name of the variable to tabulate.
#' @param weight Optional character string giving the name of the weight variable.
#'   If `NULL` (default), unweighted counts are used.
#' @param include_na Logical. If `TRUE`, NA values are included as a `"(Missing)"`
#'   row. Default is `FALSE`.
#' @param digits Integer. Number of decimal places for percentages. Default is `0`.
#'
#' @return A data frame with columns for the variable values, percentage
#'   (`weighted_pct` or `unweighted_pct`), and `unweighted_n`. The final row
#'   shows the total percentage and total unweighted N.
#'
#' @examples
#' \dontrun{
#' # Basic unweighted frequency
#' freq(data, "party3")
#'
#' # Weighted frequency
#' freq(data, "party3", weight = "WEIGHT_PID_ADJ")
#'
#' # Weighted frequency including NAs
#' freq(data, "MAGA1", weight = "WEIGHT_PID_ADJ", include_na = TRUE)
#'
#' # Weighted frequency with decimal places
#' freq(data, "party3", weight = "WEIGHT_PID_ADJ", digits = 1)
#'
#' # On a combo variable
#' freq(data, "party3MAGA1_combo", weight = "WEIGHT_PID_ADJ")
#'
#' # On a total variable
#' freq(data, "rxmany_Total", weight = "WEIGHT_PID_ADJ")
#'}
#' @importFrom dplyr filter group_by summarise mutate rename arrange bind_rows
#' @importFrom tibble tibble
#' @importFrom magrittr %>%
#' @importFrom rlang :=

freq <- function(data, var, weight = NULL, include_na = FALSE, digits = 0) {
  # Check variable exists
  if (!var %in% names(data)) {
    message(paste0(var, " is not in dataset"))
    return(NULL)
  }
  # Check weight exists if provided
  if (!is.null(weight) && !weight %in% names(data)) {
    message(paste0(weight, " is not in dataset"))
    return(NULL)
  }

  pct_label <- if (!is.null(weight)) "weighted_pct" else "unweighted_pct"
  # Work with factor levels if available to preserve order
  if (is.factor(data[[var]])) {
    existing_levels <- levels(data[[var]])
    if (include_na && anyNA(data[[var]])) {
      existing_levels <- c(existing_levels, "(Missing)")
    }
  } else {
    existing_levels <- NULL
  }

  # Resolve weight column before pipe chain to avoid NULL subsetting issue
  data[["._weight_"]] <- if (!is.null(weight)) data[[weight]] else rep(1, nrow(data))
  df <- data %>%
    dplyr::mutate(!!var := ifelse(is.na(.data[[var]]), "(Missing)", as.character(.data[[var]]))) %>%
    {if (!include_na) dplyr::filter(., .data[[var]] != "(Missing)") else .} %>%
    dplyr::group_by(.data[[var]]) %>%
    dplyr::summarise(
      weighted_n   = sum(.data[["._weight_"]], na.rm = TRUE),
      unweighted_n = dplyr::n()
    ) %>%
    dplyr::mutate(
      pct        = round(weighted_n / sum(weighted_n) * 100, digits),
      weighted_n = NULL
    ) %>%
    dplyr::rename(!!pct_label := pct)
  # Restore factor ordering
  if (!is.null(existing_levels)) {
    df <- df %>%
      dplyr::mutate(!!var := factor(.data[[var]], levels = existing_levels)) %>%
      dplyr::arrange(.data[[var]]) %>%
      dplyr::mutate(!!var := as.character(.data[[var]]))
  }
  total_pct <- sum(df[[pct_label]])
  total_n   <- if (include_na) nrow(data) else sum(!is.na(data[[var]]))
  df %>%
    dplyr::bind_rows(tibble::tibble(!!var := "Total", !!pct_label := total_pct, unweighted_n = total_n))
}



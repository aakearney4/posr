#' Create MOSE Summary Table
#'
#' Computes survey design summary statistics including DEFF, MOSE, and weighted sums
#' for a set of variables in a dataset. Supports total, single variables, and interactions.
#' Use for statistical testing or inspecting data
#'
#' @param data A data frame containing the variables and weight column.
#' @param weight_col Character string specifying the column in `data` with survey weights.
#' @param vars Character vector of variable names to summarize. Use "total" for overall summary.
#'             Interaction terms can be specified with a colon, e.g., "gender:party".
#' @param min_n Minimum unweighted sample size to include a row in the summary. Default is 75.
#' @param transpose Logical, whether to return a transposed table (rows = metrics, columns = variable/level). Default is TRUE.
#' @param digits Number of decimal places to round numeric summary statistics (except MOSE). Default is 2.
#' @param moedigits Number of decimal places to round MOSE. Default is 0 (whole number).
#'
#' @return A data frame or transposed matrix of summary statistics.
#' @export
#' @importFrom dplyr filter mutate select bind_rows pull
#' @importFrom rlang .data syms
#' @importFrom pewmethods calculate_deff

mose_table <- function(data, weight_col, vars, min_n = 75,
                       transpose = TRUE, digits = 2, moedigits = 0) {

# Build a summary row for a given variable/level/weights
  make_summary <- function(var_name, subsample, w) {
    n_unweighted <- length(w)
    if (n_unweighted < min_n) return(NULL)
    deff_res <- calculate_deff(w)

    tibble(
      Variable = var_name,
      Subsample = subsample,
      `Unweighted N` = n_unweighted,
      `Sum of weights` = round(sum(w), digits),
      `Sum of weights squared` = round(sum(w^2), digits),
      DEFF = round(deff_res$deff, digits),
      `Sqrt(DEFF)` = round(sqrt(deff_res$deff), digits),
      MOSE = paste0(round(deff_res$moe, moedigits), "%"),
      `Min weight` = round(min(w, na.rm = TRUE), digits),
      `Max weight` = round(max(w, na.rm = TRUE), digits),
      `Median weight` = round(median(w, na.rm = TRUE), digits)
    )
  }

# -------------------------------
# Check inputs
# -------------------------------
  if (!(weight_col %in% names(data))) stop("Weight column not found in data.")
  if (!is.numeric(data[[weight_col]])) stop("Weight column must be numeric.")
  missing_vars <- vars[!vars %in% names(data) & vars != "total" & !grepl(":", vars)]
  if (length(missing_vars) > 0) stop("These vars are missing in data: ", paste(missing_vars, collapse = ", "))

# -------------------------------
# Loop through each variable
# -------------------------------
  results <- lapply(vars, function(v) {

    if (v == "total") {
      w <- data %>% filter(!is.na(.data[[weight_col]])) %>% pull(.data[[weight_col]])
      make_summary("total", "Total", w)

    } else if (grepl(":", v)) {
      # Interaction variable
      vars_split <- strsplit(v, ":")[[1]]
      dat_sub <- data %>% filter(!is.na(.data[[weight_col]])) %>%
        mutate(.group = interaction(!!!rlang::syms(vars_split), sep = " | "))
      group_levels <- levels(dat_sub$.group)

      bind_rows(lapply(group_levels, function(level) {
        w <- dat_sub %>% filter(.group == level) %>% pull(.data[[weight_col]])
        make_summary(v, level, w)
      }))

    } else {
      # Single variable
      dat_sub <- data %>% filter(!is.na(.data[[v]]), !is.na(.data[[weight_col]]))
      group_levels <- unique(dat_sub[[v]])

      bind_rows(lapply(group_levels, function(level) {
        w <- dat_sub %>% filter(.data[[v]] == level) %>% pull(.data[[weight_col]])
        make_summary(v, level, w)
      }))
    }
  })

# -------------------------------
# Combine all results
# -------------------------------
  full_summary <- bind_rows(results)

# -------------------------------
# Transpose if requested
# -------------------------------
  if (transpose) {
    num_data <- full_summary[, -(1:2)] %>% as.data.frame()
    transposed_nums <- t(num_data)
    transposed <- rbind(
      as.character(full_summary$Variable),
      as.character(full_summary$Subsample),
      transposed_nums
    )
    colnames(transposed) <- NULL
    rownames(transposed) <- c(
      "Variable", "Subsample", "Unweighted N", "Sum of weights",
      "Sum of weights squared", "DEFF", "Sqrt(DEFF)", "MOSE",
      "Min weight", "Max weight", "Median weight"
    )
    return(transposed)
  } else {
    return(full_summary)
  }
}

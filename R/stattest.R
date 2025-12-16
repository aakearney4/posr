#' Perform Statistical Significance Testing
#'
#' Computes differences in proportions for survey data, supporting:
#' 1. Single survey with one question (horse race)
#' 2. Single survey with two groups
#' 3. Single survey with one group, two questions
#' 4. Two independent surveys
#'
#' Also calculates standard errors, margin of error, design effect (if weights are provided),
#' and flags whether differences are statistically significant.
#'
#' @param data Data frame. Primary survey data.
#' @param by Character. Column name of the grouping variable.
#' @param var Character. Column name of the variable to analyze in `data`.
#' @param group1 Character. Name of the first group level.
#' @param answer Character. Response value of interest for `var`.
#' @param wt Character, optional. Column name for survey weights.
#' @param group2 Character, optional. Name of second group level (for two-group tests).
#' @param answer2 Character, optional. Response value for `var2` (for one-group, two-question test).
#' @param var2 Character, optional. Second question variable (for one-group, two-question test).
#' @param data2 Data frame, optional. Second survey data (for independent surveys).
#' @param wt2 Character, optional. Weight column for `data2`.
#' @param crit Numeric. Critical value for significance (default 1.96 ~ 95% CI).
#' @param type Character. Type of test: `"auto"` (default), `"horse_race"`, `"two_groups"`, `"one_group_two_questions"`.
#'
#' @return A data frame summarizing:
#'   - Proportions (P1/P2)
#'   - Sample sizes
#'   - Weight sums, squared weights, design effects
#'   - Differences, tolerances, and significance
#'   - Test type
#'
#' @details
#' - Uses `pewmethods::get_totals()` internally.
#' - For weighted surveys, calculates design effects and adjusts standard errors.
#' - If `type = "auto"`, the function attempts to infer test type based on supplied arguments.
#' - Returns a tidy table that can be exported to Excel.
#'
#' @examples
#' \dontrun{
#' stattest(
#'   data = survey1,
#'   by = "gender",
#'   var = "support_candidate",
#'   group1 = "Male",
#'   group2 = "Female",
#'   answer = "Yes",
#'   wt = "weight"
#' )
#'
#' # Two independent surveys
#' stattest(
#'   data = survey1,
#'   data2 = survey2,
#'   var = "support_candidate",
#'   var2 = "support_candidate",
#'   by = "gender",
#'   group1 = "Male",
#'   group2 = "Female",
#'   answer = "Yes",
#'   answer2 = "Yes",
#'   wt = "weight",
#'   wt2 = "weight"
#' )
#' }
#'
#' @export
#' @importFrom dplyr as_tibble mutate select
#' @importFrom pewmethods get_totals
#'
#'
#'
#'
stattest <- function(
    data, by, var, group1, answer, wt = NULL,
    group2 = NULL, answer2 = NULL, var2 = NULL,
    data2 = NULL, wt2 = NULL,
    crit = 1.96, type = "auto"
) {

  # ---- Case: two independent surveys ----
  if(!is.null(data2)){
    test_name <- "Two independent surveys"
    cat("Running test:", test_name, "\n")

    # --- Survey 1 totals ---
    totals1 <- get_totals(var = var, df = data, wt = wt, by = by, digits = 2, na.rm = TRUE)
    p1 <- totals1[totals1[[var]] == answer, group1] / 100
    n1 <- nrow(data)
    sum_wt1 <- if(!is.null(wt)) sum(data[[wt]], na.rm = TRUE) else NA
    sum_wt_sq1 <- if(!is.null(wt)) sum(data[[wt]]^2, na.rm = TRUE) else NA
    deff1 <- if(!is.null(wt)) (sum_wt_sq1 / sum_wt1^2) * n1 else NA

    # --- Survey 2 totals ---
    totals2 <- get_totals(var = var2, df = data2, wt = wt2, by = by, digits = 2, na.rm = TRUE)
    p2 <- totals2[totals2[[var2]] == answer, group2] / 100
    n2 <- nrow(data2)
    sum_wt2 <- if(!is.null(wt2)) sum(data2[[wt2]], na.rm = TRUE) else NA
    sum_wt_sq2 <- if(!is.null(wt2)) sum(data2[[wt2]]^2, na.rm = TRUE) else NA
    deff2 <- if(!is.null(wt2)) (sum_wt_sq2 / sum_wt2^2) * n2 else NA

    # --- Standard error with design effects ---
    raw_tol <- sqrt(deff1 * (p1*(1-p1)/n1) + deff2 * (p2*(1-p2)/n2))
    tol <- crit * raw_tol
    diff_p <- p1 - p2
    significant <- abs(diff_p) > tol

    return(data.frame(
      P1 = p1, Unwgt_N1 = n1, Sum_wt1 = sum_wt1, Sum_wt_sq1 = sum_wt_sq1, Deff1 = deff1,
      P2 = p2, Unwgt_N2 = n2, Sum_wt2 = sum_wt2, Sum_wt_sq2 = sum_wt_sq2, Deff2 = deff2,
      Difference_P1_P2 = diff_p, SE = raw_tol, MOE_05 = tol, Significant_at_05 = significant,
      Test_Type = "two_groups_independent_surveys"
    ))
  }

  # ---- Single survey: auto-detect test type ----
  if(type == "auto"){
    if(!is.null(group2)){
      type <- "two_groups"
      test_name <- "Two groups, one question"
    } else if(!is.null(var2) & !is.null(answer2)){
      type <- "one_group_two_questions"
      test_name <- "One group, two questions"
    } else {
      type <- "horse_race"
      test_name <- "Horse race (one group, one question)"
    }
  } else {
    # User-specified test type
    test_name <- switch(type,
                        horse_race = "Horse race (one group, one question)",
                        two_groups = "Two groups, one question",
                        one_group_two_questions = "One group, two questions",
                        type)
  }

  cat("Running test:", test_name, "\n")

  # ---- Compute P1 and P2 ----
  if(type == "horse_race"){
    totals <- get_totals(var = var, df = data, wt = wt, by = by, digits = 2, na.rm = TRUE)
    p1 <- totals[totals[[var]] == answer, group1] / 100
    p2 <- 1 - p1
    n1 <- n2 <- nrow(subset(data, data[[by]] == group1))
  } else if(type == "two_groups"){
    totals <- get_totals(var = var, df = data, wt = wt, by = by, digits = 2, na.rm = TRUE)
    p1 <- totals[totals[[var]] == answer, group1] / 100
    p2 <- totals[totals[[var]] == answer, group2] / 100
    n1 <- sum(data[[by]] == group1)
    n2 <- sum(data[[by]] == group2)
  } else if(type == "one_group_two_questions"){
    totals1 <- get_totals(var = var, df = data, wt = wt, by = by, digits = 2, na.rm = TRUE)
    totals2 <- get_totals(var = var2, df = data, wt = wt, by = by, digits = 2, na.rm = TRUE)
    p1 <- totals1[totals1[[var]] == answer, group1] / 100
    p2 <- totals2[totals2[[var2]] == answer2, group1] / 100
    n1 <- n2 <- nrow(subset(data, data[[by]] == group1))
  }

  # ---- Sum of weights and design effect ----
  if(!is.null(wt)){
    sum_wt1 <- sum(data[[wt]][data[[by]] == group1], na.rm = TRUE)
    sum_wt2 <- if(!is.null(group2)) sum(data[[wt]][data[[by]] == group2], na.rm = TRUE) else sum_wt1

    sum_wt_sq1 <- sum(data[[wt]][data[[by]] == group1]^2, na.rm = TRUE)
    sum_wt_sq2 <- if(!is.null(group2)) sum(data[[wt]][data[[by]] == group2]^2, na.rm = TRUE) else sum_wt_sq1

    deff1 <- (sum_wt_sq1 / sum_wt1^2) * n1
    deff2 <- (sum_wt_sq2 / sum_wt2^2) * n2
  } else {
    sum_wt1 <- sum_wt_sq1 <- deff1 <- NA
    sum_wt2 <- sum_wt_sq2 <- deff2 <- NA
  }

  # ---- Standard error/Tolerances ----
  raw_tol <- sqrt(deff1 * (p1 * (1 - p1) / n1) + deff2 * (p2 * (1 - p2) / n2))
  tol <- crit * raw_tol
  diff_p <- abs(p1 - p2)
  diff_tol <- diff_p - tol

  significance <- ifelse(
    diff_tol > 0,
    "YES",
    ifelse(diff_tol > -0.009,
           "borderline",
           "NO")
  )

# --- Rounding helper ---
  round2 <- function(x) if(is.numeric(x)) round(x, 2) else x

# --- Output ---

  header1 <- if(!is.null(group2)) group1 else answer[1]
  header2 <- if(!is.null(group2)) group2 else if(length(answer) > 1) answer[2] else answer[1]

  out <- data.frame(
    Label = c(
      "--- Group 1 ---", "P1", "Unweighted N (Group 1)", "Sum of Weights (Group 1)",
      "Sum of Sq Weights (Group 1)","Design Effect (Group 1)","--- Group 2 ---",
      "P2", "Unweighted N (Group 2)", "Sum of Weights (Group 2)", "Sum of Sq Weights (Group 2)",
      "Design Effect (Group 2)", "--- Test Results ---", "Combined raw tolerance",
      "Tolerance at 0.05", "Difference P1 – P2","Difference minus Tolerance",
      "Significant at 0.05", "Test Type"
    ),

    Value = c(
      header1,                       # Group 1 header
      round2(p1),
      round2(n1),
      round2(sum_wt1),
      round2(sum_wt_sq1),
      round2(deff1),
      header2,                       # Group 2 header
      round2(p2),
      round2(n2),
      round2(sum_wt2),
      round2(sum_wt_sq2),
      round2(deff2),
      NA,                            # Test results header
      round2(raw_tol),
      round2(tol),
      round2(diff_p),
      round2(diff_tol),
      significance,
      test_name
    ),
    stringsAsFactors = FALSE
  )

 return(out)
}

#' Prepare Stat Testing Table for export
#'
#' Converts each row of a data frame to a list, coercing numeric-like values to numeric.
#' This preserves character values (e.g., labels, headers) while enabling consistent formatting for Excel.
#'
#' @param data A data frame, typically the output of `stattest()`.
#'
#' @return A list of lists, one per row, with numeric values coerced and character values preserved.
#'
#' @details
#' - Each row is converted to a list to allow per-cell formatting in Excel.
#' - Numeric-looking strings are converted to numeric; non-numeric strings remain unchanged.
#' - Intended for internal use with `statexporter()`.
#'
#' @examples
#' \dontrun{
#' df <- data.frame(Label = c("A", "B"), Value = c("1.23", "Yes"))
#' fix_statguts(df)
#' }
#'
#' @keywords internal
#' @importFrom dplyr as_tibble mutate select
#' @importFrom purrr map

fix_statguts <- function(data) {
  data %>%
    as_tibble(.name_repair = "minimal") %>%
    mutate(rownum = row_number()) %>%
    split(.$rownum) %>%
    map(function(x) {
      x <- select(x, -rownum)
      vals <- x[, , drop = FALSE]

      # Cell-by-cell numeric coercion
      for (j in seq_along(vals)) {
        val <- vals[[j]]
        num_val <- suppressWarnings(as.numeric(val))
        if (!is.na(num_val)) vals[[j]] <- num_val
      }

      # Return a list (preserves numeric vs character)
      as.list(vals)
    })
}

#' Export a Data Frame to Excel with Numeric Formatting
#'
#' Writes a data frame (typically output from `stattest()`) to an Excel workbook.
#' Numeric values are formatted to 2 decimal places, while character values are preserved.
#'
#' @param df A data frame to export.
#' @param file Character. File path for the Excel workbook to save.
#' @param sheet_name Character. Name of the sheet to write (default `"Sheet1"`).
#'
#' @return None. Writes an Excel workbook to the specified file path.
#'
#' @details
#' - Uses `fix_statguts()` internally to convert each row for proper formatting.
#' - Automatically applies numeric formatting (2 decimals) to numeric cells.
#' - Column widths are auto-adjusted.
#'
#' @examples
#' \dontrun{
#' res <- stattest(data = survey1, by = "gender", var = "support", group1 = "Male",
#'                 group2 = "Female", answer = "Yes", wt = "weight")
#' statexporter(res, "survey_results.xlsx")
#' }
#'
#' @export
#' @importFrom openxlsx createWorkbook addWorksheet writeData addStyle createStyle setColWidths saveWorkbook

statexporter <- function(df, file, sheet_name = "Sheet1") {

  wb <- createWorkbook()
  addWorksheet(wb, sheet_name)

  # Process all cells
  guts <- fix_statguts(df)

  for (i in seq_len(length(guts))) {
    excel_row <- i

    # Convert the list of cells back to data.frame row for writeData
    writeData(
      wb, sheet_name,
      x = as.data.frame(guts[[i]], stringsAsFactors = FALSE),
      startRow = excel_row,
      startCol = 1,
      colNames = FALSE
    )

    # Apply numeric formatting
    is_num <- vapply(guts[[i]], is.numeric, logical(1))
    if (any(is_num)) {
      addStyle(
        wb, sheet_name,
        style = createStyle(numFmt = "0.00"),
        rows = excel_row,
        cols = which(is_num),
        stack = TRUE
      )
    }
  }

  # Auto-width
  setColWidths(wb, sheet = sheet_name, cols = seq_len(ncol(df)), widths = "auto")

  saveWorkbook(wb, file, overwrite = TRUE)
}


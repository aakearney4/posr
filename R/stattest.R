library(dplyr)
library(pewmethods)

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

    totals1 <- get_totals(var = var, df = data, wt = wt, by = by, digits = 2, na.rm = TRUE)
    p1 <- totals1[totals1$Q1 == answer, group1] / 100
    n1 <- nrow(data)
    sum_wt1 <- if(!is.null(wt)) sum(data[[wt]], na.rm = TRUE) else NA
    sum_wt_sq1 <- if(!is.null(wt)) sum(data[[wt]]^2, na.rm = TRUE) else NA
    deff1 <- if(!is.null(wt)) (sum_wt_sq1 / sum_wt1^2) * n1 else NA

    totals2 <- get_totals(var = var2, df = data2, wt = wt2, by = by, digits = 2, na.rm = TRUE)
    p2 <- totals2[totals2[[var2]] == answer, group2] / 100
    n2 <- nrow(data2)
    sum_wt2 <- if(!is.null(wt2)) sum(data2[[wt2]], na.rm = TRUE) else NA
    sum_wt_sq2 <- if(!is.null(wt2)) sum(data2[[wt2]]^2, na.rm = TRUE) else NA
    deff2 <- if(!is.null(wt2)) (sum_wt_sq2 / sum_wt2^2) * n2 else NA

    se <- sqrt((p1*(1-p1)/n1) + (p2*(1-p2)/n2))
    diff_p <- p1 - p2
    moe <- crit * se
    significant <- abs(diff_p) > moe

    return(data.frame(
      P1 = p1, Unwgt_N1 = n1, Sum_wt1 = sum_wt1, Sum_wt_sq1 = sum_wt_sq1, Deff1 = deff1,
      P2 = p2, Unwgt_N2 = n2, Sum_wt2 = sum_wt2, Sum_wt_sq2 = sum_wt_sq2, Deff2 = deff2,
      Difference_P1_P2 = diff_p, SE = se, MOE_05 = moe, Significant_at_05 = significant,
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
    n1 <- n2 <- sum(data[[by]] == group1)
  } else if(type == "two_groups"){
    totals <- get_totals(var = var, df = data, wt = wt, by = by, digits = 2, na.rm = TRUE)
    p1 <- totals[totals[[var]] == answer, group1] / 100
    p2 <- totals[totals[[var]] == answer, group2] / 100
    n1 <- sum(data[[by]] == group1)
    n2 <- sum(data[[by]] == group2)
  } else if(type == "one_group_two_questions"){
    totals1 <- get_totals(var = var, df = data, wt = wt, by = by, digits = 2, na.rm = TRUE)
    totals2 <- get_totals(var = var2, df = data, wt = wt, by = by, digits = 2, na.rm = TRUE)
    p1 <- totals1[totals1$Q1 == answer, group1] / 100
    p2 <- totals2[totals2[[var2]] == answer2, group1] / 100
    n1 <- n2 <- sum(data[[by]] == group1)
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

  # ---- Standard error ----
  se <- switch(type,
               horse_race = sqrt((p1 + p2 - (p1 - p2)^2)/(n1 - 1)),
               two_groups = sqrt((p1*(1-p1)/n1) + (p2*(1-p2)/n2)),
               one_group_two_questions = sqrt((p1*(1-p1)/n1) + (p2*(1-p2)/n1))
  )

  # ---- Combined Raw Tolerance ----
  raw_tol <- sqrt(deff1 * (p1 * (1 - p1) / n1) + deff2 * (p2 * (1 - p2) / n2))

  # ---- Tolerance at 0.05 ----
  tol <- crit * raw_tol

  # ---- Differences and significance ----
  diff_p <- abs(p1 - p2)
  diff_tol <- diff_p - tol

  significance <- ifelse(
    diff_tol > 0,
    "YES",
    ifelse(diff_tol > -0.009,
           "borderline",
           "NO")
  )

  # ---- Return results ----
  # Helper to round only numerics
  round2 <- function(x) {
    if (is.numeric(x)) round(x, 2) else x
  }

  # Helper to round numeric values to 2 decimals but leave non-numeric as-is
  round2 <- function(x) {
    if(is.numeric(x)) round(x, 2) else x
  }

  # Determine the headers based on whether group2 exists
  header1 <- if(!is.null(group2)) group1 else answer[1]
  header2 <- if(!is.null(group2)) group2 else if(length(answer) > 1) answer[2] else answer[1]

  out <- data.frame(
    Label = c(
      "--- Group 1 ---",
      "P1",
      "Unweighted N (Group 1)",
      "Sum of Weights (Group 1)",
      "Sum of Sq Weights (Group 1)",
      "Design Effect (Group 1)",
      "--- Group 2 ---",
      "P2",
      "Unweighted N (Group 2)",
      "Sum of Weights (Group 2)",
      "Sum of Sq Weights (Group 2)",
      "Design Effect (Group 2)",
      "--- Test Results ---",
      "Combined raw tolerance",
      "Tolerance at 0.05",
      "Difference P1 – P2",
      "Difference minus Tolerance",
      "Significant at 0.05",
      "Test Type"
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

  out


  return(out)
}

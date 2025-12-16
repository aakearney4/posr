#' fix_guts
#' Internal helper function to process data frame so numbers remain
#' as numbers and text as text when exporting. Borrowed and adapted from:
#' <https://github.com/jjmoncus/wink/blob/main/R/write_banners.R>
#'
#' @param data A data frame or tibble to process.
#' @return A list of row vectors, with numeric values as numeric and text as character.
#' @keywords internal
#' @importFrom purrr map
#' @importFrom dplyr select mutate row_number
#' @importFrom tibble as_tibble

# ---- fix_guts: preserve numeric vs text safely ----
fix_guts <- function(data) {
  data %>%
    as_tibble(.name_repair = "minimal") %>%
    mutate(rownum = row_number()) %>%
    split(.$rownum) %>%
    map(function(x) {
      x <- select(x, -rownum)

      # Drop label column
      vals <- x[, -1, drop = FALSE]

      # Attempt numeric coercion
      num_try <- suppressWarnings(map(vals, as.numeric))

      # If first value converts, treat row as numeric
      if (!is.na(num_try[[1]][1])) {
        vals[] <- num_try
      }

      unlist(vals, use.names = FALSE)
    })
}

#' demo_spans
#' Internal helper function to find consecutive repeated values in a row.
#' Used to merge demo cells in the Excel sheet.
#'
#' @param demo_row A character vector representing a row of demo labels.
#' @return A tib ble with columns: label, start, end.
#' @keywords internal
#' @importFrom tibble tibble

# ---- helper to find duplicate demo label spans for merging ----
demo_spans <- function(demo_row) {
  r <- rle(demo_row)
  ends <- cumsum(r$lengths)
  starts <- ends - r$lengths + 1

  tibble(
    label = r$values,
    start = starts,
    end = ends
  )
}

#' excelmixer
#'
#' Export a data frame to Excel with mixed character and numeric values
#'
#' @param df Data frame to export.
#' @param file Output Excel file path.
#' @param sheet_name Name of the Excel sheet. Default is "Sheet1".
#' @param merge_demos Logical, should demo cells in the first row be merged? Default TRUE.
#' @return None; writes Excel file.
#' @export
#' @importFrom openxlsx createWorkbook addWorksheet writeData saveWorkbook setColWidths createStyle addStyle mergeCells
#' @importFrom dplyr select
#' @importFrom tibble rownames_to_column
#' @importFrom purrr map


# ---- Excel exporter ----
excelmixer <- function(df, file, sheet_name = "Sheet1", merge_labels = TRUE) {

  wb <- createWorkbook()
  addWorksheet(wb, sheet_name)

  # Move row names into data
  df <- tibble::rownames_to_column(df, "Row")

  # Demographic name and label rows
  demo_var <- as.character(df[1, -1])
  demo_val <- as.character(df[2, -1])
  body <- df[-c(1, 2), ]

  # ---- write demographic headers ----
  writeData(
    wb, sheet_name,
    x = as.data.frame(t(c("", demo_var))),
    startRow = 1, startCol = 1,
    colNames = FALSE
  )

  writeData(
    wb, sheet_name,
    x = as.data.frame(t(c("", demo_val))),
    startRow = 2, startCol = 1,
    colNames = FALSE
  )

  # ---- merge demographic headers when they match (e.g. PARTY PARTY) default=TRUE ----
  if (merge_labels) {
    spans <- demo_spans(demo_var)
    for (i in seq_len(nrow(spans))) {
      if (!is.na(spans$label[i]) && spans$start[i] < spans$end[i]) {
        mergeCells(
          wb,
          sheet = sheet_name,
          cols = (spans$start[i] + 1):(spans$end[i] + 1),
          rows = 1
        )
      }
    }
  }

  # ---- header style ----
  header_style <- createStyle(
    halign = "center",
    valign = "center"
  )
  addStyle(
    wb, sheet_name,
    style = header_style,
    rows = 1:2,
    cols = 1:ncol(df),
    gridExpand = TRUE,
    stack = TRUE
  )

  # ---- body ----
  guts <- fix_guts(body)
  stopifnot(length(guts) == nrow(body))

  for (i in seq_len(nrow(body))) {

    excel_row <- i + 2

    # Row label
    writeData(
      wb, sheet_name,
      x = data.frame(body$Row[i]),
      startRow = excel_row,
      startCol = 1,
      colNames = FALSE
    )

    # Row values
    row_values <- guts[[i]]

    # Write values as-is
    writeData(
      wb, sheet_name,
      x = as.data.frame(t(row_values)),
      startRow = excel_row,
      startCol = 2,
      colNames = FALSE
    )

    # Apply numeric style (without forcing digits)
    is_num <- vapply(row_values, is.numeric, logical(1))
    if (any(is_num)) {
      addStyle(
        wb, sheet_name,
        style = createStyle(),
        rows = excel_row,
        cols = which(is_num) + 1,
        stack = TRUE
      )
    }
  }

  # ---- auto-adjust first column width ----
  setColWidths(
    wb,
    sheet = sheet_name,
    cols = 1,
    widths = "auto"
  )

  saveWorkbook(wb, file, overwrite = TRUE)
}


# # Example usage:
# df <- data.frame(
#   Label = c("A", "B", "C", "D"),
#   Value = c(1.23, 45, NA, "some text"),
#   stringsAsFactors = FALSE
# )
#
# excelmixer(df, "mixed_example.xlsx")

library(openxlsx)
library(dplyr)
library(purrr)

# Keep your fix guts function from before
fix_guts <- function(data) {
  data %>%
    mutate(rownum = 1:nrow(data)) %>%
    split(~rownum) %>%
    map(function(x) {
      x <- x %>% select(-rownum)
      rest <- x[2:length(x)]
      condition <- suppressWarnings(as.numeric(rest[[1]]) %>% is.na())
      if (!condition) {
        rest <- rest %>% mutate(across(everything(), as.numeric))
      }
      rest
    }) %>%
    list_flatten()
}

#Write a dataframe to excel that has mixed styles
excelmixer <- function(df, file, sheet_name = "Sheet1") {
  wb <- createWorkbook()
  addWorksheet(wb, sheet_name)

  # Create a column for row names if they exist
  if (!is.null(rownames(df))) {
    df <- cbind(Row = rownames(df), df)
  }

  # Write headers
  writeData(wb, sheet_name, x = colnames(df), startRow = 1, startCol = 1, colNames = FALSE)

  # Apply fix_guts to split rows and handle numeric/text
  guts <- fix_guts(df)

  # Write each row iteratively
  for (i in 1:nrow(df)) {
    # Write main data (starting from column 2)
    writeData(
      wb,
      sheet = sheet_name,
      x = guts[[i]],
      startRow = i + 1,  # row 1 is headers
      startCol = 2,
      colNames = FALSE
    )

    # Write first column separately (labels/text)
    writeData(
      wb,
      sheet = sheet_name,
      x = df[i, 1, drop = TRUE],
      startRow = i + 1,
      startCol = 1,
      colNames = FALSE
    )

    # Apply numeric formatting only to numeric cells
    numeric_cells <- sapply(guts[[i]], is.numeric)
    if(any(numeric_cells)) {
      addStyle(
        wb,
        sheet = sheet_name,
        style = createStyle(numFmt = "0.00"),
        rows = i + 1,
        cols = which(numeric_cells) + 1,  # +1 because main data starts in col 2
        gridExpand = TRUE,
        stack = TRUE
      )
    }
  }

  # Save workbook
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

#' Main Execution
#'
#' This script loads data, performs statistical tests, and writes results to an Excel file.
file_name <- 'asdf.xlsx'  # Specify the Excel file name
sheet_name <- "Stratified Testing"  # Specify the sheet name
df <- load_data(file_name, sheet_name)  # Load data

# Define unique identifiers for analysis
unique_values <- c(10, 11, 12, 13, 14, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 
                   56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 
                   72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84)

data_list <- prepare_data(df, unique_values)  # Prepare data

wb <- loadWorkbook(file_name)  # Load the workbook

for (j in unique_values) {
  dff <- data_list[[as.character(j)]]  # Subset data for current unique identifier
  test_result <- perform_breslow_day_test(dff)  # Perform Breslow-Day test
  
  if (!is.null(test_result)) {
    write_results(wb, df, c(test_result$p.value), j)  # Write results if test was performed
  } else {
    p_values <- perform_chi_square_test(dff)  # Perform Chi-square test
    write_results(wb, df, p_values, j)  # Write results
  }
}

# Save the workbook
saveWorkbook(wb, file_name, overwrite = TRUE)  # Save changes to Excel file

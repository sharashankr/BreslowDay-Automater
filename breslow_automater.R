#' Load necessary libraries and load the Excel workbook
#'
#' @param file_name A character string specifying the name of the Excel file.
#' @param sheet_name A character string specifying the name of the sheet to read.
#' @return A data frame containing the data from the specified sheet.
load_data <- function(file_name, sheet_name) {
  wb <- loadWorkbook(file_name)  # Load the workbook
  df <- read.xlsx(wb, sheet = sheet_name)  # Read data from specified sheet
  return(as.data.frame(df))  # Return data as a data frame
}

#' Prepare and subset data based on unique identifiers
#'
#' @param df A data frame containing the data to be prepared.
#' @param unique_values A vector of unique identifiers used for subsetting.
#' @return A list of data frames corresponding to the unique identifiers.
prepare_data <- function(df, unique_values) {
  results <- list()
  
  for (j in unique_values) {
    matching_rows <- which(df$QN == j)  # Identify rows matching the unique identifier
    rows_to_include <- unique(c(matching_rows, matching_rows + 1, matching_rows + 2))
    dff <- df[rows_to_include[rows_to_include <= nrow(df)], ]  # Subset the data frame
    results[[as.character(j)]] <- dff  # Store result
  }
  
  return(results)  # Return a list of data frames
}

#' Perform the Breslow-Day test
#'
#' @param dff A data frame containing the data to be tested.
#' @return The result of the Breslow-Day test or NULL if no data is available.
perform_breslow_day_test <- function(dff) {
  my_list <- c()  # Initialize list to store matrices
  
  for (i in 1:nrow(dff)) {
    if (i > 1) {  
      a_size <- dff[i, 8]   
      cohort_a <- dff[i, 6] 
      b_size <- dff[i, 12]  
      cohort_b <- dff[i, 10] 
      l <- matrix(c(a_size, cohort_a - a_size, b_size, cohort_b - b_size), nrow = 2, ncol = 2)  # Create 2x2 matrix
      my_list <- append(my_list, l)  # Append matrix to list
    }
  }
  
  if (length(my_list) > 0) {
    array_3d <- array(my_list, dim = c(2, 2, nrow(dff)))  # Create 3D array
    
    # Perform the test based on the minimum values in the array
    correct_option <- ifelse(min(array_3d) <= 7, TRUE, FALSE)
    test_result <- BreslowDayTest(x = array_3d, OR = 1, correct = correct_option)  # Run the test
    
    return(test_result)  # Return test result
  }
  
  return(NULL)  # Return NULL if no data
}

#' Perform Chi-square test
#'
#' @param dff A data frame containing the data to be tested.
#' @return A vector of p-values from the Chi-square tests.
perform_chi_square_test <- function(dff) {
  p_values <- numeric(nrow(dff))  # Initialize vector for p-values
  
  for (i in 1:nrow(dff)) {
    a_size <- dff[i, 8]   
    cohort_a <- dff[i, 6] 
    b_size <- dff[i, 12]  
    cohort_b <- dff[i, 10] 
    
    observed <- matrix(c(a_size, cohort_a - a_size, b_size, cohort_b - b_size), nrow = 2, byrow = TRUE)  # Create observed matrix
    val <- min(observed)  # Find minimum value
    
    # Determine whether to apply continuity correction
    if (val <= 7) {
      test_result <- chisq.test(observed, correct = TRUE)  # Perform test with correction
    } else {
      test_result <- chisq.test(observed, correct = FALSE)  # Perform test without correction
    }
    
    p_values[i] <- test_result$p.value  # Store p-value
  }
  
  return(p_values)  # Return p-values
}

#' Write results back to Excel
#'
#' @param wb The workbook object.
#' @param df The original data frame.
#' @param p_values A vector of p-values to write to Excel.
#' @param j The unique identifier for the current analysis.
write_results <- function(wb, df, p_values, j) {
  row_indices <- which(df$QN == j)  # Find relevant rows
  for (row in row_indices + 1) {
    p_value_t <- ifelse(p_values[row - row_indices[1]] < 0.01, '<0.01', round(p_values[row - row_indices[1]], 4))  # Format p-value
    
    writeData(wb, sheet = "Stratified Testing", x = p_value_t, startRow = row, startCol = 13)  # Write p-value to Excel
    
    # Apply font and alignment styles
    if (p_values[row - row_indices[1]] > 0.05) {
      style <- createStyle(fgFill = "#90EE90", halign = "center", fontSize = 12, border = "TopBottomLeftRight")  # Define style
      addStyle(wb, sheet = "Stratified Testing", style = style, rows = row, cols = 13, gridExpand = TRUE)  # Apply style
    }
  }
}

# BreslowDay Automater

**BreslowDay Automater** is an R package designed to automate the analysis of stratified data using the Breslow-Day test and Chi-square test. It is particularly useful for evaluating the alignment of perceptions between healthcare professionals (HCPs) and patients regarding treatment-related side effects.

## Table of Contents

- [Purpose](#purpose)
- [Process Flow](#process-flow)
- [Key Functions](#key-functions)
- [Usage Example](#usage-example)
- [Installation](#installation)
- [Contributing](#contributing)
- [License](#license)

## Purpose

The purpose of this package is to streamline the process of performing statistical tests on survey data, providing insights into the homogeneity of odds ratios and the significance of observed frequencies across different strata.

## Process Flow

1. **Loading Required Libraries**:
   - The package utilizes several libraries (`DescTools`, `openxlsx`, and `dplyr`) for statistical tests and Excel file manipulation.

2. **Loading Data**:
   - The `load_data()` function loads an Excel workbook and reads data from a specified sheet, converting it into a dataframe for manipulation.

3. **Preparing Data**:
   - The `prepare_data()` function identifies unique values in the `QN` column and extracts relevant rows to create subsets for analysis.

4. **Performing Statistical Tests**:
   - The package includes two key tests:
     - **Breslow-Day Test**: Assesses the homogeneity of odds ratios across strata.
     - **Chi-square Test**: Evaluates the significance of observed frequencies.

5. **Writing Results**:
   - The `write_results()` function outputs p-values back to the Excel sheet, applying conditional formatting based on significance levels.

6. **Main Execution**:
   - Orchestrates the loading of data, preparation, statistical testing, and writing of results to the Excel file.

## Key Functions

- **`load_data(file_name, sheet_name)`**: Loads data from an Excel file and returns a dataframe.
- **`prepare_data(df, unique_values)`**: Prepares and subsets data based on unique identifiers, returning a list of dataframes.
- **`perform_breslow_day_test(dff)`**: Conducts the Breslow-Day test on the provided subset of data.
- **`perform_chi_square_test(dff)`**: Executes the Chi-square test for each row in the dataframe and returns p-values.
- **`write_results(wb, df, p_values, j)`**: Writes p-values back to the Excel sheet with appropriate formatting.

## Usage Example

Here’s how to use the **BreslowDay Automater** package:

## Input format:
##3 Stratified Testing

| Section | Survey Q | QN | Claim | Segment A | N-size | A% | A# | Segment B | N-size | B% | B# | p-value |
|---------|----------|----|-------|------------|--------|----|----|------------|--------|----|----|---------|


```r
# Load the BreslowDay Automater package
library(BreslowDayAutomater)

# Define the file and sheet names
file_name <- 'asdf.xlsx'
sheet_name <- "Stratified Testing"

# Load data
df <- load_data(file_name, sheet_name)

# Define unique QN values
unique_values <- unique(df$QN)

# Prepare data
data_list <- prepare_data(df, unique_values)

# Load workbook for writing results
wb <- loadWorkbook(file_name)

# Execute the statistical tests and write results
for (j in unique_values) {
  dff <- data_list[[as.character(j)]]
  
  # Perform Breslow-Day Test
  test_result <- perform_breslow_day_test(dff)
  
  if (!is.null(test_result)) {
    write_results(wb, df, c(test_result$p.value), j)
  } else {
    p_values <- perform_chi_square_test(dff)
    write_results(wb, df, p_values, j)
  }
}

# Save the updated workbook
saveWorkbook(wb, file_name, overwrite = TRUE)
```
## Installation

To install the **ChiSquareAutomator** package, ensure you have the required libraries (`DescTools`, `openxlsx`, `dplyr`) installed in your R environment. You can install these libraries using the following commands:

```r
install.packages("DescTools")
install.packages("openxlsx")
install.packages("dplyr")
```
## Contributing

Contributions are welcome! If you have suggestions or improvements, feel free to submit a pull request or open an issue.

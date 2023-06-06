---
title: Excel Formulas Reference Sheet
date: 2023-06-06T12:15:31+05:30
draft: false
showToc: true
---

## Arithmetic Formulas

- **Addition:** `=A1 + B1`
- **Subtraction:** `=A1 - B1`
- **Multiplication:** `=A1 * B1`
- **Division:** `=A1 / B1`
- **ROUND:** `=ROUND(A1, 2)` (Rounds the value in cell A1 to 2 decimal places)

## Common Functions

- **SUM:** `=SUM(A1:A5)` (Adds the values in cells A1 to A5)
- **AVERAGE:** `=AVERAGE(A1:A5)` (Calculates the average of values in cells A1 to A5)
- **COUNT:** `=COUNT(A1:A5)` (Counts the number of cells with values in the range A1 to A5)
- **MAX:** `=MAX(A1:A5)` (Returns the maximum value in the range A1 to A5)
- **MIN:** `=MIN(A1:A5)` (Returns the minimum value in the range A1 to A5)

## Text Functions
- **CONCATENATE:** `=CONCATENATE(A1, " ", B1)` (Joins the text in cells A1 and B1 with a space in between)
- **UPPER:** `=UPPER(A1)` (Converts the text in cell A1 to uppercase)
- **LOWER:** `=LOWER(A1)` (Converts the text in cell A1 to lowercase)
- **PROPER:** `=PROPER(A1)` (Capitalizes the first letter of each word in the text in cell A1)
- **TRIM:** `=TRIM(A1)` (Removes extra spaces from the text in cell A1)
- **SUBSTITUTE:** `=SUBSTITUTE(A1, "old", "new")` (Replaces occurrences of "old" with "new" in the text in cell A1)
- **LEN:** `=LEN(A1)` (Returns the number of characters in the text in cell A1)
- **LEFT:** `=LEFT(A1, 3)` (Returns the leftmost 3 characters from the text in cell A1)
- **RIGHT:** `=RIGHT(A1, 3)` (Returns the rightmost 3 characters from the text in cell A1)
- **MID:** `=MID(A1, 2, 5)` (Returns a substring from the text in cell A1, starting from the 2nd character and extracting 5 characters)

## Logical Formulas

- **IF:** `=IF(A1 > B1, "Yes", "No")` (Checks if A1 is greater than B1 and returns "Yes" if true, "No" if false)
- **AND:** `=AND(A1 > 5, B1 < 10)` (Checks if both A1 is greater than 5 and B1 is less than 10)
- **OR:** `=OR(A1 > 5, B1 < 10)` (Checks if either A1 is greater than 5 or B1 is less than 10)
- **NOT:** `=NOT(A1 > 5)` (Checks if A1 is not greater than 5)
- **COUNTIF:** `=COUNTIF(A1:A5, ">10")` (Counts the number of cells in the range A1 to A5 that are greater than 10)
- **SUMIF:** `=SUMIF(A1:A5, ">10", B1:B5)` (Adds the corresponding values in the range B1 to B5 for cells in the range A1 to A5 that are greater than 10)
- **AVERAGEIF:** `=AVERAGEIF(A1:A5, ">10", B1:B5)` (Calculates the average of values in the range B1 to B5 if the corresponding value in the range A1 to A5 is greater than 10)
- **IFBLANK:** `=IF(A1="", "Blank", "Not Blank")` (Checks if cell A1 is blank and returns "Blank" if true, "Not Blank" if false)

## Date and Time Functions

- **TODAY:** `=TODAY()` (Returns the current date)
- **NOW:** `=NOW()` (Returns the current date and time)
- **DATEDIF:** `=DATEDIF(A1, B1, "d")` (Calculates the number of days between dates in cells A1 and B1)

## Lookup and Reference Functions

- **VLOOKUP:** `=VLOOKUP(A1, A2:B10, 2, FALSE)` (Searches for a value in column A and returns the corresponding value from column B)
- **HLOOKUP:** `=HLOOKUP(A1, A2:F5, 3, FALSE)` (Searches for a value in row 1 and returns the corresponding value from row 3)
- **INDEX:** `=INDEX(A1:A10, 3)` (Returns the value at the third position in the range A1 to A10)
- **MATCH:** `=MATCH(A1, A1:A10, 0)` (Finds the position of a value in the range A1 to A10)

## Sort Data

1. **Filter:** Use the filter function to display only specific data based on criteria.
   - Select the range of data you want to filter.
   - Go to the Data tab and click on the Filter button.
   - Click on the filter dropdown arrows to select and display specific data.

2. **Sort:** Arrange data in a specific order based on one or more columns.
   - Select the range of data you want to sort.
   - Go to the Data tab and click on the Sort button.
   - Choose the column(s) you want to sort by and specify the sorting order (ascending or descending).

3. **Custom Sort:** Sort data based on custom criteria.
   - Select the range of data you want to sort.
   - Go to the Data tab and click on the Sort button.
   - In the Sort dialog box, click on the **Add Level** button to add additional sorting criteria.
   - Specify the column(s) and sorting order for each level.

## Miscellaneous Functions

1. **Text to Columns:** Divide data in a single column into separate columns based on a delimiter.
   - Select the range of cells containing the text you want to split.
   - Go to the Data tab and click on the Text to Columns button.
   - Choose the **Delimited** option and click Next.
   - Select the delimiter that separates the data (e.g., comma, space, tab).
   - Choose the destination where you want the split data to be placed (e.g., new columns or existing columns).
   - Click Finish to split the text into separate columns.

2. **Data Validation:** Set rules to validate and control data entry in specific cells.
   - Select the cell(s) you want to create a rule for.
   - Go to the Data tab and click on **Data Validation**.
   - Set the validation criteria, such as whole numbers, decimal numbers, dates, etc.

3. **Transpose Data:** Quickly change the orientation of data using the paste option.
   - Copy the data you want to transpose.
   - Right-click on the cell where you want to paste the transposed data.
   - Choose the **Transpose** option under **Paste Options**.

4. **Multi-Sheet Formatting:** To format multiple sheets simultaneously:
  - Press **CTRL + Click** on each sheet's tab that you want to format together.
  - Make your formatting changes in one sheet, and they will be reflected in all the selected sheets.
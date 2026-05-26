---
title: "Introduction to Excel Assignment"
date: 2023-06-06T12:45:31+05:30
draft: false
showToc: true
type: "course"
language: "bn"
layout: "assignment"
weight: 91
---

## Exercise 1: Creating a Spreadsheet

*To be completed after Video 2*

1. Create a new Excel file and save it on your device with the file name **Excel Practice**.

2. Enter the following data in your Excel sheet:

| S. No. | Name | State | Designation | Remarks | Salary | Sick Leaves |
|--------|------|-------|-------------|---------|--------|-------------|
| 1 | John | Delhi | Manager | John is a very hardworking fellow | 30,000 | 2 |
| 2 | Maria | Goa | Head Clerk | | 25,000 | 3 |
| 3 | Sukumar | West Bengal | Peon | | 15,000 | 0 |
| 4 | Shivam | Uttar Pradesh | Manager | | 30,000 | 5 |

3. Ensure the following:
   - The font size of the title row is 11, the other rows are size 10
   - The first row with the titles is bolded
   - The names are in italics
   - The words 'West Bengal' and 'Uttar Pradesh' are entirely visible
   - The cell that has remarks about John is completely visible
   - The four states are in four different colours namely, red, blue, green and yellow
   - The word 'Peon' is in orange colour

4. Press **Ctrl + S** to save the file. Another way to save it is:
   - Click on **File** (top left on your screen)
   - Select option **Save As**
   - Select the destination where you would like to save the file (it is preferred that you save it on your desktop for ease of access)
   - Name the file as **Excel Practice**
   - Click **Save**

5. You can now close the file. Remember to close the file only after saving otherwise all your changes will be lost.


## Exercise 2: Working with Multiple Sheets

*To be completed after Video 3*

1. Open your file called **Excel Practice**. You should be able to see the data that you had created in Exercise 1.

2. Create 2 new sheets in this file.

3. Rename your data sheet as **Data** and the other two sheets as **Backup** and **Graphs**.

4. Set the order of the sheets as **Backup** followed by **Data** followed by **Graphs**.

5. Paste the data from the **Data** sheet to the **Backup** sheet.

6. Create a copy of the sheet called **Data** and move it to the end. This sheet would currently be called as **Data (2)**.

7. Rename the **Data (2)** sheet as **Backup of Data**.

8. Delete **Backup** sheet. Remember that when you delete a sheet, the data can never be retrieved. Hence, be careful when deleting a sheet.


## Exercise 3: Copying and Pasting

*To be completed after Video 4*

1. In the file **Excel Practice**, go to the sheet called **Data** and select the entire data by dragging your mouse. Then right click and select **Copy** (keyboard shortcut **Ctrl + C**).

2. Create a new sheet and **Paste Values**.

3. Create another new sheet and **Paste Transpose**.

4. Go to the **Data** sheet. In cell **E6** (i.e. the row at the end of your table and in the column which has 'Remarks'), write **Total**.

5. Bold the **Total** cell and highlight it in yellow colour.

6. Sum the numbers in columns F and G.

7. Paint format from cell E6 to the cells F6 and G6. You must now see that cells F6 and G6 are also yellow and bold.


## Exercise 4: Formatting Spreadsheets

*To be completed after Video 5*

1. In the file **Excel Practice**, go to the sheet called **Data** and clear the formatting of the entire data.

2. Select Row 1 and make it bold.

3. Select the cells which have the total sum and the cell which have the word 'Total'. Select all three cells together by dragging your mouse.

4. Make the three cells bold and highlight them in grey colour.

5. Select all the cells with the data (rows 1–6 and columns A–G) and add borders to the cells.

6. Format the data in the salary column into Indian Rupee Currency and format it to one decimal place.

7. Using conditional formatting, highlight the cells which have sick leaves to be more than 3 in **RED**. Remember, do not highlight the cell that has the total.

8. Similarly, highlight the cells that have zero sick leaves in **GREEN**.


## Exercise 5: Formatting Data as a Table

*To be completed after Video 6*

1. In the file **Excel Practice**, go to the sheet called **Data** and insert rows above the row that has 'Total'. To do this, right-click where it shows the row number as '6' and click on **Insert**. Repeat this step five times. You should have 5 new empty rows above the 'Total' row.

2. Enter the following data into the five new rows:

| S. No. | Name | State | Designation | Remarks | Salary | Sick Leaves |
|--------|------|-------|-------------|---------|--------|-------------|
| 5 | Saransh | Maharashtra | Inspector | | 17,000 | 0 |
| 6 | Priya | Tamil Nadu | Analyst | | 24,000 | 1 |
| 7 | Gurpreet | Punjab | Analyst | | 27,000 | 3 |
| 8 | Mehak | Haryana | Researcher | | 32,000 | 6 |
| 9 | Robert | Manipur | Manager | | 30,000 | 1 |

3. You should now see the totals as Rs. 2,30,000 for the Salary column and as 21 for the Sick Leaves column.

4. Move the Remarks column to the end and move the cell with 'Total' below the designation column.

5. Select your data and format it as a **Table**.

6. Play with the filters in each column. Example, filter out the people who have '0' sick leaves. Or filter out the people who have their designation as 'Manager'.

7. Add a new total row at the end of your data and in that row, select **Sum** from the drop down menu for the 'Sick Leaves' and 'Salary' columns.

8. Delete the earlier row with the 'Total'.


## Exercise 6: Sorting, Filtering and Data Validation

*To be completed after Video 7*

1. In the file **Excel Practice**, go to the sheet called **Data** and select your table and convert it to a **Range**.

2. Select your data again and add a filter.

3. Sort the **Sick Leaves** column from smallest to largest.

4. Sort the **Name** column alphabetically.

5. Add data validation to the **Sick Leaves** column where the column can only take numbers between 0 and 31.

6. Try entering 'NA' in the sick leaves column — this should show an error. If this does not show an error, add the correct data validation again.

7. Add data validation to the **Remarks** column where the column can only take text with characters greater than 10 and less than 100.

8. Check if the data validation works.

9. Convert the range into a table again and check if the data validation is still working.


## Exercise 7: Using Basic Formulas and Functions

*To be completed after Video 8*

1. In the file **Excel Practice**, go to the sheet called **Data** and calculate the following for the **Salary** column:

   - SUM
   - AVERAGE
   - COUNT
   - MAX
   - MIN

2. Once complete, the output should be the following:

| Function | Result |
|----------|--------|
| SUM | 2,30,000 |
| AVERAGE | 25,555.55 |
| COUNT | 9 |
| MAX | 32,000 |
| MIN | 15,000 |

3. Now copy the formulas from the **Salary** column to the **Sick Leaves** column. You should see the following:

| Function | Result |
|----------|--------|
| SUM | 21 |
| AVERAGE | 2.33 |
| COUNT | 9 |
| MAX | 6 |
| MIN | 0 |

4. Make sure that the formatting of the **Salary** column does not show up in the **Sick Leaves** column. In case the formatting shows up, clear the formatting.

5. Round off the 'average' number in the **Sick Leaves** column by using the formula **ROUND** along with the **AVERAGE** formula. Round it off to zero decimal places.


## Exercise 8.1: Manipulating Text — CONCATENATE

*To be completed after Video 9*

1. In the file **Excel Practice**, go to the sheet called **Data**. Sort the **Name** column alphabetically.

2. Rename the title of the **Name** column to **First Name**.

3. Insert a new column before the **State** column and enter the following data:

| Last Name |
|-----------|
| Singh |
| Biju |
| Gonzalves |
| Kapoor |
| Malhotra |
| Mathew |
| Tripathi |
| Kumar |
| Ray |

4. Insert another new column before the **State** column. Call it **Full Name**.

5. In the **Full Name** column, use the **CONCATENATE** formula to combine the first and the last name together. Make sure the formatting does not change.

6. Now, delete the formula from the **Full Name** column, and use the **&** operator to combine the first and the last name together. Make sure the formatting does not change.


### Exercise 8.2: Manipulating Text — Text Functions

*To be completed after Exercise 8.1*

1. In the file **Excel Practice**, create a new sheet and name it **Rough**.

2. Go to the sheet called **Data**. Select and copy the **Full Name** column. Paste values in the first column of the **Rough** sheet.

3. Put the titles of column B and column C as **First Name** and **Last Name** respectively.

4. Using **Text to Columns**, break down the names in column A into their first name and last name columns. Make sure that column A has the Full Name, column B has the first name and column C has the last name.

5. Try out and play around with the formulas: **TRIM**, **UPPER**, **LOWER**, **PROPER**.

6. Test out the length of the names of all the people using the **LEN** function.

7. Using the **Full Name** column as the reference, in a new column, replace the blank space between the first name and the last name with a hyphen. The names should now look like:

| Hyphenated Names |
|------------------|
| Gurpreet-Singh |
| John-Biju |
| Maria-Gonzalves |
| Mehak-Kapoor |
| Priya-Malhotra |
| Robert-Mathew |
| Saransh-Tripathi |
| Shivam-Kumar |
| Sukumar-Ray |


## Exercise 9: Good Practices

*To be completed after Video 10*

1. Create a new sheet with the data that was shown in the video.

2. Convert this data into a table.

3. Give the table a suitable title.

4. **Bonus:** Create a custom list of all the districts in West Bengal as shown in the course.

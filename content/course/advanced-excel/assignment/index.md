---
title: "Advanced Excel Functions Assignment"
date: 2023-06-06T12:45:31+05:30
draft: false
showToc: true
type: "course"
layout: "assignment"
weight: 91
---

Download the worksheet and demo sheet to get started:

- [Advanced Excel Course — Work Sheet](/Advanced%20Excel%20Course%20-%20Work%20Sheet.xlsx)
- [Advanced Excel Course — Demo Sheet](/Advanced%20Excel%20Course%20-%20Demo%20Sheet.xlsx)


## Exercise 1: Data Validation

*To be completed after Video 1*

1. Download the file: [Advanced Excel Course — Work Sheet](/Advanced%20Excel%20Course%20-%20Work%20Sheet.xlsx)

2. Open the file and go to the sheet called **Data Validation**.

3. For the **District** column (without the header row), add a data validation (list type) to allow only the following value: **Kalimpong**.

4. For the **Block** column (without the header row), add a data validation (list type) to allow only the following blocks: **Kalimpong - 1**, **Kalimpong - 2**, **Gorubathan**, and **Kalimpong Municipality**.

5. For the **UDISE Code** column (without the header row), add a data validation to allow only whole numbers between **19240100000** and **19240409999**.

6. For the **Management** column (without the header row), add a data validation (list type) to allow only the following values: **Department of Education**, **Government Sponsored**, and **Madrasa Recognized**.

7. For the **Highest Class** column (without the header row), add a data validation (list type) to allow only the following values: **Class 12**, **Class 10**, **Class 8**, **Class 5**, and **Class 4**.

8. For the **Lowest Class** column (without the header row), add a data validation (list type) to allow only the following options: **Pre-Primary**, **Class 1**, **Class 5**, **Class 6**, **Class 9**, and **Class 11**.

9. For the remaining columns (column I to column S) (without the header row), add a data validation to allow only whole numbers greater than or equal to 0.

10. For columns I to S, while adding data validation, add an error message: **The entered value should be a whole number greater than 0**.

11. Try entering data for Serial numbers 12–15 to see if the data validation is working.

Self-evaluate your work with the demo sheet: [Advanced Excel Course — Demo Sheet](/Advanced%20Excel%20Course%20-%20Demo%20Sheet.xlsx)


## Exercise 2: Conditional Formatting

*To be completed after Video 2*

1. In the file **Advanced Excel Course — Work Sheet**, go to the sheet called **Conditional Formatting**.

2. Using conditional formatting, identify the schools which have more girls than boys and highlight the corresponding cells in the **Girls Enrolment** column in green.

3. Filter the **Girls Enrolment** column by green colour to count how many schools in this list have more girls than boys.

4. Using conditional formatting, identify the schools which have more than 35 boys and highlight the corresponding cells in the **Boys Enrolment** column in yellow.

5. Filter the **Boys Enrolment** column by yellow colour to count how many schools in this list have more than 35 boys.

6. Using conditional formatting, identify the schools which have zero male or female teachers or zero total teachers and highlight the corresponding cells in the columns **Male Teachers**, **Female Teachers** and **Total Teachers** in red.

7. Identify how many schools had zero female teachers and how many schools had no teachers at all.

8. Using conditional formatting, identify the schools which have more than 15 SC **or** ST students and highlight the corresponding cells in the **SC** and **ST** columns in green.

Self-evaluate your work with the demo sheet: [Advanced Excel Course — Demo Sheet](/Advanced%20Excel%20Course%20-%20Demo%20Sheet.xlsx)


## Exercise 3: IF Statements

*To be completed after Video 3*

1. In the file **Advanced Excel Course — Work Sheet**, go to the sheet called **IF Statements**.

2. Go to the **Girls>Boys** column. Using an IF statement, identify schools which have more girls than boys and mark them as **1**. Mark all other schools as **0**.

3. Go to the **SC/ST/OBC>General** column. Using an **IF** and an **AND** function together, identify schools which have SC/ST/OBC students and their number is more than the General category students. Mark all such schools as **Yes**, and others as **No**.

4. Count how many schools have more SC/ST/OBC students than General category students.

5. In cell **J34**, using a **COUNTIF** function, count the number of schools where the Girls Enrolment is more than 35.

6. In cell **J35**, using a **SUMIF** function, find the total enrolment in schools where Girls Enrolment is more than 35.

7. Using **COUNTIFS** function, count the number of schools where girls enrolment is more than 35 **and** total enrolment is more than 65 in cell **J36**.

8. Using **SUMIFS** function, sum the total enrolment of schools where girls enrolment is more than 35 **and** total enrolment is more than 65 in cell **J37**.

Self-evaluate your work with the demo sheet: [Advanced Excel Course — Demo Sheet](/Advanced%20Excel%20Course%20-%20Demo%20Sheet.xlsx)


## Exercise 4: Pivot Tables

*To be completed after Video 4*

1. In the file **Advanced Excel Course — Work Sheet**, go to the sheet called **Enrolment Data**.

2. Create a pivot table on a new sheet, with the type of management on the left side and the count on the right side.

3. Now go to the sheet **Enrolment Data** and using filters in the **Management** column, see if the values shown in the Pivot Tables sheet are matching with the total number of schools under the categories: Department of Education, Government Sponsored, Madrasa Recognized.

4. Create another pivot table on the same sheet with **Block** as rows, **Type of Management** as columns and their count as values.

5. Go to the sheet **Enrolment Data** and filter the Block as **Gorubathan** and Management as **Government Sponsored**. See if the count comes to **38**.

6. Create another pivot table on the same sheet with **Block** and **Type of Management** as rows, **Highest Class** as columns and their count as values.

7. Match your values with the demo sheet: [Advanced Excel Course — Demo Sheet](/Advanced%20Excel%20Course%20-%20Demo%20Sheet.xlsx)


## Exercise 5: VLOOKUP

*To be completed after Video 5*

1. In the file **Advanced Excel Course — Work Sheet**, go to the sheet called **VLOOKUP**.

2. Using **VLOOKUP** and school name as the reference column, find the total enrolment for schools listed in the VLOOKUP sheet (column G).

3. Using **VLOOKUP** and UDISE code as the reference column, find the electricity availability in schools that are listed in the VLOOKUP sheet (column G).

Self-evaluate your work with the demo sheet: [Advanced Excel Course — Demo Sheet](/Advanced%20Excel%20Course%20-%20Demo%20Sheet.xlsx)


## Exercise 6: INDEX & MATCH

*To be completed after Video 6*

1. In the file **Advanced Excel Course — Work Sheet**, go to the sheet called **INDEX MATCH**.

2. In the **MATCH** column, using a **MATCH** function, find the row numbers for the schools listed from serial numbers 1–10. Use the given UDISE Codes for this.

3. In the **INDEX** column, using an **INDEX** function and the values found in the MATCH column, find the school names for the schools listed from serial numbers 1–10.

4. In the **MATCH** column, using a **MATCH** function, find the row numbers for the schools listed from serial numbers 11–20. Use the School Names for this.

5. In the **INDEX** column, using an **INDEX** function and the values found in the MATCH column, find the UDISE codes for the schools listed from serial numbers 11–20.

6. Using **INDEX** and **MATCH** together, fill in the empty School Names (for serial no. 1–10) and the empty UDISE Codes (for serial no. 11–20).

7. Using **INDEX** and **MATCH** together, fill in the **Block** column (col. C).

Self-evaluate your work with the demo sheet: [Advanced Excel Course — Demo Sheet](/Advanced%20Excel%20Course%20-%20Demo%20Sheet.xlsx)

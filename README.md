# Excel-Data-Analysis-Practice
This repository contains problems completed from multiple tests using different Excel functions, tables, and graphs.
In this Repo, I plan to practice different skills and refine my story-telling and documentation through data exploration in Excel and by using markdown files like this on GitHub.
The three different test I will be using will cover these topics:

* Spreadsheet Functions to Organize Data
* Introduction to Filtering, Pivot Tables, and Charts
* Advanced Graphing and Charting

I will not be upload the entire test to this Repo. I will simply be using questions that required me to manipulate data on the spreadsheets and create formulas. Also, for naming convention any file with *Raw* in the file name is the **Original** file. All other files are the completed files. <br />
You can check out my LinkedIn profile by [clicking here](https://www.linkedin.com/in/william-hill-3ab051135/). <br />
I hope you enjoy!

## Spreadsheet Functions to Organize Data
This is the first of three quizes that required me to complete a spreadsheet <br />

### 1. Assuming you want the formula in H2 to always reference the cell directly to its left, correct the formula.  Once the formula is fixed, copy the formula down the column <br />

This question was requiring me to fix the formula located on cell `H2` from the `Store-Sales-2012 Raw` spreadsheet. This is a simple fix! As you can see in the formula <br />
`=IF($G$2>=30,"Large",IF(G2<=15,"Small","Medium"))` all you would need to do is remove the `$` from the `$G$2>30` portion of the of the `IF` function since you do not need an absolute cell reference there becuase you want to use the next cell down when using the fill handle to copy the formula to the bottom of the table and the absolute cell reference prevents this from happening. <br />
The final formula should look like this:
```
=IF(G2>=30,"Large",IF(G2<=15,"Small","Medium"))
```
After this, you can simply double-click the fill handle to copy the formula to the bottom of your table.

### 2. Using your newly created “Expanded Order Type” column, calculate the total “Sales” for all orders of “Medium” type (rounded to 2 decimal places).
In a previous question, the test asked me what type of function would return a "type" value from **Lookup Table 1** in cells `A2:B12`. The test wanted me to create this formula in cell `I2` under the column name **Expanded Order Type**. I decided that the`VLOOKUP` formula would be the best formula for this problem. With this information, I created this formula:
```
=VLOOKUP(G2,$A$3:$B$12,2,TRUE)
```
The next step wanted me to calculate the total sales for just the "Medium" order type. This type of question requires a conditional response and using a simple `SUM` function would require even more work to get the answer you are looking for. For this, a `SUMIF` function would be the perfect formula for this question. In cell `I2104` of the `Store-Sales-2012` file, you can see that I created that `SUMIF` function to answer this question. The formula used was:
```
=SUMIF(I2:I2103,"Medium", J2:J2103)
```
As you can see it will add together all the `Sales` from column `J` **_IF_** the condition of `"Medium"` was met in the `Expanded Order Type` column. This brings back a value of `$275,880.24` to answer the question being asked.

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
As you can see it will add together all the `Sales (J2:J2103)` **_IF_** the condition of `"Medium"` from cell range `I2:I2103` was met in the `Expanded Order Type` column. This brings back a value of `$275,880.24` to answer the question being asked.

### 3. The company gives a 1% discount on any Extra Large or larger orders.  In the “Discount” column, create a formula that returns 0.01 if the “Expanded Order Type” is Extra Large, XX Large, or XXX Large, and returns No Discount otherwise.
For this question, I decided to go with an `IFERROR` formula with a nested `VLOOKUP` formula. I came up with the following formula:
```
=IFERROR(VLOOKUP(I2,$A$14:$B$16,2,FALSE), "No Disc")
```
Now lets break it down! <br />
First I created a small table in cells `A14:B16` to reference in the `VLOOKUP` portion of the formula. After doing this, we can get started on creating the formula. I decided to approach the building of the formula in a similar fashion when creating a SQL query with a subquery. I built the `VLOOKUP` formula first. Starting off, I used `I2` to for my lookup value. Next, I used my table I created on cells `A14:B16` and made them absolute so that would not change when filling down the formula. the `2` represents the value I want returned if the condition was met. If it was not met, this is where the `IFERROR` formula comes in to play, but I will talk about that in a second. `FALSE` represents an exact match from the `Expanded Order Type` column. <br /> <br />
Now, here is when `IFERROR` formula steps in. If I were to not have the `IFERROR` formula there and an exact match was not met, the formula would pop out an error that looked like this: `#N/A`. We do not want this becuase the question wants me to return "No Discount" if it was not met. With the `IFERROR`, Excel will run `VLOOKUP` first. Instead of spitting out the `#N/A` error, Excel will then run the `IFERROR` function, and return the second part of the `IFERROR` function `"No Disc"` in my formula, hence the name `IFERROR`.

### 4. Create the formula from the previous question and copy the formula down to all the rows.  How many orders will have a discount applied?
This question is a pretty straightforward question. This is very similar to question number 2, but instead of using the `SUMIF` formula, we are going to use the `COUNTIF` formula. Quick side note, a basic `COUNT` will simply count the number of values in the given range. <br /> <br />
Here is my formula used in cell `K2104`:
```
=COUNTIF(K2:K2103,0.01)
```
I used this formula because it would count all the cells in the given range with `0.01` entered. I used the range from the `Discount` coulumn because that is what the question was asking. So, what I am trying to tell Excel to do is to **_COUNT_** the cell in range `K2:K2103` **_IF_** `0.01` was the value entered in the cell.

### 5. Create a formula in the “Sales with Discount” column and copy it down to all the rows.  What is the impact of the discount on total sales in 2012?  In other words, what is the difference between the sum of the “Sales” and the sum of the “Sales with Discount” (rounded to 2 decimal places)?
The first part of this question was pretty easy to complete. I came up with a basic formula to answer this:
```
Sales with Discount = Sales – (Sales * Discount)
```
If you notice with is calculation, I will have string values in the  `Discount` portion of the formula, which you cannot not quantify in Excel. To prevent the formula returning an error message, we will need to create another `IFERROR` formula to prevent an error being returned. So for this, I created this:
```
=IFERROR(J2-(J2*K2),J2)
```
Lets break it down! <br />
This will first calculate `Sales * Discount`, then it will subtract that from the corresponding cell in the `Sales` column. Now, since some cells in the `Discount` column have `"No Disc."` in the cell, it would produce an error becuase you cannot quantify a string function. This is where the `IFERROR` saves the day again! If my formula was not met, it will simple return the value from `Sales` column. <br /> <br />
The second portion of this question wants us to discover how much the discount cut into total sales. You can do this by simply using a `SUM` function for both the `Sales` and `Sales with Discount` columns. <br />
For the `Sales` column, I created this in Cell `J2104`:
```
=SUM(J2:J2103)
```
<br />

And for the `Sales with Discount` column, I created this in Cell `L2104`
```
=SUM(L2:L2103)
```
<br />

then we can create a basic subtraction formula by using this in Cell `J2108`
```
=J2104-L2104
```
Now We have the Answer to the question!


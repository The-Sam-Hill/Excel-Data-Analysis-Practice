# Excel-Data-Analysis-Practice
This repository contains problems completed from multiple tests using different Excel functions, tables, and graphs.
In this Repo, I plan to practice different skills and refine my story-telling and documentation through data exploration in Excel and by using markdown files like this on GitHub.
The three different tests I will be using will cover these topics:

* Spreadsheet Functions to Organize Data
* Introduction to Filtering, Pivot Tables, and Charts
* Advanced Graphing and Charting

I will not be uploading the entire test to this Repo. I will simply be using questions that required me to manipulate data on the spreadsheets and create formulas. Also, for naming convention any file with *Raw* in the file name is the **Original** file. All other files are the completed files. <br />
You can check out my LinkedIn profile by [clicking here](https://www.linkedin.com/in/william-hill-3ab051135/). <br />
I hope you enjoy!

## Spreadsheet Functions to Organize Data

### 1. Assuming you want the formula in H2 to always reference the cell directly to its left, correct the formula.  Once the formula is fixed, copy the formula down the column <br />

This question was requiring me to fix the formula located on cell `H2` from the `Store-Sales-2012 Raw` spreadsheet. This is a simple fix! As you can see in the formula `=IF($G$2>=30,"Large",IF(G2<=15,"Small","Medium"))` all you would need to do is remove the `$` from the `$G$2>30` portion of the `IF` function since you do not need an absolute cell reference there becuase you want to use the next cell down when using the fill handle to copy the formula to the bottom of the table and the absolute cell reference prevents this from happening. <br />
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
First I created a small table in cells `A14:B16` to reference in the `VLOOKUP` portion of the formula. After doing this, we can get started on creating the formula. I decided to approach the building of the formula in a similar fashion when creating a SQL query with a subquery. I built the `VLOOKUP` formula first. Starting off, I used `I2` to for my lookup value. Next, I used my table I created on cells `A14:B16` and made them absolute so that would not change when filling down the formula. the `2` represents the column of the table for the value I want returned if the condition was met. If it was not met, this is where the `IFERROR` formula comes in to play, but I will talk about that in a second. `FALSE` represents an exact match from the `Expanded Order Type` column. <br /> <br />
Now, here is when `IFERROR` formula steps in. If I were to not have the `IFERROR` formula there and an exact match was not met, the formula would pop out an error that looked like this: `#N/A`. We do not want this becuase the question wants me to return "No Discount" if it was not met. With the `IFERROR`, Excel will run `VLOOKUP` first. Instead of spitting out the `#N/A` error, Excel will then run the `IFERROR` function, and return the second part of the `IFERROR` function `"No Disc"` in my formula, hence the name `IFERROR`.

### 4. Create the formula from the previous question and copy the formula down to all the rows.  How many orders will have a discount applied?
This question is a pretty straightforward question. This is very similar to question number 2, but instead of using the `SUMIF` formula, we are going to use the `COUNTIF` formula. Quick side note, a basic `COUNT` will simply count the number of values in the given range. <br /> <br />
Here is my formula used in cell `K2104`:
```
=COUNTIF(K2:K2103,0.01)
```
I used this formula because it would count all the cells in the given range with `0.01` entered. I used the range from the `Discount` coulumn because that is what the question was asking. So, what I am trying to tell Excel to do is to **_COUNT_** the cells in range `K2:K2103` **_IF_** `0.01` was the value entered in the cell.

### 5. Create a formula in the “Sales with Discount” column and copy it down to all the rows.  What is the impact of the discount on total sales in 2012?  In other words, what is the difference between the sum of the “Sales” and the sum of the “Sales with Discount” (rounded to 2 decimal places)?
The first part of this question was pretty easy to complete. I came up with a basic formula to answer this:
```
Sales with Discount = Sales – (Sales * Discount)
```
If you notice within this calculation, I will have string values in the  `Discount` portion of the formula, which you cannot not quantify with an integer in Excel. To prevent the formula returning an error message, we will need to create another `IFERROR` formula to prevent an error being returned. So, I created this:
```
=IFERROR(J2-(J2*K2),J2)
```
Lets break it down! <br />
This will first calculate `Sales * Discount` since it is the innermost part of the formula surrounded by parenthesis, then it will subtract that from the corresponding cell in the `Sales` column. Now, since some cells in the `Discount` column have `"No Disc."` in the cell, it would produce an error becuase you cannot quantify a string function with a integer. This is where the `IFERROR` saves the day again! If my formula was not met, it will simply return the value from `Sales` column. <br /> <br />
The second portion of this question wants us to discover how much the discount cut into total sales. You can do this by using a `SUM` function for both the `Sales` and `Sales with Discount` columns. <br />

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

### 6. Currently, customers are responsible for paying the shipping costs.  The sales team suggests that customers really dislike paying shipping costs, and that offering “free shipping” instead of the 1% discount would likely increase sales.  Create a formula for the “Sales with Free Shipping” column that subtracts the “Shipping Cost” from the “Sales” only if the “Expanded Order Type” is Extra Large, XX Large, or XXX Large.  Copy the formula down to all the rows.  What would total 2012 sales have been if the company had offered free shipping instead of the 1% discount (rounded to 2 decimal places)?
Ok, there is a lot of information to absorb with the question, but lets break it down. There are two objectives in this question we need to complete to come to a conclusion.
First we need to create a formula similar to the formula to the formula that decided if there were to be a 1% discount or if no discount would apply. This question will have a formula that is structured differently to achieve similar results.<br />

We want to create a formula that subtracts the Shipping costs found in `Shipping Cost` column from thes sales in the `Sales` Column. To do this, we will need to create an `IF` formula with multiple `IF` statements nested within it. I came up with this formula for the first part of this answer:
```
=IF(I2="Extra Large", J2-P2, IF(I2="XX Large", J2-P2, IF(I2="XXX Large", J2-P2,J2)))
```
Ok, now we lets break this down. At first glance, it looks like a lot, but in theory it is not so bad. The first part of the formula, `(I2="Extra Large", J2-P2`, is pretty easy to understand after looking at some of the previous formulas we created. So what it is saying is **_IF_** `I2`, or the cell value in the `Expanded order type`, *Equals* `"Extra Large"`, then we subtract `Sales`**-**`Shipping Costs` (`J2-P2` of the formula). If that is not satisifed then we move on to the next `IF` formula. We will rinse and repeat the steps again for `XX Large` and `XXX Large` parts of the of the formula. Now, if none of conditions are met then we want it to return just the number in the `Sales` column show that the end of the formula, which is represented by `J2` at the end of the formula. New we will close of the formula with `)))` since that is how many open parenthesis we have.
<br />

Now that we have the first part of the qeustion completed, we can now move on to the second part of the question to get our final answer! <br />
All we have do now is create a basic `SUM` function to finish the problem and it can created like so:
```
=SUM(M2:M2103)
```
This will give us a total of `$3,712,048.65` to answer the question.

### 7. How much money would the company have saved in 2012 if it had offered free shipping instead of the 1% discount on Extra Large, XX Large, or XXX Large orders (rounded to 2 decimal places)?
This might be the easiest formula we have created. It will be a simple subtraction formula like so:
```
Money Saved = Sales with Free Shipping - Sales with Discount
```
Since we already have sum functions for both of these columns (`L2104` for `Sales with discount` and `M2104` for `Sales with Free Shipping`) from previous questions we can create this formula for the answer:
```
=M2104-L2104
```
This will give us a savings amount of `$9,008.97`

### 8. What would 2012 total “Sales” have been if the company had offered free shipping on any order shipped by Delivery Truck, and no additional discounts (rounded to 2 decimal places)?
First, we will need to create a new column. I named mine `Shipping cost discount for delivery truck`, but it does not have to be named like that and could easily be shortened down, but I digress. <br />
Now back to the formula, we will need to create a formula that subtracts the Shipping Cost from Sales if the Shipping Mode is by delivery truck. To do this, the `IF` formula will be the easiest way to do this. Here is what I came up with:
```
=IF(N2="Delivery Truck",J2-P2,J2)
```

Here is the analysis. So **_IF_** `Ship Mode`**=** `"Delivery Truck"` then we subtract `Sales`**-**`Shipping Cost`. Then, **_IF_** `Ship Mode`**_<>_** `"Delivery Truck"` then return `Sales`! <br />

Now that we have done that we can create a `SUM` formula to give us the answer to our question. <br />
That answer looks like this:
```
=SUM(O2:O2103)
```
And this can be found in cell `O2105`! <br /> <br />

That is all for this test! I hope you like my breakdown of my thought process for this test. Next I will be breaking down the next test, Introduction to Filtering, Pivot Tables, and Charts, soon!

# VBA-challenge

# Mutliple Year Stock Data

![stockmarket](https://github.com/Pooja14n/VBA-challenge/assets/144713762/6521bfb0-f5e6-4e4f-8aba-5b4c8de2e5a4)



In this homework assignment, VBA scripting is used to analyze generated stock market data for 2018, 2019, and 2020.

# Requirements

A script that loops through all the stocks for one year and outputs the following information, is created:

1. The ticker symbol
2. Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
3. The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
4. The total stock volume of the stock.
5. To return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".


# Step 1: The ticker symbol

1. A variable `summary_table_row` iss created to keep track of the location for each stock symbol/ticker.
2. All the ticker symbols are identified and were listed in a column, by creating a variable named `ticker_name` that is used to hold the name of the ticker.


# Step 2: Yearly Change

1. Three variables are used to hold the `opening` price, `closing` price, and `yearly_change`.
2. Yearly change is calculated from the difference between the closing price at the end of a given year to the opening price at the beginning of that year (yearly_change = closing - opening).
3. `yearly_change` variable gets reset to 0 after the yearly change for a particular ticker gets calculated, to start new calculation for the next ticker.
4. This column is formatted as "currency".
5. Conditional Formatting is applied to change the cell color to green if yearly change is positive, else, color changes to red.


# Step 3: Percent Change

1. A variable to hold the `percent_change` is created.
2. `percent_change` is calculated using the formula percent_change = (yearly_change / opening) * 100.
3. This column is formatted as "%".
4. Conditional Formatting is applied to change the cell color to green if yearly change and hence, the percent change is positive, else, color changes to red.
   

# Step 4: Total Stock Volume

1. A variable to hold the `total_stock_volume` is created.
2. The `total_stock_volume` for each stock symbol/ticker is calculated using `if-statement` within `for-loop` that ranges from row 2 to lastrow.
3. `total_stock_volume` variable gets reset to 0 after the total stock volume for a particular ticker gets calculated, to start new calculation for the next ticker.


# Step 5: The following three values are calculated and displayed in the output:

1. Greatest % Increase - `max` function is used to derive at this value.
2. Greatest % Decrease - `min` function is used to derive at this value.
3. Greatest Total Volume - `max` function is used to derive at this value.
4. `match` function is used to make reference to each of the above three values calculated, to its respective ticker symbol.


# Note
 1. `For Each` ws In Worksheets is used to loop through all the worksheets.
 2. A variable `lastrow` is used to count the number of rows.


# References
Referred to various class activity exercises, Microsoft Support Documentation, Stack Overflow, got support from BCS Learning Assistant in identifying and solving the error in my code for calculating `yearly_change` and for `max` function. 


# Files submitted including this README File
1. Three individual screenshots of the results (for year 2018, 2019, and 2020).
2. VBA script file - `Multiple_year_stock_data_Final.vb`.
3. Excel File with VBA Macro code under the `Module` for it - Multiple_year_stock_data_Final.xlsm.

   

# VBA Stock Price Analysis
## Stock price analysis using Excel VBA (macros)

![Stock Chart Art](Images/stock_chart_art.png)

The beauty of automating repetitive tasks in Excel with VBA scripting is demonstrated in this analysis, which seeks to answer the following questions:

From 2012-2016...
* Which NYSE stock had the highest trading volume in each year? The lowest?
* Which NYSE stock performed best in each year? The worst?

## Data Structure
The data began as an Excel workbook consisting of 5 spreadsheets, one for each year between 2012 and 2016.

Each spreadsheet consists of over half a million rows of data, one row representing a single stock's trading stats for one day. Each of the ~2,800 stock tickers on the NYSE takes up ~250 rows in each spreadsheet, one for each trading day of the year. Columns consist of:
* `ticker` - the stock ticker symbol
* `date` - the date for which the stock statistics are recorded
* `open` - the opening price of the stock
* `high` - the highest daily price of the stock
* `low` - the lowest daily price of the stock
* `close` - the closing price of the stock
* `vol` - the daily trading volume of the stock at close

The data is sorted alphabetically by ticker, then by date for each ticker.

## Testing Data
Because the dataset is so large, it was necessary to create a smaller set for testing. In order to cut down on processing time during writing and debugging the VBA script, the larger dataset was truncated to ~10,000 rows per spreadsheet.

This testing file is included in the repo, as `Raw_Stock_Data_Testing.xlsx`. Due to GitHub's 100MB file size limit, the larger raw dataset, as well as the completed analyzed workbook, is stored locally. Try [Yahoo Finance](https://finance.yahoo.com/) to download a similar dataset.

## Scripting
### Step 1: Summarizing daily data
The first task in discovering which stocks performed best and worst in each year, was to reduce the daily trading statistics for each stock to three yearly aggregates:
* Yearly Change ($USD difference in price between last and first trading day of the year)
* Percent Change ([Yearly Change / price on first trading day of year * 100])
* Total Volume (Sum of all daily trading volumes)

Because this would be unfeasible to do by hand, VBA was used. The script used can be found in this repo as "Analysis Module 1". Its function is described below:

#### Outer loop
The outer loop of the script iterates through the 5 spreadsheets in the workbook (`For Each ws In Worksheets`), each time:
* Writing headers for the new yearly aggregate table (to the right of the raw data)
    * format is `worksheet_object.Cells(x, y).Value = "text_to_write"`
* Determining the last used row of the spreadsheet (on line 35) by:
    * `RowNum = ws.Cells(Rows.Count, 1).End(xlUp).Row`:
        * navigates to the last cell of column A using `ws.Cells(Rows.Count, 1)` 
            * (`Rows.Count` returns last possible cell in the spreadsheet)
        * jumps up to the last populated cell using `.End(xlUp)`,
        * uses `.Row` to access the row number of that cell, 
        * stores this row number in the `RowNum` variable
* Running the inner loop
* Resetting the `PasteOffset` variable, which has been incremented during the inner loop to write data into successive cells.
* Moving to the next spreadsheet in the workbook

#### Inner loop
The inner loop iterates through cells in column A (`ticker` column), first row to last (last row is determined in outer loop). It serves to:
1. Ensure that the row below is for the same stock ticker (not in the last row of current ticker), in which case it:
    * Reads the `volume` from the current row and adds it to the running total for the current ticker, `RunTot`, which starts as 0.
    * Increments the loop counter `FirstTime`.
    * Checks if this is the first row of a new ticker (`FirstTime = 1`), in which case it reads the `open` price of the current row and stores it as the current ticker's opening price for the year, `OpenPrice`. 
        * (`OpenPrice` is used to calculate yearly percent change later)
        * If it is not the first row of a new ticker (`FirstTime != 1`, this step is not executed.
2. If the row below the current row has a different ticker symbol, we have reached the last row of that ticker. In this condition:
    * `volume` is added to the `RunTot` as normal
    * the `ticker` symbol is read, and written to the new aggregate table
    * the running volume total `RunTot` is written to the new aggregate table
    * the closing price for that day, in column `close` is read and stored as yearly closing price, `ClosePrice`, for later percent change calculation
    * yearly $USD change and percent change are calculated using stored yearly `OpenPrice` and `ClosePrice` and written to new aggregate table
        * an `If` statement catches opening stock prices recorded as 0 from breaking the script, which would cause division by 0.
    * if the yearly $USD change is positive, the cell it is written to is colored green, if negative, colored red.
    * `RunTot` is reset to 0 to be ready to hold volume running total for the next ticker group
    * `PasteOffset` is incremented so the next ticker and values writing in the new aggregate table occurs one row below the last written row.
    * `FirstTime` loop counter is reset for the new ticker group

#### Output
This results in the following aggregate table:

![Conditional Formatting](Images/conditional_formatting.png)

### Step 2: Identifying the winners and losers
After generating the aggregate table for each spreadsheet in the workbook, a second script was written to find the best and worst performers for each year, and write them into a new table. 

The goal was to create a table showing the ticker and value of the stock with the:
* Greatest yearly % increase in price
* Greatest yearly % decrease in price
* Greatest total trading volume
* Least total trading volume

The script can be found in the repo as "Analysis Module 2" and is described below:

#### Outer loop
The outer loop of this script iterates over spreadsheets in the workbook, each time:
1. Writing the table row and column headers
2. Resetting variables that hold:
    * paste offset `PstOS`
    * highest value found so far during iteration `TopDog`
    * lowest value found so far during iteration `UnderDog`
3. Determining the number of rows in the aggregate table, to bound later iteration through rows in the innermost loop.
3. Running middle loop
4. Changing results of middle loop to the proper "0.00%" format
5. Proceeding to the next spreadsheet in the workbook

#### Middle loop
The middle loop of this script iterates through the two columns for which minimum and maximum values must be recorded: Percent Change and Total Stock Volume. It serves to:
1. Execute inner loop
2. Write results of the inner loop to the new summary table
3. Reset variables used by the inner loop to store maximum, minimum, and ticker values
4. Proceed to the next column

#### Inner loop
The inner loop of this script iterates through rows in the column designated by the middle loop. It starts by storing two values, `TopDog` and `UnderDog`, both initially 0.

It then iterates through each value in a column. If the read value is larger than `TopDog`, that value is stored as the new `TopDog`. The corresponding ticker symbol is then recorded in the `HiTicker` variable.

Similarly, if the read value is lower than `UnderDog`, that value is stored as the new `UnderDog`. The corresponding ticker symbol is then recorded in the `LwTicker` variable.

The loop then moves to the next row, repeating the above steps.

In this way, when the loop is terminated after the last row, `TopDog` will be the maximum value in the column, and `UnderDog` will be the minimum. These values are stored for writing by the middle loop.

#### Output
This results in the final summary table:

![Yearly Winners and Losers](Images/yearly_winners.png)


Now that the script is written, it can summarize any number of year's worth of stock data to find the best and worst performing stocks. The results for each year in the workbook can be seen below.

##### 2012:
![2012 Screenshot](Images/2012_analysis_screenshot.png)

##### 2013:
![2013 Screenshot](Images/2013_analysis_screenshot.png)

##### 2014:
![2014 Screenshot](Images/2014_analysis_screenshot.png)

##### 2015:
![2015 Screenshot](Images/2015_analysis_screenshot.png)

##### 2016:
![2016 Screenshot](Images/2016_analysis_screenshot.png)



# Stock Analysis using VBA and EXCEL

## Overview of weekly challenge project

## Purpose

This project's scope was to analyze stock market data using VBA macros to sum total daily volumes and annual returns for each stock ticker. The challenge was to refactored the macro in order to make it run faster. 


## Original Project Details:

Steve wants to help his parents analyze stock to find investment opportunities. To help him we need to find the total daily volume and yearly return for each stock. Daily volume is the total number of shares traded throughout the day; it measures how actively a stock is traded. The yearly return is the percentage difference in price from the beginning of the year to the end of the year. 



## VBA Macro Workflow:
### Subroutine: All Stocks Analysis 
#### 1). Format timer
#### 3). Create header on worksheet for analysis table
#### 4). Initialize array of all tickers


![Original VBA Macro steps 1-4.](https://github.com/ClayMack/Stock-Analysis/blob/master/Resources/Original%20Code%201.png "VBA code.")

__(Original VBA Macro steps 1-4.)__

#### 5). Initialize variables for starting price and ending price
#### 6). Get the number of rows to loop over

![Original VBA code #2.](https://github.com/ClayMack/Stock-Analysis/blob/master/Resources/Original%20Code%202.png "VBA code.")

__(Original VBA Macro steps 5-6.)__


#### 7). Loop through tickers
#### 8). loop through rows in the data
#### 9). Get total volume for current ticker
#### 10). Get starting price for current ticker
#### 11). Get ending price for current ticker
#### 12). Output data for current ticker


![Original VBA code #3.](https://github.com/ClayMack/Stock-Analysis/blob/master/Resources/Original%20Code%203.png "VBA code.")

__(Original VBA Macro steps 7-12.)__


#### 13). Call a Subroutine for formatting the output data table
#### 14). End the timer


![Original VBA code #4.](https://github.com/ClayMack/Stock-Analysis/blob/master/Resources/Original%20Code%204.png "VBA code.")

__(Original VBA Macro steps 13-14.)__





## Refactored Code:

#### 1). Create a tickerIndex variable and set equal to zero before iterating over all the rows
#### 2). Use tickerIndex to index the ticker array and the three output arrays (step #3)
#### 3). Create three output arrays: tickerVolumes, tickerStartingPrices. and tickerEndingPrices

![Refactored VBA code #1.](https://github.com/ClayMack/Stock-Analysis/blob/master/Resources/Refractored%20Code%20%231.png "VBA code.")

__(Refactored VBA Macro steps 1-3.)__


#### 4). Create for loops to intitialize the tickerVolumes to zero and to loop over all the rows of data in the spreadsheet
#### 5). Write an if-then statement to check if the current row is the first row with the selected tickerIndex. If it is, then assign the current starting price to the tickerStartingPrices variable
#### 6). Write a script that increases the tickerIndex if the next row’s ticker doesn’t match the previous row’s ticker

![Refactored VBA code #2.](https://github.com/ClayMack/Stock-Analysis/blob/master/Resources/Refractored%20Code%20%232.png "VBA code.")

__(Refactored VBA Macro steps 4-6.)__


#### 7). Use a for loop to loop through your arrays (tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices) to output the “Ticker,” “Total Daily Volume,” and “Return” columns in your spreadsheet


![Refactored VBA code #3.](https://github.com/ClayMack/Stock-Analysis/blob/master/Resources/Refractored%20Code%20%233.png "VBA code.")

__(Refactored VBA Macro step 7.)__




## Results: 

After refactoring the macro we were able to cut the run time to 0.125 seconds for the 2017 data and 0.129 seconds for the 2018 data. The attached spreadsheet contains the macro with both the original subroutine as well as the new refactored subroutine with buttons for to run the refractored subroutine and to clear the data.

![Refactored VBA code #3.](https://github.com/ClayMack/Stock-Analysis/blob/master/Resources/VBA_Challenge_2017.png "VBA code.")

__(Timing for 2017 refactored subroutine)__


![Refactored VBA code #3.](https://github.com/ClayMack/Stock-Analysis/blob/master/Resources/VBA_Challenge_2018.png "VBA code.")

__(Timing for 2018 refactored subroutine)__


## Summary: 

As we see in this project, there are advantages to refactoring the code we originally wrote. The obvious one is speed. By simplifing complext code, you can dramatically shorten the run time of a macro. Secondly, refractoring can make the code easier to understand. This can make problematic bugs easier to spot, and it can lead to further performance enhancements. At the same time there could be disadvantages of refractoring code. If the code is very long, there could be bugs that develope from the refactoring if a subroutine contains a bug that later gets recalled into another subroutine. This could lead to calculation errors in macros.

While the original VBA script was logically sound and ran without error. It was slower, and if the dataset was much larger, it could have resulted in overall performance degradation. Given the project at hand, this performance increase may not have been worth the effort as there was less than a second speed increase.  





















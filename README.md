#Stock Analysis

## Overview of Project

### Project Goals
The goal of this project was to write a macro to facilitate the analysis of yearly stock data. The analysis included the yearly trade volume and the yearly return for each stock.

### Project Task
This project had two parts:

First we needed to create a Macro that would evaluate stock tables for a given year and then diplay the trade volume and return for each stock in a separate worksheet. We created a launch button for the analysis so anyone could perform the
calculations without haveing to install the developer tab. I did some basic formatting to make the results easier to read and highlighted the positive returns.

The second part of the project involved reviewing the already written code to make it more efficient. Using arrays to hold all of our variables makes is easier for the computer to compute and store the data all at once, instead of having
to switch between the sheets writing, and rewriting, multiple times.

## Results

The initial code got the job done, but it took more than 10 seconds to evaluate, calculate, and display the results for one year of a small sample of stocks for the year. The revized code that stores all the calculated values in arrays took less than one second to run.
Should this ever be used to evaluate a larger sample of stock data the calculations won't bog down the computer.

### Original Code
The original code stored values as singles variable that had to be zeroed out with each stock ticker change:

 '3.Create variables for start price, end price

    Dim startingPrice As Double
    Dim endingPrice As Double

'Find number of rows in sheet for loop
    Worksheets(yearValue).Activate
    'RowEnd code taken from https://stackoverflow.com/questions/18088729/row-count-where-data-exists
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row!
'4.Loop through the tickers
    For i = 0 To 11
    
        ticker = tickers(i)
        totalVolume = 0
This resulted in a very long run time:
![2017_StockAnalysis_Homework](https://user-images.githubusercontent.com/86027932/124340537-da06a780-db83-11eb-8756-60c496cfd8c5.png)


### New Code
The new code store each calculation in an array so that which makes it easier for the computer to work with:
    '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Double
    Dim tickerEndingPrices(12) As Double

    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
     For tickerIndex = 0 To 11
        
        tickerVolumes(tickerIndex) = 0
This really sped up the macro run time:
[VBA_Challenge_2017](https://user-images.githubusercontent.com/86027932/124340491-9d3ab080-db83-11eb-9870-9987182ff7af.png)

## Summary

### Advantages:
Original Code: Easier for beginners to work with and can be easier to read since you only have to look at the variable name when reading your code.
Refactored Code: Much faster to run. Able to store multiple iterations of the same variable at the same time. 

### Disadvantages:
Original Code: Super slow on my computer. A large data set would take several minutes to run. 
Refactored Code: Can get confusing for an inexperiened coder when trying to find error. Looking for variable name & array index, using the correct one in the correct place.

Whole project: This code is very specific in that the return calculations will only work if the sheet is sorted in date order. Since we using an array to hold the stock tickers, the start & end prices
would still work if it wasn't sorted alphabetically, but if the dates get out of order the code doesn't look for to determine the start & end price.

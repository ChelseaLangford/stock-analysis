# VBA of Wall Street

## Overview of Project
The purpose of this project was to creata a VBA script which analyzed the stock return percentage and total daily volume traded for a specific subset of stocks for a given year. My (hypothetical) client had a particular interest in one green energy company, in which his parents were considering investing. However, he was concerned that investing everything in a single company might not be a wise decision. To confirm his suspicions, my client asked for an easy way of comparing the stock return percentage for the years 2017 and 2018 for the company his parents were interested in, as well as a group of comparable green energy companies.

Once this script was created, it was determined that while the script ran efficiently when comparing the volume and return percentages for a set of 12 stocks, it might not be as quick when asked to run the analysis on a much larger data set, say, all publicly traded stocks in a given year. The second phase of this project involved refactoring the code to find a more efficient way of scanning through the daily volume and stock prices for each stock and generating the percent return from the ending price compared to the stock's starting price.

## Results

### Initial Code 
In order to generate the the return percentage comparison for year's worth of stock prices, I first needed to create code that would recognize the individual stock tickers. To do this, I created an array which defined an index for each of the 12 tickers:

```
   ' Initialize array of all tickers
   Dim tickers(12) As String
   tickers(0) = "AY"
   tickers(1) = "CSIQ"
   tickers(2) = "DQ"
   tickers(3) = "ENPH"
   tickers(4) = "FSLR"
   tickers(5) = "HASI"
   tickers(6) = "JKS"
   tickers(7) = "RUN"
   tickers(8) = "SEDG"
   tickers(9) = "SPWR"
   tickers(10) = "TERP"
   tickers(11) = "VSLR"
   ```
After setting variables for the starting price and ending price of the stocks, as well as defining a variable for the end row count, I then created a loop for the code to loop through all the tickers within the array, through all the rows within the dataset: 
   
   ```  
   ' Loop through tickers
   For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       
       ' Loop through rows in the data
   Sheets(yearValue).Activate
       For j = 2 To RowCount
   ```
   Once the loop was set, I then created a series of If statements to get the starting price and ending price for each ticker:
   
   ```
   ' Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If
           ' Get starting price for current ticker
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value

           End If

           'Get ending price for current ticker
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               endingPrice = Cells(j, 6).Value
               ```
The output of these statments was the total daily volume as well as the return percentage (endingPrice / startingPrice - 1) for each ticker. 

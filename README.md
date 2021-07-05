# Multi-Year Stock Analysis
## Overview
After reviewing the workbook provided that allowed him to analyze the stock information from 2017 
and 2018, Steve is interested in having the workbook run analysis of the entire stock market over 
the last few years. In order to make sure the workbook will be able to handle the increase in input 
data, the original code has been refactored to make it more efficient.

## Results
In looking at the performance of the stocks between 2017 and 2018, there is a pretty significant 
difference between the two. The 2017 stocks overall performed much better than they did in 2018,even 
though the overall total volumes were very close to the same. To determie the total volume the code 
compares the ticker to the current stock being analyzed, and if they match it adds the given volume 
to the total. Here is the original code that was used to get this result:
      
      For j = 2 To RowCount
            'Activate data worksheet
            Worksheets(yearValue).Activate
            'get total volume for current ticker
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value

In order to get the return, a little more analysis was required. Return isn't a value on the original data, so code was used to find the starting and ending prices, and the return was calculated using those two values:
      
      'get starting price for current ticker
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                startingPrice = Cells(j, 6).Value
            End If
            
            'get ending price for current ticker
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                endingPrice = Cells(j, 6).Value
            End If

Because there was such a significant difference in return for each year, performing this analysis 
on further years is definitely going to give a better picture of how each of the stocks performs 
over time.

Prior to refactoring the code, each of the 2017 and 2018 analyses ran for approximately 60 
seconds. They really seemed to lag the system. After refactoring the code the run-times were 
significantly improved for both years. The following screenshots show the refactored run times for 
each year:

![VBA_Challenge_2017](https://user-images.githubusercontent.com/85318060/124412698-efa2db00-dd03-11eb-8087-61878d2ba0c9.PNG)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/85318060/124412708-f4678f00-dd03-11eb-8ae0-1ca3a269afd1.PNG)

In order to refactor the code to run more efficiently, the first thing that was done was setting up each of the three output values as their own arrays, as follows:

     '1b) Create three output arrays
          Dim tickerVolumes(12) As Long
          Dim tickerStartingPrices(12) As Double
          Dim tickerEndingPrices(12) As Double
          
By setting up these arrays, and then creating a ticker Index, we allow the code to run through as 
many iterations as needed to get all of the stock value, which will help when looking at further 
stock data for other years.

The arrays changed the code slightly for how we captured each of the three output values noted above. Instead of going through each ticker and output separately, the code is able to go through each of the tickers using the ticker Index value:
      
      '2b) Loop over all the rows in the spreadsheet.
      For i = 2 To RowCount
          Worksheets(yearValue).Activate
          '3a) Increase volume for current ticker
          If Cells(i, 1).Value = tickers(tickerIndex) Then
              tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
          End If
        
          '3b) Check if the current row is the first row with the selected tickerIndex.
          'If it is Then set the starting price
              If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = 
              tickers(tickerIndex) Then
                  tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            End If
        
        '3c) check if the current row is the last row with the selected ticker
                'If it is Then set the ending price
            
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            End If

             'If the next rows ticker does not match, increase the tickerIndex.
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i - 1, 1).Value = tickers(tickerIndex) Then
                'Then increase the ticker index
                tickerIndex = tickerIndex + 1

Then at the end all of the ticker outputs are gone through one after another, utilizing one final for loop:
 
     '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
        For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1
        
      Next i

After all of these changes, the code runs much more smoothly, and it will be much more efficient
when adding additional data to analyze.

## Summary
In summary, refactoring code has several advantages, but there is also a big potential 
disadvantage.Some of the advantages are that it can make the code easier to read, and it can make 
it run more efficiently, depending on what kind of refactoring you are doing. The one big 
disadvantage is that any time code is rewritten there is a chance to introduce bugs that may not be 
caught by initial testing.

For this code the refactoring has definitely made it run more efficiently, which should help with 
any further analysis that Steve would like to complete. The code also contains more notes on what 
has been completed, which will help anyone who is trying to look at the code in the future. The 
disadvantage on this code is that the refactored code is slightly more complex than the original 
code, which may make it more difficult to follow, and there is the potential that there could be 
bugs that may not show themselves until the code is being used for more data. Overall refactoring 
is a good way to reorganize code that has been completed to put it in the most logical order and 
make sure that everything is running as efficiently as possible.

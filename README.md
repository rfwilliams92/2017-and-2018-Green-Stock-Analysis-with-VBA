# Green Stock Analysis with VBA in Excel

## Overview of Project

I created a VBA module to loop through a large set of green stock data to create a well defined table representing stock trends. Then refactored the module to be more flexible and run faster with much larger data sets. 

### Purpose
The objective was to create a VBA module that could move through a large set of green stock data to provide analysis on which stocks were more profitable in 2017 and 2018. Then refactor data to run faster.

## Results 
The biggest difference in the refactored code was the change from variables to variable arrays.
```
Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single
```
This change allows me to get rid of one `For` loop. These arrays allows you to group similar elements and easier to sort the data. But before that I had to set all the 
tickerVolumes(tickerName) = 0
```
For tickerName = 0 To 11
        
        tickerVolumes(tickerName) = 0
    
    Next tickerName
```
And lastly the variables needed to be modified to include the increasing ticker index count.

```
'2b) Loop over all the rows in the spreadsheet.
    For i = firstRow To RowCount

'3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    
'3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
    
'3c) check if the current row is the last row with the selected ticker
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
    

'3d) Increase the tickerIndex.
        tickerIndex = tickerIndex + 1
        
        End If

```

## Summary

###Advantages of Refactoring Code
A faster runtime to accommodate larger and larger data sets
Creates flexibility in the code to be reused elsewhere.

###Disadvantages of Refactoring Code
With more variables and arrays created to provide more flexibility can make the code harder to follow due to the increased abstraction.

###Advantages and Disadvantages of Refactored VBA script
Specifically the refactored code seems to speed up run time by about 5x.
Initial Analysis for 2017
![This is an image](/Resources/2017_Initial_Analysis.png)
Refactored
![This is an image](/Resources/2017_Refactored.png)
The disadvantage if this VBA code is that although being more much more flexible (easily applied to different stocks or a much larger data set), is that it is delicate in construction. A misspelled variable or a mistake with the syntax in VBA make it very easy to get a compile error.

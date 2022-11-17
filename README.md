## Overview of Project

### Purpose
The purpose of this project was to allow Steve to quickly perform research on certain stocks and to streamline the time it takes to find that information. In order to get this completed the supplied file with the Microsoft Excel VBA code had to be refactored for quicker processing time.


## Results
I began by copying the code and created a macro for that code then refactored the code into another macro. I then ran both codes on the 2018 data and verified the results. The refactored code ran quicker than the original code that was supplied. The below code was added to the original script. 

    '1a) Create a ticker Index
    
     tickerIndex = 0

    
    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
        
    '2a) Create a for loop to initialize the tickerVolumes to zero.
    
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    
    Next i
        
    '2b) Loop over all the rows in the spreadsheet.
        
    For i = 2 To RowCount
        
        '3a) Increase volume for current ticker
       
       tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

            End If
            
            '3d Increase the tickerIndex.
            
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
            
            End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    
    Next i



## Summary
Refactoring the code allowed Excel to perform the function quicker. As you can see below the time it took the original code was about 332 milliseconds and the refactored code ran about 58 milliseconds. So, the difference between the two is about 274 milliseconds!

- The advantage of refactoring code is that you may not have to write the whole code from scratch. A refactored code also can perform functions quicker. 

### Original Code

![Original](https://github.com/princetonduarte/stock-analysis/blob/8b27541f0a1e371f6ea10bac993148ad76f35141/Resources/VBA_Challenge_2018.png)

### Refactored Code
![Refactored](https://github.com/princetonduarte/stock-analysis/blob/0329b6194ff44dad8b00aebe009b7b78cb60ef4b/Resources/VBA_Challenge_2018_refactored.png)

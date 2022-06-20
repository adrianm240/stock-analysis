# stock-analysis
# Analyzing Stock Returns with Refactored Excel VBA Macro

## Overview of Project
Used Excel VBA to refactor a macro's code to automate the analysis of stock information.

### Purpose
To refactor and improve the efficiency of an Excel VBA macro code that automates the analysis of stock information. The customer originally wanted a VBA macro that analyzed twelve stock tickers, but in the future he plans to analyze thousands of stocks simultaneously and in order to do requires a more efficient code that will run the script faster.

## Results
The refactored code was significantly faster than the original. The 2018 refactored script was 725% faster (0.078125 vs 0.6445313), while the 2017 refactored script was 735% faster(0.078125 vs 0.6523438).

![VBA_Challenge_2017](https://user-images.githubusercontent.com/106203262/174647248-bfb98e61-b902-497c-878b-c00d547f4748.png)   ![VBA_Challenge_2018](https://user-images.githubusercontent.com/106203262/174647261-24fcd1d4-8ee1-4b1a-822f-5b06e2f9cd9a.png)

### Refactored Code
The primary difference in the refactored code was the use of a ticker index and three array variables which allowed the script to loop through the data only 1 time to collect the same information as the original, which did it in 12 loops.
'1a) Create a ticker Index
    tickerIndex = 0
    
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

The code then initiated two for loops; one to set the variable tickerVolumes to 0 and the other to begin looping over all the rows in the data.
'2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        
        tickerVolumes(i) = 0
    
    Next i
    
    '2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount

The next steps were to increase the volume amounts for the current ticker, check if the current row was the first row with the selected tickerIndex and then with the last row and if it didn't match, increase the tickerIndex.
'3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                       
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then

            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
        'If the next row’s ticker doesn’t match, increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        End If

The final step in the refactored portion of the code was to loop through the arrays to output the analysis on a spreadsheet.
For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i

## Summary

### Refactoring Code in General
A major advantage of refactoring the Excel VBA code is that it significantly improved the efficiency and speed of the script. As a result, the script can be used on larger datasets. However, a disadvantage to refactoring the code was that it took time and effort away from producing something new to focusing on something that already worked. In certain circumstances you may not want to dedicate resources to improve already working code, such as a soon to end/expire program, but if the life shelf of a code is long, then refactoring it makes perfect sense to improve its longevity and reduce its technical debt.

### Refactoring The Stock Ticker Excel VBA Code
For the original stock ticker Excel VBA code in particular, a pro of refactoring it was that it both improved its speed and expanded the number of stocks it could analyze in the future. Aside from the time it took to refactor the code, there really was no apparent con in this case.

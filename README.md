# VBA Challenge - Refactor VBA Code

## Overview of Project

### Purpose 
The original All Stocks Analysis project was developed for the end-user to push a control button and be able to quickly analyze a single tab of stock data grouped by the stock year. The performance of the All Stocks Analysis VBA code was good but it needs to be taken into consideration that the analysis was for only 12 stocks and two years. There's concern that the code will not scale well with thousands of stocks over multiple years. Therefore, there's a need to refactor the All Stocks Analysis code so that it will loop through all the data one time in order to collect the same information we did in the original VBA code and improve the efficency and execution time. 

## Results

### Refactoring The Code
The original AllStocksAnalysis() uses a nested For loop to move through the ticker array one stock at a time where it captures the needed information of totalVolume, endingPrice and startingPrice and outputs the information onto the excel All Stocks Anlaysis sheet. It thens moves on to the next stock ticker in the array and repeats the process. 

Rather than loop through the data one stock and output to the sheet, the refactored code instead loops through capturing the needed information of totalVolume, endingPrice and startingPrice and once all the data is stored for all stocks it then uses a loop to output the information onto the spreadsheet.

Below is the entire code:

Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single
    yearValue = InputBox("What year would you like to run the analysis on?")
    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
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
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
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
        If Cells(i, 1).Value = tickers(tickerIndex) Then
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        End If
                
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
                
        '3c) check if the current row is the last row with the selected ticker
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        End If
                
        'If the next row’s ticker doesn’t match, increase the tickerIndex.
            
        '3d Increase the tickerIndex.
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
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
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            Cells(i, 3).Interior.Color = vbGreen            
        Else        
            Cells(i, 3).Interior.Color = vbRed            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub


### Comparing Execution Times of Original Script adn Refactored Script 
The execution of the original code for both years completed in about .75 seconds. 

After the refactoring of the code, there was a noticable difference in the execution time dropping down to .11 seconds.


## Summary Statement

What are the advantages or disadvantages of refactoring code?
How do these pros and cons apply to refactoring the original VBA script?

Refactoring is a key part of the coding process. The first version of a coding project might not always be the most efficient. Refactoring provides an opprotunity to clean up the code, reduce steps, improve the logic, use less memory and add documentation for future users to read. 

Refactoring can be difficult if your refactoring someone else's code and they didn't provide adequate documentation. It can be difficult understanding their logic if the code doesn't follow good programming logic and they provided enough commentary. 

For the purpose of this challenge, it was good to see that there was more than one way to work with loops and how the structure of loops can have an effect on overall performance and execution time. Also, writing out the steps as part of the documentation proved to be very helpful at refreshing my memory when I came back to the code several days later to write up my summary. 

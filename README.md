# VBA Stock Analysis

## Overview of Project
* Using VBA and Macros to provide Technical Analysis on Green Stocks 

### Purpose
* Taking our initial code for Technical Analysis on 12 similar stocks and taking yearly return and total volume for the past 2 years to identify and highlight any key metric and trends that provide insight on which stock is worth investing in

## Results
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
    Dim tickervolumes(12) As Long
    Dim tickerstartingPrices(12) As Single
    Dim tickerendingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickervolumes(i) = 0
        tickerstartingPrices(i) = 0
        tickerendingPrices(i) = 0
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
    tickervolumes(tickerIndex) = tickervolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
           If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerstartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
            
            
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerendingPrices(tickerIndex) = Cells(i, 6).Value
         End If

            '3d Increase the tickerIndex.
          If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            End If
         
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
       Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = tickerendingPrices(i) / tickerstartingPrices(i) - 1
        
        
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

### Insight
* When dealing with the original DQ Analysis, we're left with the agreggate of every single function and progress, the code is very rudimentary and overflowing with manual data and code. Processing time for all of this is good enough, but can be improved through refactoring. 

#### What are the advantages or disadvantages of refactoring code?
* The advantages of refactoring code is being able to take different frameworks and lenses when it comes to coding and find which style, syntax, and methodology are the most effecient, concise, and logical. Every coder will have very similar yet different layouts and structure. Through debugging and inquiry, one is able to see which method is better when it comes to processing and presneting the data highlighted within the code. At the same time, there's also disadvantages to this; coders will have more or less the same code. The matter lies in the difference, when working through it, what works for one code, wont necessarily work for the other and can make or break a VBA script depending on how the data is pulled and utilized. Refactoring is essentially a double edged sword, it begs the question of if it's not broken, dont fix it or reinventing the wheel. This varies and the volume of data and the scale of these analysis projects.
#### How do these pros and cons apply to refactoring the original VBA script?
*  We're able to shorten the length of the "chain" or "loop" by condensing the VBA script through restructuring and reformatting the layout of the original code. Overall this leads to a shorter processing time. By creating an index array and ticker index, fine tuning the code to provide more specific data such as specifying the year and row control, leads to more precise codes and scripts. This becomes more and more important as the amount of data increases and operational effieciency is sustainable and maximized. 


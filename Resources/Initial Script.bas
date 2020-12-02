Attribute VB_Name = "All_Stocks_Analysis"
'Format the output sheet on the "All Stocks Analysis" worksheet.
'Initialize an array of all tickers.
'Prepare for the analysis of tickers.
'Initialize variables for the starting price and ending price.
'Activate the data worksheet.
'Find the number of rows to loop over.
'Loop through the tickers.
'Loop through rows in the data.
'Find the total volume for the current ticker.
'Find the starting price for the current ticker.
'Find the ending price for the current ticker.
'Output the data for the current ticker.

Sub AllStocksAnalysis_2()
    'set timer to measure code performance
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer

    'Format the sheet that will show the results
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
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
    
    Dim startingPrice As Double
    Dim endingPrice As Double
    
    Worksheets(yearValue).Activate

    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    rowStart = 2
    
    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
        
        Worksheets(yearValue).Activate
        For j = rowStart To RowCount
        '5a) Get total volume for current ticker
            If Cells(j, 1).Value = ticker Then
            
                totalVolume = totalVolume + Cells(j, 8).Value
                
            End If
        
        '5b) get starting price for current ticker
        
            If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
                
                startingPrice = Cells(j, 6).Value
            
            End If
        
        '5c) get ending price for current ticker
            If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
                
                endingPrice = Cells(j, 6).Value
            
            End If
        
        Next j
        
        '6a) Output data for current ticker
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 1).HorizontalAlignment = xlLeft
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 2).HorizontalAlignment = xlLeft
        Cells(4 + i, 3).Value = (endingPrice / startingPrice) - 1
        Cells(4 + i, 3).HorizontalAlignment = xlLeft
        Cells(4 + i, 3).NumberFormat = "0.0%"
    
    Next i
    
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

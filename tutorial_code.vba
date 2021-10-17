Sub AllStocksAnalysis()

    Worksheets("All Stocks Analysis").Activate
    Dim startTime As Single
    Dim endTime As Single

    yearvalue = InputBox("What year do you want your analysis in?")
    startTime = Timer
    
    Cells(1, 1).Value = "All Stocks ( " + yearvalue + ")"
    
    'give headers to columns
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volumn"
    Cells(3, 3).Value = "Return"
    
    'array of all tickers
    Dim tickers() As Variant
    tickers() = Array("AY", "CSIQ", "DQ", "ENPH", "FSLR", "HASI", "JKS", "RUN", "SEDG", "SPWR", "TERP", "VSLR")
    
    'variables for starting price and ending price
    Dim startingprice As Single
    Dim endingprice As Single
    
    Worksheets(yearvalue).Activate
    RowCount = Cells(Rows.Count, "a").End(xlUp).row
    
    For i = 0 To 11
        ticker = tickers(i)
        'do stuff with ticker
        totalvolumn = 0
    
    Worksheets(yearvalue).Activate
        For j = 2 To RowCount
            If Cells(j, 1).Value = ticker Then
                totalvolumn = totalvolumn + Cells(j, 8).Value
            End If
        
            If Cells(j - 1, 1) <> ticker And Cells(j, 1) = ticker Then
                startingprice = Cells(j, 6).Value
            End If
            
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                endingprice = Cells(j, 6).Value
            End If
        Next j
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalvolumn
        Cells(4 + i, 3).Value = endingprice / startingprice - 1
    Next i
endTime = Timer
msgbox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearvalue)


End Sub

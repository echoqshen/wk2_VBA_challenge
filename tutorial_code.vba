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
    
    
' for loop that checks each ticker with every row of the selected sheet 
    For i = 0 To 11
        ticker = tickers(i)
        'do stuff with ticker
' totalvolumn is initialized at the start of each loop and used again to calculate vol for each ticker 
        totalvolumn = 0
    
    Worksheets(yearvalue).Activate
'for each ticker in the out loop ( next i ) we look at each row of the selected data set 
        For j = 2 To RowCount

' if we hit the ticker we are looking at in the outer loop increment the total for that ticker to eventually find total for that ticker
            If Cells(j, 1).Value = ticker Then
                totalvolumn = totalvolumn + Cells(j, 8).Value
            End If
' the dataset is sorted by ticker. so if we detect a change in ticker then we have hit the starting price 
            If Cells(j - 1, 1) <> ticker And Cells(j, 1) = ticker Then
                startingprice = Cells(j, 6).Value
            End If
' the dataset is sorted by ticker. so if we detect a change in ticker then we have hit the ending price            
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                endingprice = Cells(j, 6).Value
            End If

' END OF LOOP FOR ROW 

        Next j
        
        Worksheets("All Stocks Analysis").Activate
' populate results 
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalvolumn
        Cells(4 + i, 3).Value = endingprice / startingprice - 1

' END OF LOOP FOR EACH TICKER 
    Next i
endTime = Timer
msgbox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearvalue)


End Sub

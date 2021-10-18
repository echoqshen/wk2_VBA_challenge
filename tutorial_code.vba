Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single
    

'    'use 4th sheet for debuging
    Set Db = Worksheets("debug")
'    'debug counter
    Dc = 1
'    ' debug step
    Db.Cells(Dc, 1) = "debug output"
   Dc = Dc + 1
'
   yearValue = InputBox("What year would you like to run the analysis on?")
'
   startTime = Timer
'
'    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
'
   Range("A1").Value = "All Stocks (" + yearValue + ")"
'
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
    RowCount = Cells(Rows.Count, "A").End(xlUp).row

    '1a) Create a ticker Index
    Dim TickerIndex As Integer
    TickerIndex = 0


    '1b) Create three output arrays

    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

    ''2a) Create a for loop to initialize the tickerVolumes to zero.

    For j = 0 To 11
    tickerVolumes(j) = 0
    tickerStartingPrices(j) = 0
    tickerEndingPrices(j) = 0
    Next j
    
    
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount

        '3a) Increase volume for current ticker
        tickerVolumes(TickerIndex) = tickerVolumes(TickerIndex) + Cells(i, 8).Value


        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        'sean
        'If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
        'echo
        If Cells(i - 1, 1) <> tickers(TickerIndex) And Cells(i, 1) = tickers(TickerIndex) Then
        
        tickerStartingPrices(TickerIndex) = Cells(i, 6).Value
        End If


        'End If

'        '3c) check if the current row is the last row with the selected ticker
'        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
'        tickerEndingPrices(TickerIndex) = Cells(i, 8).Value
'        End If
'       3C and 3D should be one condition not TWO
'     'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
'        'If  Then

        'sean
        'If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        'echo
        If Cells(i + 1, 1).Value <> tickers(TickerIndex) And Cells(i, 1).Value = tickers(TickerIndex) Then

        tickerEndingPrices(TickerIndex) = Cells(i, 6).Value

            Db.Cells(Dc, 1) = "ticker index is now"
            Dc = Dc + 1
            Db.Cells(Dc, 1) = TickerIndex
            Dc = Dc + 1
            Db.Cells(Dc, 1) = "Row is now"
            Dc = Dc + 1
            Db.Cells(Dc, 1) = i
            Dc = Dc + 1
            Db.Cells(Dc, 1) = "Next Row is now"
            Dc = Dc + 1
            Db.Cells(Dc, 1) = Cells(i + 1, 1)
            Dc = Dc + 1
            Db.Cells(Dc, 1) = "This Row is now"
            Dc = Dc + 1
            Db.Cells(Dc, 1) = Cells(i, 1)
            Dc = Dc + 1
            Db.Cells(Dc, 1) = "Total Ticker volume for this ticker is:"
            Dc = Dc + 1
            Db.Cells(Dc, 1) = tickerVolumes(TickerIndex)
            Dc = Dc + 1
            Db.Cells(Dc, 1) = "Starting Price for this ticker is:"
            Dc = Dc + 1
            Db.Cells(Dc, 1) = tickerStartingPrices(TickerIndex)
            Dc = Dc + 1
            Db.Cells(Dc, 1) = "Ending Price for this ticker is:"
            Dc = Dc + 1
            Db.Cells(Dc, 1) = tickerEndingPrices(TickerIndex)
            Dc = Dc + 1

          '3d Increase the tickerIndex.
        TickerIndex = TickerIndex + 1
        
        If TickerIndex = 1 Then
            Db.Cells(Dc, 1) = "TICKER INDEX IS 1 :"
            Dc = Dc + 1
            
            Db.Cells(Dc, 1) = "Starting Price for ticker 0 is:"
            Dc = Dc + 1
            Db.Cells(Dc, 1) = tickerStartingPrices(0)
            Dc = Dc + 1
            Db.Cells(Dc, 1) = "Ending Price for  ticker  0 is:"
            Dc = Dc + 1
            Db.Cells(Dc, 1) = tickerEndingPrices(0)
        End If
        
        'End If
        End If
    Next i
    
    
    

    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For kk = 0 To 11
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + kk, 1).Value = tickers(kk)
        Cells(4 + kk, 2).Value = tickerVolumes(kk)
        Cells(4 + kk, 3).Value = tickerEndingPrices(kk) / tickerStartingPrices(kk) - 1
        
        Db.Cells(Dc, 1) = "Starting Price for this ticker is:"
        Dc = Dc + 1
        Db.Cells(Dc, 1) = tickerStartingPrices(kk)
        Dc = Dc + 1
        Db.Cells(Dc, 1) = "Ending Price for this ticker is:"
        Dc = Dc + 1
        Db.Cells(Dc, 1) = tickerEndingPrices(kk)
        Dc = Dc + 1
        
        
        

    Next kk

    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    datarowstart = 4
    datarowend = 15

    For i = datarowstart To datarowend

        If Cells(i, 3) > 0 Then

            Cells(i, 3).Interior.Color = vbGreen

        Else

            Cells(i, 3).Interior.Color = vbRed

        End If

    Next i

    endTime = Timer
    msgbox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)


    ActiveSheet.UsedRange.Delete


End Sub

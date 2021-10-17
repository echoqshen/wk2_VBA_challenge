Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single
    
    'use 4th sheet for debuging
    Set Db = Worksheets("debug")
    'debug counter
    Dc = 1
    ' debug step
    Db.Cells(Dc, 1) = "debug output"
    Dc = Dc + 1

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer

    'Format the output sheet on All Stocks Analysis worksheet
    ' "All Stocks Analysis" is the output sheet
    ' the sheets for 2017 and 2018 have input data which should be immutable
    
    'configure the program to make changes to the output sheet by activating it
    Worksheets("All Stocks Analysis").Activate

    ' add a heading to the output sheet
    Range("A1").Value = "All Stocks (" + yearValue + ")"

    ' add headers to the output sheet
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    ' there are 12 different stock indexes which have string shorthad
    'Initialize array of all tickers
    Dim tickers(12) As String
    
    ' store the names of stock tickers as ordered set in tickers array

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

    ' we inputed a year which is used to select the sheet for that year with data
    ' "activate" that sheet so it is in reference for the program at this point
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    ' figure out how many rows the activated years sheet has
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).row
    
    
    ' above this was code given . below this is assignment questions
    '1a) Create a ticker Index
    Dim tickerIndex As Integer
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For j = 0 To 11
        tickerVolumes(j) = 0
    Next j
'
'
'
   '2b) Loop over all the rows in the spreadsheet.
   For i = 2 To RowCount
   
    'Db.Cells(Dc, 1) = i
    'Dc = Dc + 1
   
   'START OF LOOP FOR EACH ROW
   
   ' 3a) Increase volume for current ticker
   tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
   
    Db.Cells(Dc, 1) = tickerIndex
    Dc = Dc + 1
   
'        If Cells(i, 1).Value = tickerIndex Then
'
'        '3b) Check if the current row is the first row with the selected tickerIndex.
'        'If  Then
        'If Cells(j - 1, 1) <> ticker And Cells(j, 1) = ticker Then
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
'        End If
'
'        '3c) check if the current row is the last row with the selected ticker
'         'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
'        'If  Then
'        If Cells(i + 1, 1).Value <> tickerIndex And Cells(i, 1).Value = tickerIndex Then
'
'            '3d Increase the tickerIndex.
'            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
'        'End If
'        End If
'
'
'
'    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
'    For i = 0 To 11
'
'        Worksheets("All Stocks Analysis").Activate
'
'   END OF LOOP FOR EACH ROW

   Next i
'
'    'Formatting
'    Worksheets("All Stocks Analysis").Activate
'    Range("A3:C3").Font.FontStyle = "Bold"
'    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
'    Range("B4:B15").NumberFormat = "$" + "#,##0"
'    Range("C4:C15").NumberFormat = "0.00%"
'    Columns("B").AutoFit
'
'    dataRowStart = 4
'    dataRowEnd = 15
'
'    For i = dataRowStart To dataRowEnd
'
'        If Cells(i, 3) > 0 Then
'
'            Cells(i, 3).Interior.Color = vbGreen
'
'        Else
'
'            Cells(i, 3).Interior.Color = vbRed
'
'        End If
'
'    Next i
'
'    endTime = Timer
'    msgbox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

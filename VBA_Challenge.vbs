Sub AllStocksAnalysisOptimized()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")
    
    startTime = Timer
    
   '1) Format the output sheet on All Stocks Analysis worksheet
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    Range("A1").Font.FontStyle = "Bold"
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    'Format Table
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A4:C15").Borders(xlEdgeBottom).LineStyle = xlDouble
    Range("A4:C15").Borders(xlEdgeTop).LineStyle = xlDouble
    Range("A4:C15").Borders(xlEdgeRight).LineStyle = xlDouble
    Range("A4:C15").Borders(xlEdgeLeft).LineStyle = xlDouble
    'Format Numbers
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.00%"
    Columns("B").AutoFit

   '2) Initialize array of all tickers
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

    '3a Ticker Index
    tickerIndex = 0
    
    '3b) Create Arrays
    Dim tickerVolumes(12) As Long
    Dim tickerEndingPrices(12) As Single
    Dim tickerStartingPrices(12) As Single


   '4) Loop to initialize tickers
   For i = 0 To 11
       tickerVolumes(i) = 0
   Next i
       
    '5) loop through rows in the worksheet
       
       For i = 2 To RowCount
       
           '5a) Increase volume for current ticker
           tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

           '5b) Check for first ocurrence of ticker
           If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
           
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            End If
           
            '5c) Check for Last ocurrence of ticker
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            '5d Increase the ticker by 1
            tickerIndex = tickerIndex + 1

           End If
           
        Next i
    '6 Loop the array for output and calculate percentaje
    
        For i = 0 To 11
        
            Sheets("All Stocks Analysis").Activate
            Cells(4 + i, 1).Value = tickers(i)
            Cells(4 + i, 2).Value = tickerVolumes(i)
            Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
        Next i
   
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

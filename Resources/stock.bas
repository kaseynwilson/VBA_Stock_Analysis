Sub StockTicker()

         Dim WS_Count As Integer
         Dim I As Integer
         Dim Numberoftherow As Integer
         
         ' Set WS_Count equal to the number of worksheets in the active
         ' workbook.
         WS_Count = ActiveWorkbook.Worksheets.Count

         ' Begin the loop.
         For I = 1 To WS_Count
            Sheets(I).Select
            ' Insert your code here.
            ' The following line shows how to reference a sheet within
            ' the loop by displaying the worksheet name in a dialog box.
            
            'Add Column Headers for summary table
            Cells(1, 9).Value = "Ticker"
            Range("J1").Value = "Yearly Change"
            Range("K1").Value = "Percent Change"
            Range("L1").Value = "Total Stock Volume"
        
            
            'Find last row of worksheet
            LastRow = Cells(Rows.Count, 1).End(xlUp).Row
           
            'set initial row number to 2 for summary table
            Numberoftherow = 2
            'Create counter for stock volume and set initial value to 0
            Dim TotalStockVolume As Double
            TotalStockVolume = 0
            'Create counter for ticker number
            Dim m As Integer
            m = 0
    
            
            'Create loop of worksheet to lastrow to input data in summary table
            For j = 2 To LastRow
                'Create variable to hold each Ticker name
                Dim Ticker As String
                Ticker = Cells(j, 1).Value
                'Create variable to hold the Ticker's opening price.
                Dim TickerOpen As Double
                TickerOpen = Cells(j - m, 3).Value
                
                'Create Variable for stock close price
                Dim TickerClose As Double
                TickerClose = Cells(j, 6).Value
        
                'Create variable for yearly change
                Dim TickerChange As Double
                TickerChange = TickerClose - TickerOpen
                
                'If ticker is different than previous row
                If Cells(j + 1, 1).Value <> Cells(j, 1).Value Then
                    'Then input value into Column I "Ticker"
                    Cells(Numberoftherow, 9).Value = Cells(j, 1).Value
                    
                    
                    'Also add stock volume of the row to the stock volume counter.
                    TotalStockVolume = TotalStockVolume + Cells(j, 7).Value
                    'input the end totalstockvolume into the summary table.
                    Cells(Numberoftherow, 12).Value = TotalStockVolume
                    
                    'If the ticker opening price does not equal zero
                    If TickerOpen <> 0 Then
                        'input into summary table the ticker change and percent change
                        Cells(Numberoftherow, 10).Value = TickerChange
                        Cells(Numberoftherow, 11).Value = (TickerChange / TickerOpen) '*100 Took out multiplying by 100 after adding number format %
                        Cells(Numberoftherow, 11).NumberFormat = "0.00%"
                    'If the ticker opening price DOES equal zero
                    Else
                        'input into summary table the ticker change and N/A for the percentage change (can't calculate a percent change w/ a zero starting point)
                        Cells(Numberoftherow, 10).Value = TickerChange
                        Cells(Numberoftherow, 11).Value = 0
                    End If
                    
                    'If yearly change/change in ticker is greater than 0 make cell green
                    If TickerChange > 0 Then
                        Cells(Numberoftherow, 10).Interior.ColorIndex = 4
                    'Else if yearly change in ticker is less than 0 make cell red.
                    ElseIf TickerChange < 0 Then
                        Cells(Numberoftherow, 10).Interior.ColorIndex = 3
                    End If
                
                    'Add 1 to the number of the row count so that next unique ticker, yearly change, etc. goes into the subsequent row.
                    Numberoftherow = Numberoftherow + 1
                
                    'Reset Stock Volume and counter for the ticker to zero
                    TotalStockVolume = 0
                    m = 0
                
                    'Redefine YearOpen with next row after we've gone through other calculations.
                    'YearOpen = WS.Cells(j + 1, 3).Value
                
                'if ticker is the same as previous ticker
                Else
                    
                    'Add stock volume to stock volume counter
                    TotalStockVolume = TotalStockVolume + Cells(j, 7).Value
                    
                    'Add quantity of 1 to the counter for the ticker
                    m = m + 1
                    
                End If
            
            'Go to next row in workbook
            Next j
            
            'Find Last Row of the Summary Table of Ticker
            Dim LastSummaryTableRow As Integer
            LastSummaryTableRow = Cells(Rows.Count, 9).End(xlUp).Row
            
    
            'BONUS SECTION
                
            'set headers/labels for table
            Range("O2").Value = "Greatest % Increase"
            Range("O3").Value = "Greatest % Decrease"
            Range("O4").Value = "Greatest Total Volume"
            Range("P1").Value = "Ticker"
            Range("Q1").Value = "Value"
            
            'ORIGINAL LOCATION OF SUMMARY TABLE
            'Find Last Row of the Summary Table of Tickers
            'Dim LastSummaryTableRow As Integer
            'LastSummaryTableRow = Cells(Rows.Count, 9).End(xlUp).Row
            'Loop through rows 2 to end of summary table
            For k = 2 To LastSummaryTableRow
                
                'Set Variables for the Greatest% increase, Greatest% decrease and Greatest total volume.
                'Set value to the corresponding cell in the table.
                Dim MaxIncrease As Double
                    MaxIncrease = Range("Q2").Value
                    'Range("Q2").NumberFormat = "0.00%"
                Dim MaxDecrease As Double
                    MaxDecrease = Range("Q3").Value
                    'Range("Q3").NumberFormat = "0.00%"
                Dim MaxVolume As LongLong
                    MaxVolume = Range("Q4").Value
    
                'If cells in row k column 10 ("yearly change") is greater than current Max Increase Then
                If Cells(k, 11).Value > MaxIncrease Then
                    'Set that value in the table
                    Range("Q2").Value = Cells(k, 11).Value
                    'Set that corresponding ticker value in the table
                    Range("P2").Value = Cells(k, 9).Value
                'Else if cells in row k column 10 "yearly change" is less than the current Max Decrease then
                ElseIf Cells(k, 11).Value < MaxDecrease Then
                    'Set that value in the table
                    Range("Q3").Value = Cells(k, 11).Value
                    'Set that corresponding ticker value in the table
                    Range("P3").Value = Cells(k, 9).Value
                
                End If
                
                'If cells in row k column 12 "Total Stock volume" is greater than current Max Volume then
                If Cells(k, 12).Value > MaxVolume Then
                    'Set that value in the table
                    Range("Q4").Value = Cells(k, 12).Value
                    'Set that corresponding ticker value in the table.
                    Range("P4").Value = Cells(k, 9).Value
                End If
                
            'Go to next row in ticker summary table
            Next k
            
            'Autofit columns
            Range("A:Q").EntireColumn.AutoFit
            
         'Go to Next workbook
         Next I

        
        
        
      End Sub


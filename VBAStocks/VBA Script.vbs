Sub VBAStocks():

'loop through all the worksheets
For Each ws In Worksheets

    ' add ticker, yearlychange, percentchange and total stock vol column headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    ' define key variables to iterate through the rows
    
    Dim individual_tickers As String
    Dim stockvol As Double
    Dim startrow As Double
    Dim endrow As Double
    Dim yearopen As Double
    Dim yearclose As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    
    ' tally up the stock volume as we iterate
    Dim totalstockvol As Double
    
    ' keep track of where to write out the summary
    Dim SummaryRow As Integer
    SummaryRow = 2
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
    ' loop through all the rows in ticker column
    For r = 2 To LastRow
        
        ' use conditionals to obtain yearopan and yearclose values for each ticker type
        If ws.Cells(r - 1, 1).Value <> ws.Cells(r, 1).Value Then
            startrow = r
            yearopen = ws.Cells(startrow, 3).Value
        End If
        
        endrow = 0
        
        If ws.Cells(r + 1, 1).Value <> ws.Cells(r, 1).Value Then
            endrow = r
            yearclose = ws.Cells(endrow, 6).Value
        End If
       
       
        ' assign values to column 1 tickers and column 7 stock volume
        individual_tickers = ws.Cells(r, 1).Value
        stockvol = ws.Cells(r, 7).Value

        
        ' when the ticker is the same, we add on the volume
        totalstockvol = totalstockvol + stockvol
        
        
        ' if ticker changes, print output in summary table
        If ws.Cells(r + 1, 1).Value <> individual_tickers Then
                
            yearly_change = (yearclose - yearopen)
            
            ws.Cells(SummaryRow, 10).Value = yearly_change
            ws.Cells(SummaryRow, 9).Value = individual_tickers
            ws.Cells(SummaryRow, 12).Value = totalstockvol
            ws.Cells(SummaryRow, 11).Value = percent_change
           
        
            If (yearopen = 0) Then
                percent_change = 0
                
            Else
                percent_change = (yearly_change / yearopen) / 100
            
            End If
            
            If ws.Cells(SummaryRow, 10).Value >= 0 Then
                ws.Cells(SummaryRow, 10).Interior.ColorIndex = 4
                
            Else
                ws.Cells(SummaryRow, 10).Interior.ColorIndex = 3
            
            End If
            
            
            ' move down each row of summary table
            SummaryRow = SummaryRow + 1
            
            ' once ticker changes, we reset
            totalstockvol = 0
            yearly_change = 0
            
    
        End If
  
    
    Next r
    
    
    ' fix number formatting
    ws.Range("J:J").NumberFormat = "0.00"
    ws.Range("K:K").NumberFormat = "0.00%"
    
    Next ws
    

End Sub

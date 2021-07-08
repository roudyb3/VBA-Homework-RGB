Sub AlphabeticalTesting()
    
  'Loop through Worksheets
  Dim w As Worksheet
  For Each w In ActiveWorkbook.Worksheets
  w.Activate
    
    'Set an initial variable for the ticker name
    Dim Ticker As String
    
    'Set an initial variable for the opening price
    Dim Open_Total As Double
    Open_Total = Cells(2, 3).Value
     
    'Set a variable for the closing price
    Dim Closing_Total As Double
    Closing_Total = 0
    
    'Set a variable for the volume
    Dim Volume_Total As Double
    Volume_Total = 0
    
    'Keep track of the location of all the tickers in the summary brand
    Dim Ticker_Table_Row As Integer
    Ticker_Table_Row = 2
    
    'Determine the last row
    NumRows = Range("A2", Range("A2").End(xlDown)).Rows.Count
    
    'Establish the for loop to go through it all now
    For i = 2 To NumRows
        
        'Move on if it's not the same value
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
        'Set the ticker name
        Ticker = Cells(i, 1).Value
        
        'Add to the closing total
        Closing_Total = Closing_Total + Cells(i, 6).Value
        
        'Add to the volume total
        Volume_Total = Volume_Total + Cells(i, 7).Value
        
        'Print the ticker in the summary
        Range("J1") = "Ticker"
        Range("J" & Ticker_Table_Row).Value = Ticker
        
        'Add in the difference and format it
        Range("K1") = "Yearly Change"
        If ((Closing_Total - Open_Total) > 0) Then
            Range("K" & Ticker_Table_Row).Value = Closing_Total - Open_Total
            Range("K" & Ticker_Table_Row).Interior.ColorIndex = 4
        Else
            Range("K" & Ticker_Table_Row).Value = Closing_Total - Open_Total
            Range("K" & Ticker_Table_Row).Interior.ColorIndex = 3
    
    End If
    
        'Add in the percent change and format
    If Closing_Total <> 0 And Open_Total <> 0 Then
        Range("L1") = "Percent Change"
        Range("L" & Ticker_Table_Row).NumberFormat = "#.##%"
        Range("L" & Ticker_Table_Row).Value = (Closing_Total / Open_Total) - 1
    Else
    End If
        'Add in total stock volume
        Range("M1") = "Total Stock Volume"
        Range("M" & Ticker_Table_Row).Value = Volume_Total
 
        'Add one to the Ticker Table
        Ticker_Table_Row = Ticker_Table_Row + 1
    
        'Reset the open and closing total
        Open_Total = Cells(i + 1, 3).Value
        Closing_Total = 0
        Volume_Total = 0
        
    'if the cell immediately following a ticker is the same then
    Else
    
        'get to the last number of the closing price
        Closing_Total = Closing_Total
        Volume_Total = Volume_Total + Cells(i, 7).Value
        
    End If
        
    Next i
    
    Next w
    
End Sub


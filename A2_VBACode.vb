Sub StocksTwo()

'Create variable for worksheet and start loop
Dim WS As Worksheet

For Each WS In ThisWorkbook.Worksheets

'Create variables needed
Dim Ticker As String
Dim Summary_Table As Integer
Summary_Table = 2

Start = 0
Total = 0

'Inputting headers to empty cells
WS.[J1] = "Ticker"
WS.[K1] = "Yearly Change"
WS.[L1] = "Percent Change"
WS.[M1] = "Total Stock Volume"
WS.[P2] = "Greatest % Increase"
WS.[P3] = "Greatest % Decrease"
WS.[P4] = "Greatest Total Volume"
WS.[Q1] = "Ticker"
WS.[R1] = "Value"

'Create variable for last row of data
Dim LRow As Long
LRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

'Loop thru Tickers and add new names to Summary Table
    For i = 2 To LRow
    
    If Start = 0 Then
        Start = WS.Cells(i, "C")
    
    End If
    
    Total = Total + WS.Cells(i, "G")
    
    'If Tickers are the same....
    If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
        Ticker = WS.Cells(i, 1).Value
        
'Print Ticker to table
    WS.Range("J" & Summary_Table).Value = Ticker
    Ticker = 0

 'Print Yearly Change
    Yearly_Change = WS.Cells(i, "F") - Start
    WS.Cells(Summary_Table, "K") = Yearly_Change
        
        'Add conditionals for highlighting Yearly Change cells
        If Yearly_Change < 0 Then
        WS.Cells(Summary_Table, "K").Interior.ColorIndex = 3
        Else
        WS.Cells(Summary_Table, "K").Interior.ColorIndex = 4
        End If
        
 'Print Percentage change and Total Stock Value
        WS.Cells(Summary_Table, "L") = FormatPercent(Yearly_Change / Start)
        WS.Cells(Summary_Table, "M") = Total
        Start = 0
        Total = 0
        
'Add one to Summary Row
Summary_Table = Summary_Table + 1


End If

Next i

Next WS

End Sub


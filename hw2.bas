Attribute VB_Name = "Module1"
Sub module2():

For Each ws In Worksheets

Dim i As Long
Dim TickerColumn As Long
' Dim NextRow As Integer
Dim NextRow As Long
' Dim NextTicker As Integer
Dim NextTicker As Long

' WorksheetName = ws.Name
TickerColumn = 1
NextRow = 2
NextTicker = 2
ColumnAFirstRow = 2

'Populate column headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

'Ticker details
ColumnALastRow = ws.Cells(Rows.Count, TickerColumn).End(xlUp).Row
    ' MsgBox (ColumnALastRow)

For i = ColumnAFirstRow To ColumnALastRow
    If ws.Range("A" & i + 1) <> ws.Range("A" & i) Then
        ws.Range("I" & NextRow) = ws.Range("A" & i)
    
            Lastrow_by_ticker = ws.Range("F" & i)
            Firstrow_by_ticker = ws.Range("C" & NextTicker)
            Ticker_close = Lastrow_by_ticker
            Ticker_open = Firstrow_by_ticker
    
        ws.Range("J" & NextRow) = Ticker_close - Ticker_open
        ws.Range("J" & NextRow).NumberFormat = "0.00"
        ws.Range("K" & NextRow) = (ws.Range("F" & i) - ws.Range("C" & NextTicker)) / ws.Range("C" & NextTicker)
        ws.Range("K" & NextRow).NumberFormat = "0.00%"
        ws.Range("L" & NextRow) = WorksheetFunction.Sum(Range("G" & NextTicker, "G" & i))
        ws.Range("L" & NextRow).NumberFormat = "0"
        
        NextRow = NextRow + 1
        NextTicker = i + 1
    End If
Next i

'Populate category names
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Value"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"

Dim TickerValue As Long
Dim ColumnIStartRow As Long

TickerValue = 9
ColumnIFirstRow = 2

'Greatest % increase, greatest % decrease, greatest stock volume
ColumnILastRow = ws.Cells(Rows.Count, TickerValue).End(xlUp).Row
    ' MsgBox (ColumnILastRow)

' Greatest % Increase lookup
ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K" & ColumnIFirstRow, "K" & ColumnILastRow))
' Greatest % Decrease lookup
ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K" & ColumnIFirstRow, "K" & ColumnILastRow))
' Greatest Total Value lookup
ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L" & ColumnIFirstRow, "L" & ColumnILastRow))
' Format cells Q2 and Q3 as percentage
ws.Range("Q2", "Q3").NumberFormat = "0.00%"
' Range("Q4").NumberFormat = "0"

' Greatest % Increase Ticker symbol lookup
For i = ColumnIFirstRow To ColumnILastRow
    If ws.Range("K" & i).Value = ws.Range("Q2").Value Then
        ws.Range("P2").Value = ws.Range("I" & i).Value
    End If
Next i

' Greatest % Decrease Ticker symbol lookup
For i = ColumnIFirstRow To ColumnILastRow
    If ws.Range("K" & i).Value = ws.Range("Q3").Value Then
        ws.Range("P3").Value = ws.Range("I" & i).Value
    End If
Next i

' Greatest Total Value Ticker symbol lookup
For i = ColumnIFirstRow To ColumnILastRow
    If ws.Range("L" & i).Value = ws.Range("Q4").Value Then
        ws.Range("P4").Value = ws.Range("I" & i).Value
    End If
Next i

' Yearly Change column color format
For i = ColumnIFirstRow To ColumnILastRow
    If ws.Range("J" & i).Value < 0 Then
        ws.Range("J" & i).Interior.Color = vbRed
    ElseIf ws.Range("J" & i).Value > 0 Then
        ws.Range("J" & i).Interior.Color = vbGreen
' No change - Cell Background blank
    Else: ws.Range("J" & i).Interior.Color = xlNone
    End If
Next i
'Autofit column width
ws.Columns("A:Q").AutoFit
Next ws
End Sub


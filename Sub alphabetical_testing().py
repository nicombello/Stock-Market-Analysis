Sub alphabetical_testing()
Dim ticker As String
Dim open_value, close_value, percent_change, volume, diff As Double
Dim start, j As Integer

For Each ws In Worksheets

'Input Titles
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

volume = 0
start = 2
j = 2

'Detect where row ends
last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To last_row

If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
'Is this the last row for the ticker we're on?

volume = volume + ws.Cells(i, 7).Value
ticker = ws.Cells(i, 1).Value
ws.Cells(j, 9).Value = ticker
ws.Cells(j, 12).Value = volume

open_price = ws.Cells(start, 3).Value
close_price = ws.Cells(i, 6).Value

diff = close_price - open_price

ws.Range("J" & j).Value = diff
ws.Range("K" & j).Value = diff / open_price

If open_price = 0 Then
open_price = 1

If ws.Range("K" & j).Value > 0 Then ws.Range("K" & j).Interior.ColorIndex = 3

start = i + 1
volume = 0
j = j + 1

Else

volume = volume + ws.Cells(i, 7).Value
 
End If

End If

Next i

Next ws

End Sub
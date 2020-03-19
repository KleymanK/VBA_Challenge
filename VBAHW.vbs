# VBA_Challenge

This is my VBA homework

Sub stockData()

On Error Resume Next

Dim ws As Worksheet
For Each ws In Worksheets
ws.Activate

Dim lastRow As Long
Dim x_summary As Long
Dim Total As Double
Dim Opening As Double
Dim Closing As Double

x_summary = 2
Total = 0

lastRow = Cells(Rows.Count, "A").End(xlUp).Row

On Error Resume Next

For x = 2 To lastRow
If Cells(x + 1, 1).Value <> Cells(x, 1).Value Then

Total = Total + Cells(x, 7).Value
Closing = Cells(x, 6).Value

Cells(x_summary, 9).Value = Cells(x, 1).Value
Cells(x_summary, 12).Value = Total
Cells(x_summary, 10).Value = Closing - Opening
Cells(x_summary, 11).Value = (Closing - Opening) / Opening


Total = 0
x_summary = x_summary + 1

Else: Total = Total + Cells(x, 7).Value

End If


If Cells(x - 1, 1).Value <> Cells(x, 1).Value Then
    
        Opening = Cells(x, 3).Value

End If
Next x

lastRow = Cells(Rows.Count, "K").End(xlUp).Row
For y = 2 To lastRow
If Cells(y, 11).Value >= 0 Then

Cells(y, 11).Interior.ColorIndex = 4

Else: Cells(y, 11).Interior.ColorIndex = 3

End If

Next y
Next

End Sub



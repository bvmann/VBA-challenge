Attribute VB_Name = "Module1"
Sub tickersymbol()

Dim ticksymbol As String

ticker = 2

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow



ticksymbol = Cells(i, 1).Value

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
Range("H" & ticker).Value = ticksymbol


ticker = ticker + 1


End If

Next i


End Sub



Attribute VB_Name = "Module1"
Sub stonks()


For Each ws In Worksheets

Dim ticker As Integer
Dim yearlychange As Integer
Dim closeprice As Double
Dim openprice As Double
Dim openpricecount As Long
Dim percentchangeyr As Integer
Dim totalstock As Double


ticker = 2

yearlychange = 2

openpricecount = 2

percentchangeyr = 2

totalstock = 0

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


ws.Range("h1").value = "ticker symbol"
ws.Range("i1").value = "yearly change"
ws.Range("j1").value = "percent change"
ws.Range("k1").value = "total stock volume"
ws.Range("J2:J" & lastrow).NumberFormat = "0.00%"
For i = 2 To lastrow


closeprice = ws.Cells(i, 6).value


ticksymbol = ws.Cells(i, 1).value

openprice = ws.Cells(openpricecount, 3).value

totalstock = totalstock + ws.Cells(i, 7)

If ws.Cells(i + 1, 1).value <> ws.Cells(i, 1).value Then
ws.Range("H" & ticker).value = ticksymbol

ws.Range("I" & yearlychange).value = (closeprice - openprice)

ws.Range("k" & yearlychange).value = totalstock

    If openprice <> 0 Then
    
    ws.Range("j" & percentchangeyr).value = ws.Range("I" & yearlychange).value / openprice
    
    End If
    
ticker = ticker + 1

yearlychange = yearlychange + 1

percentchangeyr = percentchangeyr + 1

openpricecount = 1 + i

totalstock = 0

End If



Next i

lastrow2 = ws.Cells(Rows.Count, 8).End(xlUp).Row

For i = 2 To lastrow2

If ws.Cells(i, 9).value > 0 Then

ws.Cells(i, 9).Interior.ColorIndex = 4

Else

ws.Cells(i, 9).Interior.ColorIndex = 3

End If

Next i

Next ws


End Sub







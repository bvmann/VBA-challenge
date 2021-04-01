Attribute VB_Name = "Module1"
Sub stonks()



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

lastrow = Cells(Rows.Count, 1).End(xlUp).Row


Range("h1").Value = "ticker symbol"
Range("i1").Value = "yearly change"
Range("j1").Value = "percent change"
Range("k1").Value = "total stock volume"
Range("J2:J" & lastrow).NumberFormat = "0.00%"
For i = 2 To lastrow


closeprice = Cells(i, 6).Value


ticksymbol = Cells(i, 1).Value

openprice = Cells(openpricecount, 3).Value

totalstock = totalstock + Cells(i, 7)

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
Range("H" & ticker).Value = ticksymbol

Range("I" & yearlychange).Value = (closeprice - openprice)

Range("k" & yearlychange).Value = totalstock
    
    If openprice <> 0 Then

    Range("j" & percentchangeyr).Value = Range("I" & yearlychange).Value / openprice

    End If
    

ticker = ticker + 1

yearlychange = yearlychange + 1

percentchangeyr = percentchangeyr + 1

openpricecount = 1 + i

totalstock = 0

End If



Next i

lastrow2 = Cells(Rows.Count, 8).End(xlUp).Row

For i = 2 To lastrow2

If Cells(i, 9).Value > 0 Then

Cells(i, 9).Interior.ColorIndex = 4

Else

Cells(i, 9).Interior.ColorIndex = 3

End If

Next i


End Sub


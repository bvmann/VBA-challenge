Attribute VB_Name = "Module2"
Sub bonus()

For Each ws In Worksheets


Dim maxtick As String
Dim maxpercent As Double

Dim minpercent As Double
Dim mintick As String

Dim maxvolume As Double
Dim voltick As String


maxtick = ws.Cells(2, 1).value
maxpercent = ws.Cells(2, 10).value

mintick = ws.Cells(2, 1).value
minpercent = ws.Cells(2, 10).value

maxvolume = ws.Cells(2, 11).value
voltick = ws.Cells(2, 1).value

ws.Range("p1").value = "value"
ws.Range("o1").value = "ticker"

ws.Range("n2").value = "Greatest % increase"
ws.Range("n3").value = "Greatest % decrease"
ws.Range("n4").value = "Greatest total volume"

ws.Range("p2").NumberFormat = "0.00%"
ws.Range("p3").NumberFormat = "0.00%"
lastrow = ws.Cells(Rows.Count, 8).End(xlUp).Row

For i = 2 To lastrow

    If ws.Cells(i, 10).value > maxpercent Then
    
    maxpercent = ws.Cells(i, 10).value
    maxtick = ws.Cells(i, 1).value
    End If
Next i

    ws.Range("p2").value = maxpercent
    ws.Range("o2").value = maxtick
    
For i = 2 To lastrow
    If ws.Cells(i, 10).value < minpercent Then
    
    minpercent = ws.Cells(i, 10).value
    mintick = ws.Cells(i, 1).value
    End If
Next i

    ws.Range("p3").value = minpercent
    ws.Range("o3").value = mintick


For i = 2 To lastrow
    If ws.Cells(i, 12).value > maxvolume Then
    
    maxvolume = ws.Cells(i, 12).value
    voltick = ws.Cells(i, 1).value
    End If
Next i

ws.Range("p4").value = maxvolume
ws.Range("o4").value = voltick

Next ws

    
    

End Sub

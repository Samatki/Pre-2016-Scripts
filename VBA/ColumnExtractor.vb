Sub FillDataSheet()

Dim i As Integer
Dim k As Integer
Dim l As Integer

Dim Columnsa(9) As Integer

Columnsa(1) = 1
Columnsa(2) = 2
Columnsa(3) = 4
Columnsa(4) = 6
Columnsa(5) = 8
Columnsa(6) = 12
Columnsa(7) = 14
Columnsa(8) = 9
Columnsa(9) = 10

'For i = 1 To 9
'Worksheets("Test2").Activate
'Range("A" & i) = Columnsa(i)
'Next i


For i = 1 To 16
For l = 1 To 9

k = Columnsa(l)

Worksheets("Rawdata").Activate

ActiveSheet.Range(Cells(1, k + 14 * i - 14), Cells(1, k + 14 * i - 14).End(xlDown)).Copy _
Destination:=Worksheets("Data").Range(Cells(3, l + 10 * i - 10).Address)
    
Next l

Next i

End Sub

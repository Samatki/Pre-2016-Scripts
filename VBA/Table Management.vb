Sub hidecolumns()

Dim c1 As String
Dim c2 As String
Dim cT As String

For i = 0 To 23

c1 = ActiveSheet.Cells(1, (i * 6) + 5).Address
c2 = ActiveSheet.Cells(1, (i * 6) + 7).Address
cT = c1 & ":" & c2

ActiveSheet.Range(cT).EntireColumn.Hidden = True

Next i

End Sub

Sub FillTable()

Dim Title As String

diameters = Array(154.1, 203.3, 254.5, 298.5, 325.4, 368.2)
Tdiameters = Array("6in", "8in", "10in", "12in", "14in", "16in")
uval = Array(3, 0.7, 1, 6)
wT = Array(7.1, 7.9, 9.3, 12.7, 15.1, 19.1)

Sheet4.Activate

For i = 1 To 6
For j = 1 To 4

Title = Tdiameters(i - 1) & " / " & uval(j - 1) & "W/m2K"
ActiveSheet.Cells(6, (i - 1) * 24 + (j - 1) * 6 + 2) = Title

ActiveSheet.Cells(10, (i - 1) * 24 + (j - 1) * 6 + 2) = diameters(i - 1)
ActiveSheet.Cells(11, (i - 1) * 24 + (j - 1) * 6 + 2) = wT(i - 1)

ActiveSheet.Cells(63, (i - 1) * 24 + (j - 1) * 6 + 2).GoalSeek Goal:=uval(j - 1), ChangingCell:=Range(Cells(47, (i - 1) * 24 + (j - 1) * 6 + 2).Address)

ActiveSheet.Cells(47, (i - 1) * 24 + (j - 1) * 6 + 2).Copy Destination:=Sheet5.Cells(6 * (i) + (j - 1), 3)

Next j
Next i

End Sub

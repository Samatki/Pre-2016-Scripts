Sub spacer()

ActiveSheet.Range("U1").EntireColumn.Insert
ActiveSheet.Range("AP1").EntireColumn.Insert
ActiveSheet.Range("BK1").EntireColumn.Insert
ActiveSheet.Range("A:B").EntireColumn.Cut
ActiveSheet.Range("I:I").Insert
ActiveSheet.Range("V:W").EntireColumn.Cut
ActiveSheet.Range("AD:AD").Insert
ActiveSheet.Range("AQ:AR").EntireColumn.Cut
ActiveSheet.Range("AY:AY").Insert
ActiveSheet.Range("BL:BM").EntireColumn.Cut
ActiveSheet.Range("BT:BT").Insert

End Sub

Sub DataCopier()
Dim sname As String
Dim ws As Worksheet

For Each ws In Workbooks("Raw Data2.xlsm").Worksheets

sname = ws.Name

ws.Range("A5:CE17286").Copy _
Destination:=Workbooks("Cooldown Results P2.xlsm").Worksheets(sname).Range("A5")

Next

End Sub


Sub sheetmaker()

Dim SheetName As String
Dim i As Integer
Dim l As Integer
Dim shData As Worksheet

Set shData = ThisWorkbook.Sheets("Data")

For i = 1 To 8

SheetName = Worksheets("Case List").Range("C" & i)
'Worksheets(SheetName).Delete

Worksheets("Template").Copy before:=Worksheets("Results")

ActiveSheet.Name = SheetName
l = 10 * i - 10

shData.Range(shData.Cells(4, 1 + l), shData.Cells(4, 5 + l).End(xlDown)).Copy _
Destination:=Worksheets(SheetName).Range("A7")

shData.Range(shData.Cells(4, 6 + l), shData.Cells(4, 7 + l).End(xlDown)).Copy _
Destination:=Worksheets(SheetName).Range("AV8")

shData.Range(shData.Cells(4, 8 + l), shData.Cells(4, 9 + l).End(xlDown)).Copy _
Destination:=Worksheets(SheetName).Range("AY7")

Next i

End Sub
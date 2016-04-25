Sub Labelmaker()

Dim JobNumber As String
Dim JobName As String
Dim Company As String
Dim Date1 As String
Dim Date2 As String
Dim noEntries As Integer
Dim k As Integer
Dim countl As Integer
Dim noSheets As Double

Sheet1.Activate

noEntries = Range(Cells(1, 1), Cells(1, 1).End(xlDown)).Count
nolabels = (noEntries / 10)
MsgBox (nolabels)

countl = 1

For k = 1 To WorksheetFunction.RoundUp(nolabels, 0)
On Error Resume Next
Worksheets("Label-" & k).Delete


Worksheets.Add
ActiveSheet.Name = "Label-" & k

Worksheets("Label-" & k).Range("A2:C33").Font.Size = "8"
Worksheets("Label-" & k).Range("A1") = "Archived Front Sheets"
Worksheets("Label-" & k).Range("A1:C1").Merge
Worksheets("Label-" & k).Range("A1").Font.Bold = True
Worksheets("Label-" & k).Range("A1").HorizontalAlignment = xlCenter

Worksheets("Label-" & k).Range("A1:C32").BorderAround _
Weight:=xlThick

For i = 1 To 10

If countl <= noEntries Then

Sheet1.Activate
JobNumber = ActiveSheet.Range(Cells((k - 1) * 10 + i, 1).Address).Value
JobName = ActiveSheet.Range(Cells((k - 1) * 10 + i, 2).Address).Value
Company = ActiveSheet.Range(Cells((k - 1) * 10 + i, 3).Address).Value
Date1 = ActiveSheet.Range(Cells((k - 1) * 10 + i, 4).Address).Value
Date2 = Range(Cells((k - 1) * 10 + i, 5).Address).Value

Worksheets("Label-" & k).Range("A" & (3 * i)) = JobNumber
Worksheets("Label-" & k).Range("C" & (3 * i)) = Date1 & " - " & Date2
Worksheets("Label-" & k).Range("A" & (3 * i) + 1) = JobName
Worksheets("Label-" & k).Range("A" & (3 * i) + 2) = Company

Worksheets("Label-" & k).Range("A" & (3 * i) + 1 & ":C" & (3 * i) + 1).Merge
Worksheets("Label-" & k).Range("A" & (3 * i) + 2 & ":C" & (3 * i) + 2).Merge

Worksheets("Label-" & k).Range("A" & (3 * i) + 1 & ":C" & (3 * i) + 1).HorizontalAlignment = xlLeft
Worksheets("Label-" & k).Range("A" & (3 * i) + 2 & ":C" & (3 * i) + 2).HorizontalAlignment = xlLeft

Worksheets("Label-" & k).Range("C:C").ColumnWidth = 12.71

Worksheets("Label-" & k).Range("A" & (3 * i)).Font.Bold = True
Worksheets("Label-" & k).Range("A" & (3 * i)).Font.Size = "10"

countl = countl + 1

Else
End If

Next i

Next k

End Sub


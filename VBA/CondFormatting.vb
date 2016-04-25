Sub CondFormatting()

Dim i As Integer
Dim j As Integer
Dim dataTable As Range
Dim currentSection As Range

Worksheets("Results").Activate
ActiveSheet.Cells.Select
ActiveSheet.Columns.Hidden = False
ActiveSheet.Rows.Hidden = False

Application.ScreenUpdating = False

Set dataTable = Range(Cells(6, 5).End(xlToRight), Cells(14, 5))

With dataTable
    .FormatConditions.Delete
End With

For i = 1 To 16
For j = 1 To 8

Set currentSection = Range(dataTable(j, 1 + (i - 1) * 8).Address & ":" & dataTable(j, i * 8).Address)

With currentSection.FormatConditions _
   .Add(xlExpression, xlExpression, "=IF(" & currentSection(1, 8).Address & "=3,True,False)")
   .Interior.Color = RGB(255, 255, 153)
End With

With currentSection.FormatConditions _
   .Add(xlExpression, xlExpression, "=IF(" & currentSection(1, 8).Address & "=1,True,False)")
   .Interior.Color = RGB(198, 224, 180)
End With

With currentSection.FormatConditions _
   .Add(xlExpression, xlExpression, "=IF(" & currentSection(1, 8).Address & "=2,True,False)")
   .Interior.Color = RGB(226, 239, 218)
End With

Next j

Next i

For i = 1 To 16
For j = 1 To 9

Set currentSection = Range(dataTable(j, 6 + (i - 1) * 8).Address & ":" & dataTable(j, i * 8).Address)

currentSection.EntireColumn.Hidden = True

Next j
Next i

Application.ScreenUpdating = True

End Sub

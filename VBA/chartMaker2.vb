Sub Chartsa()

Dim chartNamex As String
Dim chartHeader As String
Dim IndeValue As String
Dim DepValue As String
Dim SeriesName As String

Flowcomp = Array("HIGH#2", "Low", "Ref#2", "HIGH#2 Combined", "Low Combined", "Ref#2 Combined")

For i = 1 To 6

Sheet10.Activate

chartNamex = Flowcomp(i - 1) & "Pressure"
chartHeader = "Yeoman to Piper B Flowline Arrival Pressure vs Liquid Flowrate with " & Flowcomp(i - 1) & " Composition, U = 3W/m2K"

IndeValue = "Liquid Flowrate (STB/d)"
DepValue = "Arrival Pressure (barg)"
'DepValue = "Arrival Temperature(" & Chr(167) & "C)"

Charts.Add

With ActiveChart

.Name = chartNamex

Do Until .SeriesCollection.Count = 0
.SeriesCollection(1).Delete
Loop

.HasTitle = True
.chartTitle.Text = chartHeader

.HasLegend = True
.Legend.Position = xlLegendPositionBottom

.ChartType = xlXYScatterLinesNoMarkers

.Axes(xlCategory, xlPrimary).HasTitle = True
.Axes(xlCategory, xlPrimary).AxisTitle.Caption = IndeValue
.Axes(xlCategory).TickLabelPosition = xlTickLabelPositionLow

.Axes(xlValue, xlPrimary).HasTitle = True
.Axes(xlValue, xlPrimary).AxisTitle.Caption = DepValue

End With
        
        For p = 1 To 6
                    
            SeriesName = Sheet10.Cells(9, 3 + p).Value & Chr(34)
            ActiveChart.SeriesCollection.NewSeries
            ActiveChart.SeriesCollection(p).Name = SeriesName
            
            ActiveChart.SeriesCollection(p).Values = "=" & Chr(39) & Sheet10.Name & Chr(39) & "!" & Sheet10.Cells(10 + 69 * (i - 1), 60 + p).Address & ":" & Sheet10.Cells(72 + 69 * (i - 1), 60 + p).Address
            ActiveChart.SeriesCollection(p).XValues = "=" & Chr(39) & Sheet10.Name & Chr(39) & "!" & Sheet10.Cells(10 + 69 * (i - 1), 2).Address & ":" & Sheet10.Cells(72 + 69 * (i - 1), 2).Address
            ActiveChart.SeriesCollection(p).Format.Line.Weight = 2.5
                           
        Next p

Next i

End Sub

Sub Tester()
MsgBox (Sheet10.Name)

End Sub

Sub cleanTable()

Dim currentselection As Range

Sheet10.Activate

For i = 1 To 150
For j = 79 To 109

Set currentselection = ActiveSheet.Cells(j, i)

If currentselection.Interior.Color = RGB(242, 242, 242) Then

With currentselection.FormatConditions _
   .Add(xlExpression, xlExpression, "=ISNA(" & currentselection.Address & ")")
   .Font.Color = RGB(242, 242, 242)
End With

Else

With currentselection.FormatConditions _
   .Add(xlExpression, xlExpression, "=ISNA(" & currentselection.Address & ")")
   .Font.ColorIndex = 2
End With

End If


Next j
Next i

End Sub

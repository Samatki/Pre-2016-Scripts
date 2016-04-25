Sub CommandButton1_Click()

'Creates Profile Plot Graphs (Parameter vs Flowline Lengths)

Dim CheckLine As Range
Dim CheckNames As Range
Dim SeriesNames As Range
Dim Dimensions As Range
Dim ResultColumnNumber As Range

Dim ChartCheck As String
Dim ChartName As String
Dim SeriesName As String
Dim Dimension As String
Dim ProfileSheetName As String
Dim SingleSheetCheck As String
Dim SheetName As String

Dim EndSeriesLine As Integer
Dim StartSeriesLine As Integer
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim ColNum As Integer

'Disables Screen Flickering and Misc Errors
Application.ScreenUpdating = False
Application.DisplayAlerts = False

SheetName = "Profile Charts"
SingleSheetCheck = Worksheets("Processing Sheet").Range("K9").Value
ProfileSheetName = Worksheets("Processing Sheet").Range("C2").Value
StartSeriesLine = Worksheets("Processing Sheet").Range("C3").Value
EndSeriesLine = Worksheets("Processing Sheet").Range("C4").Value

Set Dimensions = Worksheets("Processing Sheet").Range(Range("D18"), Range("D18").End(xlToRight))
Set CheckLine = Worksheets("Processing Sheet").Range(Range("D15"), Range("D15").End(xlToRight))
Set CheckNames = Worksheets("Processing Sheet").Range(Range("D17"), Range("D17").End(xlToRight))
Set SeriesNames = Worksheets("Processing Sheet").Range(Range("A19"), Range("A19").End(xlDown))
Set ResultColumnNumber = Worksheets("Processing Sheet").Range("AQ27").CurrentRegion


'NB: If 'SheetName' Worksheet already exists, Code will delete sheet and repopulate
Call CheckIfSheetExists(SheetName)
Worksheets.Add
ActiveSheet.Name = SheetName
ActiveSheet.Move After:=Worksheets(Worksheets.Count)

'Loops Through Each Recorded Property
For i = 1 To CheckLine.Columns.Count
    
   'Checks to See if chart should be made or not, according to spreadsheet Y/N Value
   If CheckLine(1, i).Value = "N" Then
     
   Else

    Charts.Add
    
    'Clears Any Misc Series
    With ActiveChart
        Do Until .SeriesCollection.Count = 0
                .SeriesCollection(1).Delete
        Loop
    End With
    
    'Chart Name
    ChartName = CheckNames(1, i).Value
    ActiveChart.Name = "P-" & ChartName
    

    
    With ActiveChart
        'Chart Title
        .HasTitle = True
        .ChartTitle.Text = "Profile Results - " & ChartName
    
        'Chart Legend
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        
        'Chart Type
        .ChartType = xlXYScatterLinesNoMarkers
        
        'Chart Axes Titles
        Dimension = Dimensions(1, i).Value
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Caption = "Horizontal Distance (m)"
        .Axes(xlCategory).TickLabelPosition = xlTickLabelPositionLow
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Caption = ChartName & " (" & Dimension & ")"
              
        'Adding Series Data - loops through Series List, corresponding to list item k
        For k = 2 To SeriesNames.Rows.Count
            SeriesName = SeriesNames(k, 1).Value
            ActiveChart.SeriesCollection.NewSeries
            ActiveChart.SeriesCollection(k - 1).Name = SeriesName
            ColNum = ResultColumnNumber(k - 1, i).Value
            ActiveChart.SeriesCollection(k - 1).Values = "=" & Chr(39) & ProfileSheetName & Chr(39) & "!" & Cells(StartSeriesLine, ColNum).Address & ":" & Cells(EndSeriesLine, ColNum).Address
            ActiveChart.SeriesCollection(k - 1).XValues = "=" & Chr(39) & ProfileSheetName & Chr(39) & "!" & Cells(StartSeriesLine, (ColNum - 1)).Address & ":" & Cells(EndSeriesLine, (ColNum - 1)).Address
            ActiveChart.SeriesCollection(k - 1).Format.Line.Weight = 2.5

        
        Next k
        
    End With

    If SingleSheetCheck = "Y" Then
    
         ActiveChart.PlotArea.Select
         ActiveChart.ChartArea.Copy
         ActiveChart.Delete
      
         Worksheets(SheetName).Activate
         ActiveSheet.Paste
    
    Else
    
    End If

End If

Next i

Worksheets(SheetName).Activate

Call ArrangeMyCharts(5)

'ActiveChart.PlotArea.Select

Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub

Private Sub CommandButton2_Click()

'Creates Trend Plot Graphs (Parameter vs Time at a Single Point)

Dim CheckLine As Range
Dim CheckNames As Range
Dim SeriesNames As Range
Dim Dimensions As Range
Dim ResultColumnNumber As Range
Dim PositionNames As Range

Dim ChartCheck As String
Dim ChartName As String
Dim SeriesName As String
Dim Dimension As String
Dim TrendSheetName As String
Dim SingleSheetCheck As String
Dim SheetName As String

Dim EndSeriesLine As Integer
Dim StartSeriesLine As Integer
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim ColNum As Integer
Dim NoPositions As Integer
Dim p As Integer
Dim nColumns As Integer

'Disables Screen Flickering and Misc Errors
Application.ScreenUpdating = False
Application.DisplayAlerts = False

'Initiating Variables
SheetName = "OptA Trend Charts"

SingleSheetCheck = Worksheets("Processing Sheet").Range("O9").Value
TrendSheetName = Worksheets("Processing Sheet").Range("C6").Value
StartSeriesLine = Worksheets("Processing Sheet").Range("C7").Value
EndSeriesLine = Worksheets("Processing Sheet").Range("C8").Value

'Describes Arays in "Processing Sheet" - Uses position 1 as a basis
Set PositionNames = Worksheets("Processing Sheet").Range(Range("AI1"), Range("AI1").End(xlDown))
Set Dimensions = Worksheets("Processing Sheet").Range(Range("D50"), Range("D50").End(xlToRight))
Set CheckLine = Worksheets("Processing Sheet").Range(Range("D47"), Range("D47").End(xlToRight))
Set CheckNames = Worksheets("Processing Sheet").Range(Range("D49"), Range("D49").End(xlToRight))

'NB: If 'SheetName' Worksheet already exists, Code will delete sheet and repopulate
Call CheckIfSheetExists(SheetName)
Worksheets.Add
ActiveSheet.Name = SheetName
ActiveSheet.Move After:=Worksheets(Worksheets.Count)

'Defines No. Positions
NoPositions = PositionNames.Rows.Count

'Loops Through Each Recorded Property
For i = 1 To CheckLine.Columns.Count - 1
      
   'Checks to See if chart should be made or not, according to spreadsheet Y/N Value
   If CheckLine(1, i).Value = "N" Then
   
   '(do nothing/skip)
     
   Else

    For p = 2 To NoPositions

    j = (p - 2) * 30
 
    Set SeriesNames = Worksheets("Processing Sheet").Range(Cells(50, 1), Cells(50, 1).End(xlDown))
    Set ResultColumnNumber = Worksheets("Processing Sheet").Range(Cells(51 + j, 35), Cells(51 + j + SeriesNames.Rows.Count - 1, 64))
        
    Charts.Add
    
    'Clears Any Misc Series
    With ActiveChart
        Do Until .SeriesCollection.Count = 0
                .SeriesCollection(1).Delete
        Loop
    End With
    
    'Chart Name
    ChartName = CheckNames(1, i).Value
    ActiveChart.Name = "T1-" & ChartName & " P" & p - 1
    
    With ActiveChart
    
        'Chart Title
        .HasTitle = True
        .ChartTitle.Text = "Trend Results - " & ChartName & " (Position: " & PositionNames(p, 1) & ")"
    
        'Chart Legend
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        
        'Chart Type
        .ChartType = xlXYScatterLinesNoMarkers
        
        'Chart Axes Titles
        Dimension = Dimensions(1, i).Value
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Caption = "Time (Hours)"
        .Axes(xlCategory).TickLabelPosition = xlTickLabelPositionLow
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Caption = ChartName & " (" & Dimension & ")"
              
        'Adding Series Data - loops through Series List, corresponding to list item k
       For k = 2 To SeriesNames.Rows.Count
          ' If k Mod 2 = 0 Then
           ' Else
            
            SeriesName = SeriesNames(k, 1).Value
            ActiveChart.SeriesCollection.NewSeries
            ActiveChart.SeriesCollection(k - 1).Name = SeriesName
            ColNum = ResultColumnNumber(k - 1, i).Value
            ActiveChart.SeriesCollection(k - 1).Values = "=" & Chr(39) & TrendSheetName & Chr(39) & "!" & Cells(StartSeriesLine, ColNum).Address & ":" & Cells(EndSeriesLine, ColNum).Address
            ActiveChart.SeriesCollection(k - 1).XValues = "=" & Chr(39) & TrendSheetName & Chr(39) & "!" & Cells(StartSeriesLine, (ColNum - 1)).Address & ":" & Cells(EndSeriesLine, (ColNum - 1)).Address
            ActiveChart.SeriesCollection(k - 1).Format.Line.Weight = 2.5
        
           ' End If
        Next k
              
    End With
    
        If SingleSheetCheck = "Y" Then
    
           ActiveChart.PlotArea.Select
           ActiveChart.ChartArea.Copy
           ActiveChart.Delete
      
           Worksheets(SheetName).Activate
           ActiveSheet.Paste
    
        Else
    
        End If
    
    Next p

End If

Next i

Worksheets(SheetName).Activate

Call ArrangeMyCharts(NoPositions - 1)

'ActiveChart.PlotArea.Select

Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub

Private Sub CommandButton3_Click()

'Creates Trend Plot Graphs (Graph for Each Case and Property, Different Positions on same plot)

Dim CheckLine As Range
Dim CheckNames As Range
Dim SeriesNames As Range
Dim Dimensions As Range
Dim ResultColumnNumber As Range
Dim PositionNames As Range

Dim ChartCheck As String
Dim ChartName As String
Dim SeriesName As String
Dim Dimension As String
Dim TrendSheetName As String
Dim ChartHeader As String
Dim SingleSheetCheck As String
Dim SheetName As String

Dim EndSeriesLine As Integer
Dim StartSeriesLine As Integer
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim ColNum As Integer
Dim NoPositions As Integer
Dim p As Integer

'Disables Screen Flickering and Misc Errors
Application.ScreenUpdating = False
Application.DisplayAlerts = False

'Initiating Variables
SheetName = "OptB Trend Charts"

SingleSheetCheck = Worksheets("Processing Sheet").Range("T9").Value
TrendSheetName = Worksheets("Processing Sheet").Range("C6").Value
StartSeriesLine = Worksheets("Processing Sheet").Range("C7").Value
EndSeriesLine = Worksheets("Processing Sheet").Range("C8").Value

'Describes Arays in "Processing Sheet" - Uses position 1 as a basis
Set PositionNames = Worksheets("Processing Sheet").Range(Range("AI1"), Range("AI1").End(xlDown))
Set Dimensions = Worksheets("Processing Sheet").Range(Range("D50"), Range("D50").End(xlToRight))
Set CheckLine = Worksheets("Processing Sheet").Range(Range("D47"), Range("D47").End(xlToRight))
Set CheckNames = Worksheets("Processing Sheet").Range(Range("D49"), Range("D49").End(xlToRight))

'NB: If 'SheetName' Worksheet already exists, Code will delete sheet and repopulate
Call CheckIfSheetExists(SheetName)
Worksheets.Add
ActiveSheet.Name = SheetName
ActiveSheet.Move After:=Worksheets(Worksheets.Count)

'Defines No. Positions
NoPositions = PositionNames.Rows.Count

'Loops Through Each Recorded Property
For i = 1 To CheckLine.Columns.Count
      
   'Checks to See if chart should be made or not, according to spreadsheet Y/N Value
   If CheckLine(1, i).Value = "N" Then
   
   '(do nothing/skip)
     
   Else

    Set SeriesNames = Worksheets("Processing Sheet").Range(Cells(50, 1), Cells(50, 1).End(xlDown))
    
    For k = 2 To SeriesNames.Rows.Count

        Charts.Add
    
        'Clears Any Misc Series
        With ActiveChart
            Do Until .SeriesCollection.Count = 0
                    .SeriesCollection(1).Delete
            Loop
        End With
    
        ChartName = CheckNames(1, i).Value
        ChartHeader = "T2-" & SeriesNames(k, 1).Value
        ActiveChart.Name = ChartHeader
            
        With ActiveChart
            .HasTitle = True
            .ChartTitle.Text = "Trend Results - " & SeriesNames(k, 1).Value & " (" & ChartName & ")"
        
            .HasLegend = True
            .Legend.Position = xlLegendPositionBottom
        
            .ChartType = xlXYScatterLinesNoMarkers
        
            Dimension = Dimensions(1, i).Value
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Caption = "Time (Hours)"
            .Axes(xlCategory).TickLabelPosition = xlTickLabelPositionLow
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Caption = ChartName & " (" & Dimension & ")"
        
          For p = 2 To NoPositions
          
          j = (p - 2) * 30

            Set ResultColumnNumber = Worksheets("Processing Sheet").Range(Cells(51 + j, 35), Cells(51 + j + SeriesNames.Rows.Count - 1, 64))
                SeriesName = PositionNames(p, 1).Value
                ActiveChart.SeriesCollection.NewSeries
                ActiveChart.SeriesCollection(p - 1).Name = SeriesName
                ColNum = ResultColumnNumber(k - 1, i).Value
                ActiveChart.SeriesCollection(p - 1).Values = "=" & Chr(39) & TrendSheetName & Chr(39) & "!" & Cells(StartSeriesLine, ColNum).Address & ":" & Cells(EndSeriesLine, ColNum).Address
                ActiveChart.SeriesCollection(p - 1).XValues = "=" & Chr(39) & TrendSheetName & Chr(39) & "!" & Cells(StartSeriesLine, (ColNum - 1)).Address & ":" & Cells(EndSeriesLine, (ColNum - 1)).Address
                ActiveChart.SeriesCollection(p - 1).Format.Line.Weight = 2.5
                
          Next p
        
        End With
    
        If SingleSheetCheck = "Y" Then
    
            ActiveChart.PlotArea.Select
            ActiveChart.ChartArea.Copy
            ActiveChart.Delete
      
            Worksheets(SheetName).Activate
            ActiveSheet.Paste
    
        Else
    
        End If
    
    Next k
       
    End If
    
Next i

Worksheets(SheetName).Activate

Call ArrangeMyCharts(SeriesNames.Rows.Count)

Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub

Private Sub CommandButton4_Click()

'Plots with Cases on x axis, Final Point on Series on y Axis (all positions)

Dim CheckLine As Range
Dim CheckNames As Range
Dim SeriesNames As Range
Dim Dimensions As Range
Dim ResultColumnRegion As Range
Dim PositionNames As Range
Dim ResultColumn As Range

Dim ChartCheck As String
Dim ChartName As String
Dim SeriesName As String
Dim Dimension As String
Dim TrendSheetName As String
Dim ChartHeader As String
Dim SingleSheetCheck As String
Dim SheetName As String

Dim EndSeriesLine As Integer
Dim StartSeriesLine As Integer
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim NoPositions As Integer
Dim p As Integer

'Disables Screen Flickering and Misc Errors
Application.ScreenUpdating = False
Application.DisplayAlerts = False

'Initiating Variables
SheetName = "OptC Trend Charts"

SingleSheetCheck = Worksheets("Processing Sheet").Range("Y9").Value
TrendSheetName = Worksheets("Processing Sheet").Range("C6").Value
StartSeriesLine = Worksheets("Processing Sheet").Range("C7").Value
EndSeriesLine = Worksheets("Processing Sheet").Range("C8").Value

'Describes Arays in "Processing Sheet" - Uses position 1 as a basis
Set PositionNames = Worksheets("Processing Sheet").Range(Range("AI1"), Range("AI1").End(xlDown))
Set Dimensions = Worksheets("Processing Sheet").Range(Range("D50"), Range("D50").End(xlToRight))
Set CheckLine = Worksheets("Processing Sheet").Range(Range("D46"), Range("D46").End(xlToRight))
Set CheckNames = Worksheets("Processing Sheet").Range(Range("D49"), Range("D49").End(xlToRight))

'Defines No. Positions
NoPositions = PositionNames.Rows.Count

'NB: If "Profile Charts" Worksheet already exists, Code will not run, delete to Restart (delete existing charts as well)
Call CheckIfSheetExists(SheetName)
Worksheets.Add
ActiveSheet.Name = SheetName
ActiveSheet.Move After:=Worksheets(Worksheets.Count)

'Loops Through Each Recorded Property
For i = 1 To CheckLine.Columns.Count
'
   'Checks to See if chart should be made or not, according to spreadsheet Y/N Value
   If CheckLine(1, i).Value = "N" Then

   '(do nothing/skip)

   Else
   
   Charts.Add
   
   'Clears Any Misc Series
    With ActiveChart
           Do Until .SeriesCollection.Count = 0
                    .SeriesCollection(1).Delete
           Loop
    End With
   
   ChartName = CheckNames(1, i).Value
   ChartHeader = ChartName
   ActiveChart.Name = "T3- " & ChartHeader
   
    Set SeriesNames = Worksheets("Processing Sheet").Range(Cells(50, 1), Cells(50, 1).End(xlDown))

        With ActiveChart
            .HasTitle = True
            .ChartTitle.Text = ChartName & " (at End of Trend Series)"
            
            .HasLegend = True
            .Legend.Position = xlLegendPositionBottom

            .ChartType = xlLineMarkers
            
            Dimension = Dimensions(1, i).Value
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Caption = "Case"
            .Axes(xlCategory).TickLabelPosition = xlTickLabelPositionLow
            .Axes(xlCategory).TickLabels.Orientation = 45 ' degrees
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Caption = ChartName & " (" & Dimension & ")"

        For p = 2 To NoPositions

        j = (p - 2) * 30

            Set ResultColumnRegion = Worksheets("Processing Sheet").Cells(51 + j, 81).CurrentRegion
                SeriesName = PositionNames(p, 1).Value
                ActiveChart.SeriesCollection.NewSeries
                ActiveChart.SeriesCollection(p - 1).Name = SeriesName
            Set ResultColumn = Worksheets("Processing Sheet").Range(ResultColumnRegion(1, i), ResultColumnRegion(SeriesNames.Count - 1, i))
                ActiveChart.SeriesCollection(p - 1).Values = "=" & Chr(39) & "Processing Sheet" & Chr(39) & "!" & ResultColumn.Address
                ActiveChart.SeriesCollection(p - 1).XValues = "=" & Chr(39) & "Processing Sheet" & Chr(39) & "!" & Worksheets("Processing Sheet").Cells(51, 1).Address & ":" & Worksheets("Processing Sheet").Cells(50, 1).End(xlDown).Address

          Next p
       End With
       
    If SingleSheetCheck = "Y" Then
    
    ActiveChart.PlotArea.Select
    ActiveChart.ChartArea.Copy
    ActiveChart.Delete
      
    Worksheets(SheetName).Activate
    ActiveSheet.Paste
    
    Else
    
    End If
    
End If

Next i

Worksheets(SheetName).Activate

Call ArrangeMyCharts(4)

Application.ScreenUpdating = True
Application.DisplayAlerts = True


End Sub

Sub ArrangeMyCharts(nColumns As Integer)
    Dim iChart As Long
    Dim nCharts As Long
    Dim dTop As Double
    Dim dLeft As Double
    Dim dHeight As Double
    Dim dWidth As Double

    dTop = 30      ' top of first row of charts
    dLeft = 50    ' left of first column of charts
    dHeight = 450  ' height of all charts
    dWidth = 750   ' width of all charts
    nCharts = ActiveSheet.ChartObjects.Count

    For iChart = 1 To nCharts
        With ActiveSheet.ChartObjects(iChart)
            .Height = dHeight
            .Width = dWidth
            .Top = dTop + Int((iChart - 1) / nColumns) * dHeight
            .Left = dLeft + ((iChart - 1) Mod nColumns) * dWidth
        End With
    Next
End Sub
Sub CheckIfSheetExists(SName As String)

On Error Resume Next
Sheets(SName).Delete
On Error GoTo 0

End Sub
Private Sub CommandButton5_Click()

Call CommandButton1_Click
Call CommandButton2_Click
Call CommandButton3_Click
Call CommandButton4_Click

End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub



'Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'End Sub
Sub Modelmaker()

Sheet2.Activate

If CheckTableErrors = "ERRORS" Then GoTo LastLineOnErrors

Dim SingleBranchObj As New NET32COMLib.ISingleBranchModel

Dim ModelSaveDirectory As String
Dim TemplatePath As String
Dim CurrentProfileRange As Range
Dim l As Integer
Dim CurrentProfileRangeS As String
Dim CurrentProfileRange1 As String
Dim CurrentProfileRange2 As String
Dim sheetname As String
Dim ProfileRangeString As String
Dim CurrentInsulationRange As Range
Dim CurrentInsulationRangeS As String
Dim CurrentInsulationRange1 As String
Dim CurrentInsulationRange2 As String
Dim insulvalues As Variant

'Sets Save Directory
ModelSaveDirectory = Range("ModelSaveDirectory")

For i = 1 To Range("ModelsNumber").Count

If Not Range("ModelsActivate").Cells(i, 1).Value = "Y" Then
    'Do Nothing
    'Add check for excel errors here to also do nothing (and Msgbox?)
Else

'Sets Model Path
TemplatePath = Range("ModelsPath")

'Open Model
SingleBranchObj.OpenModel TemplatePath

'Creates New File For each Model based on Template (No changes yet)
'bOK = SingleBranchObj.SaveModel(ModelSaveDirectory & Range("ModelsName").Cells(i, 1).Value & ".bps")

'********************************************** FLOW CORRELATIONS ************************************************************
'*****************************************************************************************************************************

'Model currently assumes 3 phase for OLGA or a default for other flow correlations - abbreviations are located in PIPESIM help or userdll.dat

Dim Flowcorr As New FLOWCORRELATIONCOMLib.CIFlowCorrelation

'Horizontal Flow Correlation

If Range("ModelsFCHoriz").Cells(i, 1) = "bja" Then
    Flowcorr.SetHorizontalCorrelation "BJA", "BBR"
ElseIf Range("ModelsFCHoriz").Cells(i, 1) = "tulsa" Then
    Flowcorr.SetHorizontalCorrelation "TULSA", "TBB"
ElseIf Range("ModelsFCHoriz").Cells(i, 1) = "OLGA-S 2000 Version 6.2.7 flow correlations, October 2010" Then
    Flowcorr.SetHorizontalCorrelation "OLGA6.2.7", "olga3pe"
ElseIf Range("ModelsFCHoriz").Cells(i, 1) = "OLGA-S 2000 Version 5.3.2 flow correlations, February 2009" Then
    Flowcorr.SetHorizontalCorrelation "OLGA5.3.2", "olga3pd"
ElseIf Range("ModelsFCHoriz").Cells(i, 1) = "OLGA-S 2000 Version 5.3 flow correlations, February 2008" Then
    Flowcorr.SetHorizontalCorrelation "OLGA5.3", "olga3pc"
ElseIf Range("ModelsFCHoriz").Cells(i, 1) = "OLGA-S 2000 Version 5.0 flow correlations, June 2006" Then
    Flowcorr.SetHorizontalCorrelation "OLGA5.0", "olga3pb"
ElseIf Range("ModelsFCHoriz").Cells(i, 1) = "TUFFP Unified 3-Phase (v.2011.1) (override emul visc.)" Then
    Flowcorr.SetHorizontalCorrelation "TUFFP", "TUFFPU3P_e-2"
ElseIf Range("ModelsFCHoriz").Cells(i, 1) = "LedaFlow Point Model (v.1.0.231.1, June 2011)" Then
    Flowcorr.SetHorizontalCorrelation "LEDA 1.0", "LEDA_3P_e-2"
ElseIf Range("ModelsFCHoriz").Cells(i, 1) = "Segregated Flow - GRE Mechanistic Model BP" Then
    Flowcorr.SetHorizontalCorrelation "BPD", "BP1"
Else
    MsgBox ("Model " & Range("ModelsNumber").Cells(i, 1) & " missing Horizontal Flow Correlation")
End If

'Vertical Flow Correlation (no BPD Mode)

If Range("ModelsFCVert").Cells(i, 1) = "bja" Then
    Flowcorr.SetVerticalCorrelation "BJA", "BBR"
ElseIf Range("ModelsFCVert").Cells(i, 1) = "tulsa" Then
    Flowcorr.SetVerticalCorrelation "TULSA", "TBB"
ElseIf Range("ModelsFCVert").Cells(i, 1) = "OLGA-S 2000 Version 6.2.7 flow correlations, October 2010" Then
    Flowcorr.SetVerticalCorrelation "OLGA6.2.7", "olga3pe"
ElseIf Range("ModelsFCVert").Cells(i, 1) = "OLGA-S 2000 Version 5.3.2 flow correlations, February 2009" Then
    Flowcorr.SetVerticalCorrelation "OLGA5.3.2", "olga3pd"
ElseIf Range("ModelsFCVert").Cells(i, 1) = "OLGA-S 2000 Version 5.3 flow correlations, February 2008" Then
    Flowcorr.SetVerticalCorrelation "OLGA5.3", "olga3pc"
ElseIf Range("ModelsFCVert").Cells(i, 1) = "OLGA-S 2000 Version 5.0 flow correlations, June 2006" Then
    Flowcorr.SetVerticalCorrelation "OLGA5.0", "olga3pb"
ElseIf Range("ModelsFCVert").Cells(i, 1) = "TUFFP Unified 3-Phase (v.2011.1) (override emul visc.)" Then
    Flowcorr.SetVerticalCorrelation "TUFFP", "TUFFPU3P_e-2"
ElseIf Range("ModelsFCVert").Cells(i, 1) = "LedaFlow Point Model (v.1.0.231.1, June 2011)" Then
    Flowcorr.SetVerticalCorrelation "LEDA 1.0", "LEDA_3P_e-2"
Else
    MsgBox ("Model " & Range("ModelsNumber").Cells(i, 1) & " missing Vertical Flow Correlation")
End If

'Unused properties - all take a double as input
'Flowcorr.HorizontalFrictionFactor
'Flowcorr.VerticalFrictionFactor
'Flowcorr.HorizontalHoldup
'Flowcorr.VerticalHoldup
'Flowcorr.SwapAngle

'If Single Phase (not currently used)
'Flowcorr.SinglePhaseCorrelation 'As String
'Flowcorr.SinglePhaseFactor 'As Double

'Makes changes to model
SingleBranchObj.FlowCorrelation = Flowcorr

'********************************************** MISC PARAMETERS **************************************************************
'*****************************************************************************************************************************

'Erosional Velocity Constant

'Model 0 = API14e, 1 = SALAMA (2000)
SingleBranchObj.ErosionCorrosion.ErosionModel = 0

'Sets API 14e Erosional Velocity Constant
SingleBranchObj.ErosionCorrosion.ErosionVelocityConst = Range("ModelsEVC").Cells(i, 1)

'********************************************** BLACK OIL PARAMETERS *********************************************************
'*****************************************************************************************************************************

Dim BlackOil As New FLUIDMODELCOMLib.IBlackOil

gortype = Range("ModelsBOGORType")
WCType = Range("ModelsBOWCtype")
APIType = Range("ModelsBOAPItype")

Dim gorval As Integer
Dim WCval As Integer
Dim APIval As Integer

'BO Model GLR
If (gortype = "GLR") Then
    gorval = 0
    BlackOil.GLR_SI = Range("ModelsBOGOR").Cells(i, 1)
ElseIf (gortype = "GOR") Then
    gorval = 1
    BlackOil.GOR_SI = Range("ModelsBOGOR").Cells(i, 1)
ElseIf (gortype = "LGR") Then
    gorval = 2
    BlackOil.LGR_SI = Range("ModelsBOGOR").Cells(i, 1)
ElseIf (gortype = "OGR") Then
    gorval = 3
    BlackOil.OGR_SI = Range("ModelsBOGOR").Cells(i, 1)
End If
BlackOil.GOR_Type = gorval

'BO Model WaterCut
If (WCType = "WCut") Then
    WCval = 0
    BlackOil.Watercut = Range("ModelsBOWatercut").Cells(i, 1)
ElseIf (WCType = "GWR") Then
    WCval = 1
    BlackOil.GWR_SI = Range("ModelsBOWatercut").Cells(i, 1)
ElseIf (WCType = "WGR") Then
    WCval = 2
    BlackOil.WGR_SI = Range("ModelsBOWatercut").Cells(i, 1)
End If
BlackOil.WGR_Type = WCval

'BO Model API
If (APIType = "API") Then
    APIval = 0
    BlackOil.API = Range("ModelsBOAPI").Cells(i, 1)
ElseIf (APIType = "DOD") Then
    APIval = 1
    BlackOil.API = Range("ModelsBOAPI").Cells(i, 1)
End If
BlackOil.API_Type = APIval

'Names Fluid
If Range("ModelsBOFluidName").Cells(i, 1) = "" Then
   BlackOil.Name = "Model " & i & " Fluid"
Else
   BlackOil.Name = Range("ModelsBOFluidName").Cells(i, 1)
End If

'Sets Black Oil Properties
BlackOil.GasSG = Range("ModelsBOGasSG").Cells(i, 1)
BlackOil.WaterSG = Range("ModelsBOWaterSG").Cells(i, 1)
BlackOil.Comment = Range("ModelsBOComment").Cells(i, 1)

'Makes Changes
SingleBranchObj.BlackOil = BlackOil

'*********************************************** SOURCE PROPERTIES ***********************************************************
'*****************************************************************************************************************************

SingleBranchObj.SetPropertyVal "Source", "SOURCE PRESSURE", Range("ModelsSourcePressure").Cells(i, 1), Range("ModelsSourceUnits").Cells(1, 1)
SingleBranchObj.SetPropertyVal "Source", "SOURCE TEMPERATURE", Range("ModelsSourceTemperature").Cells(i, 1), Range("ModelsSourceUnits").Cells(1, 2)

'************************************************ GEOMETRIES *****************************************************************
'*****************************************************************************************************************************

Dim FlowlObj As New NET32COMLib.FlowlineObj
Dim UObj As New NET32COMLib.HeatTransfer

'************************************************* FLOWLINE *******************************************************************

Set FlowlObj = SingleBranchObj.ObjectProperties("Flowline")

If Range("ModelsGeoComplexity").Cells(i, 1) = "Simple" Then

    FlowlObj.UseDetailedProfile False
    
    SingleBranchObj.SetPropertyVal "Flowline", "HORIZONTAL LENGTH", Range("ModelsSimpleLength").Cells(i, 1), Range("ModelsFlSimpleUnits").Cells(1, 1)
    SingleBranchObj.SetPropertyVal "Flowline", "VERTICAL LENGTH", Range("ModelsSimpleElevation").Cells(i, 1), Range("ModelsFlSimpleUnits").Cells(1, 2)
    SingleBranchObj.SetPropertyVal "Flowline", "PIPE ID", Range("ModelsID").Cells(i, 1), Range("ModelsFlSimpleUnits").Cells(1, 4)
    
    SingleBranchObj.SetPropertyVal "Flowline", "PIPE AMB TEMPERATURE", Range("ModelsSimpleAmbTemp").Cells(i, 1), Range("ModelsFlSimpleUnits").Cells(1, 7)
    SingleBranchObj.SetPropertyVal "Flowline", "PIPE ROUGHNESS", Range("ModelsRoughness").Cells(i, 1), Range("ModelsFlSimpleUnits").Cells(1, 6)
    
    If Range("ModelsUMode").Cells(i, 1) = "Calculated" Then
    Else
    SingleBranchObj.SetPropertyVal "Flowline", "PIPE WT", Range("ModelsWT").Cells(i, 1), Range("ModelsFlSimpleUnits").Cells(1, 5)
    End If
    SingleBranchObj.SetPropertyVal "Flowline", "HEIGHT UNDULATIONS", Range("ModelsSimpleROU").Cells(i, 1), ""

ElseIf Range("ModelsGeoComplexity").Cells(i, 1) = "Complex" Then

    FlowlObj.UseDetailedProfile False
    SingleBranchObj.SetPropertyVal "Flowline", "PIPE ID", Range("ModelsID").Cells(i, 1), Range("ModelsFlSimpleUnits").Cells(1, 4)
    SingleBranchObj.SetPropertyVal "Flowline", "PIPE ROUGHNESS", Range("ModelsRoughness").Cells(i, 1), Range("ModelsFlSimpleUnits").Cells(1, 6)
    
If Range("ModelsUMode").Cells(i, 1) = "Calculated" Then
Else
    SingleBranchObj.SetPropertyVal "Flowline", "PIPE WT", Range("ModelsWT").Cells(i, 1), Range("ModelsFlSimpleUnits").Cells(1, 5)
End If

    FlowlObj.UseDetailedProfile True        ' set the flowline to use detailed description
    FlowlObj.ClearDetailedProfile        ' clear all existing nodes

For M = 1 To Sheet3.Range("ProfilesList").Count + 1
        If M < Sheet3.Range("ProfilesList").Count + 1 Then
            If Sheet3.Range("ProfilesList").Cells(M, 1) = Sheet2.Range("ModelsComplexGeo").Cells(i, 1) Then
                CurrentProfileRange2 = Sheet3.Range("ProfileTableReference").Offset(2, (M - 1) * 8).End(xlDown).Address
                CurrentProfileRange1 = Sheet3.Range("ProfileTableReference").Offset(2, (M - 1) * 8).Offset(0, 6).Address
                CurrentProfileRangeS = CurrentProfileRange2 & ":" & CurrentProfileRange1
                Set CurrentProfileRange = Sheet3.Range(CurrentProfileRangeS)
                    If CurrentProfileRange.Rows.Count > 1000 Then
                        MsgBox (Sheet2.Range("ModelsComplexGeo").Cells(i, 1).Value & " Information not entered correctly")
                        GoTo LastLineOnErrors
                    End If
                GoTo ProfileDataSolved:
            End If
            Else
                MsgBox ("Profile Not Present in List - Check Model Geometry Names")
                GoTo LastLineOnErrors
      End If
Next M

ProfileDataSolved:

For k = 1 To CurrentProfileRange.Rows.Count 'all values in strict SI
        FlowlObj.AddProfileNode_SI CurrentProfileRange.Cells(k, 1), CurrentProfileRange.Cells(k, 2), CurrentProfileRange.Cells(k, 3), CurrentProfileRange.Cells(k, 4), CStr(CurrentProfileRange.Cells(k, 7))
Next k
     
Else
    'If Complexity value missing goes to next model - Should be unneccessary
    MsgBox ("Error - Missing Flowline Complexity Value")
    GoTo Lastline
End If

'U-Value
Set UObj = FlowlObj.HeatTransfer
If Range("ModelsUMode").Cells(i, 1) = "Input" Then
    
    UObj.UValueType = 5
    UObj.SetUValue Range("ModelsInputU").Cells(i, 1), Range("FlUUnits").Cells(1, 1)

ElseIf Range("ModelsUMode").Cells(i, 1) = "Calculated" Then

    UObj.UValueType = 0
       
For M = 1 To Range("InsulationsList").Count + 1
        If M < Range("InsulationsList").Count + 1 Then
            If Range("InsulationsList").Cells(M, 1) = Range("ModelsCalcUCoating").Cells(i, 1) Then
                CurrentInsulationRange2 = Range("InsulationTableReference").Offset(2, (M - 1) * 4).End(xlDown).Address
                CurrentInsulationRange1 = Range("InsulationTableReference").Offset(2, (M - 1) * 4).Offset(0, 1).Address
                CurrentInsulationRangeS = CurrentInsulationRange2 & ":" & CurrentInsulationRange1
                Set CurrentInsulationRange = Sheet5.Range(CurrentInsulationRangeS)
                    If Sheet5.Range(CurrentInsulationRangeS).Rows.Count > 12 Then
                        MsgBox (Range("ModelsCalcUCoating").Cells(i, 1).Value & " Information not entered correctly")
                        GoTo LastLineOnErrors
                    End If
                GoTo FLInsulationDataSolved:
            End If
            Else
                MsgBox ("Insulation Not Present in List - Check Model Insulation Names")
                GoTo LastLineOnErrors
      End If
Next M


FLInsulationDataSolved:

Sheet5.Activate
CurrentInsulationRange.Select
MsgBox (CurrentInsulationRange.Address)

insulvalues = arraymaker(Selection)

    'SingleBranchObj.
'    UObj.SetObjectRef(
'
'    UObj.SetCoatingData_SI insulvalues
    
    'MsgBox (UObj.GetCoatingData_SI("Flowline"))
       
    'FlowlObj.HeatTransfer.SetCoatingData_SI insulvalues
        
    UObj.SetPipeConductivity CurrentInsulationRange.End(xlUp).Offset(16, 0), "W/m/K"
    UObj.SetAmbientFluidVelocity Sheet2.Range("ModelsCalcUAmbFluidVel").Cells(i, 1), Sheet2.Range("FlUUnits").Cells(1, 5)
    UObj.SetBurialDepth CurrentInsulationRange.End(xlUp).Offset(17, 1).Value / 1000, "m"
    UObj.SetGroundConductivity CurrentInsulationRange.End(xlUp).Offset(17, 0), "W/m/K"
    SingleBranchObj.SetPropertyVal "Flowline", "PIPE WT", CurrentInsulationRange.End(xlUp).Offset(16, 1), "mm"

    If Range("ModelsCalcUAmbFluid").Cells(i, 1) = "Air" Then
    UObj.AmbientFluidType = 0
    Else
    UObj.AmbientFluidType = 1
    End If

ElseIf Range("ModelsUMode").Cells(i, 1) = "Bare (in Air)" Then
    If Range("ModelsGeoComplexity").Cells(i, 1) = "Complex" Then
    MsgBox ("Model " & Range("ModelsNumber").Cells(i, 1) & "Flowline" & " U value can only have mode 'input' or 'calculated' if geometry is 'complex'")
    GoTo LastLineOnErrors
    Else
    UObj.UValueType = 3
    End If
ElseIf Range("ModelsUMode").Cells(i, 1) = "Bare (in Water)" Then
    If Range("ModelsGeoComplexity").Cells(i, 1) = "Complex" Then
    MsgBox ("Model " & Range("ModelsNumber").Cells(i, 1) & "Flowline" & " U value can only have mode 'input' or 'calculated' if geometry is 'complex'")
    GoTo LastLineOnErrors
    Else
    UObj.UValueType = 4
    End If
ElseIf Range("ModelsUMode").Cells(i, 1) = "Insulated" Then
    If Range("ModelsGeoComplexity").Cells(i, 1) = "Complex" Then
    MsgBox ("Model " & Range("ModelsNumber").Cells(i, 1) & "Flowline" & " U value can only have mode 'input' or 'calculated' if geometry is 'complex'")
    GoTo LastLineOnErrors
    Else
    UObj.UValueType = 1
    End If
ElseIf Range("ModelsUMode").Cells(i, 1) = "Coated" Then
    If Range("ModelsGeoComplexity").Cells(i, 1) = "Complex" Then
    MsgBox ("Model " & Range("ModelsNumber").Cells(i, 1) & "Flowline" & " U value can only have mode 'input' or 'calculated' if geometry is 'complex'")
    GoTo LastLineOnErrors
    Else
    UObj.UValueType = 2
    End If
End If


'************************************************* RISER 1 *******************************************************************

Set FlowlObj = SingleBranchObj.ObjectProperties("Riser1")

If Range("ModelsR1Mode").Cells(i, 1) = "Simple" Then

    FlowlObj.UseDetailedProfile False

    SingleBranchObj.SetPropertyVal "Riser1", "HORIZONTAL LENGTH", Range("ModelsR1SimpleLength").Cells(i, 1), Range("ModelsR1SimpleUnits").Cells(1, 1)
    SingleBranchObj.SetPropertyVal "Riser1", "VERTICAL LENGTH", Range("ModelsR1SimpleElevation").Cells(i, 1), Range("ModelsR1SimpleUnits").Cells(1, 2)
    SingleBranchObj.SetPropertyVal "Riser1", "PIPE ID", Range("ModelsR1ID").Cells(i, 1), Range("ModelsR1SimpleUnits").Cells(1, 3)

    SingleBranchObj.SetPropertyVal "Riser1", "PIPE AMB TEMPERATURE", Range("ModelsR1AmbTemp").Cells(i, 1), Range("ModelsR1SimpleUnits").Cells(1, 6)
    SingleBranchObj.SetPropertyVal "Riser1", "PIPE ROUGHNESS", Range("ModelsR1Rough").Cells(i, 1), Range("ModelsR1SimpleUnits").Cells(1, 5)
    SingleBranchObj.SetPropertyVal "Riser1", "PIPE WT", Range("ModelsR1WT").Cells(i, 1), Range("ModelsR1SimpleUnits").Cells(1, 4)

ElseIf Range("ModelsR1mode").Cells(i, 1) = "Complex" Then
    
    'Fudge for easier data entry - enters ID, Roughness and WT in 'Simple' geometry mode first, then switches to detailed geometry mode for profile dimension entry.
    FlowlObj.UseDetailedProfile False
    SingleBranchObj.SetPropertyVal "Riser1", "PIPE ID", Range("ModelsR1ID").Cells(i, 1), Range("ModelsR1SimpleUnits").Cells(1, 3)
    SingleBranchObj.SetPropertyVal "Riser1", "PIPE ROUGHNESS", Range("ModelsR1Rough").Cells(i, 1), Range("ModelsR1SimpleUnits").Cells(1, 5)
    SingleBranchObj.SetPropertyVal "Riser1", "PIPE WT", Range("ModelsR1WT").Cells(i, 1), Range("ModelsR1SimpleUnits").Cells(1, 4)

    FlowlObj.UseDetailedProfile True        ' set the flowline to use detailed description
    FlowlObj.ClearDetailedProfile        ' clear all existing nodes

'Gets Flowline Geometry from Model Geometry sheet
For M = 1 To Sheet3.Range("ProfilesList").Count + 1
        If M < Sheet3.Range("ProfilesList").Count + 1 Then
            If Sheet3.Range("ProfilesList").Cells(M, 1) = Sheet2.Range("ModelsR1ComplexGeo").Cells(i, 1) Then
                CurrentProfileRange2 = Sheet3.Range("ProfileTableReference").Offset(2, (M - 1) * 8).End(xlDown).Address
                CurrentProfileRange1 = Sheet3.Range("ProfileTableReference").Offset(2, (M - 1) * 8).Offset(0, 6).Address
                CurrentProfileRangeS = CurrentProfileRange2 & ":" & CurrentProfileRange1
                Set CurrentProfileRange = Sheet3.Range(CurrentProfileRangeS)
                    If CurrentProfileRange.Rows.Count > 1000 Then
                        MsgBox (Sheet2.Range("ModelsR1ComplexGeo").Cells(i, 1).Value & " Information not entered correctly")
                        GoTo LastLineOnErrors
                    End If
                GoTo ProfileDataSolvedR1:
            End If
            Else
                MsgBox ("Profile Not Present in List - Check Model Geometry Names")
                GoTo LastLineOnErrors
      End If
Next M

ProfileDataSolvedR1:

For k = 1 To CurrentProfileRange.Rows.Count 'all values in strict SI
        FlowlObj.AddProfileNode_SI CurrentProfileRange.Cells(k, 1), CurrentProfileRange.Cells(k, 2), CurrentProfileRange.Cells(k, 3), CurrentProfileRange.Cells(k, 4), CStr(CurrentProfileRange.Cells(k, 7))
Next k
     
Else
    'If Complexity value missing - Should be unneccessary
    MsgBox ("Error - Missing Riser 1 Complexity Value")
    GoTo Lastline
End If

'U -Value
Set UObj = FlowlObj.HeatTransfer
If Range("ModelsR1UMode").Cells(i, 1) = "Input" Then

    UObj.UValueType = 5
    UObj.SetUValue Range("ModelsR1UInput").Cells(i, 1), Range("ModelsR1UUnits").Cells(1, 1)

ElseIf Range("ModelsR1UMode").Cells(i, 1) = "Calculated" Then

    UObj.UValueType = 0

    'UObj.SetPipeConductivity Range("ModelsCalcUPipeCond").Cells(i, 1), Range("FlUUnits").Cells(1, 4)
    UObj.SetAmbientFluidVelocity Range("ModelsR1AmbFluidVel").Cells(i, 1), Range("ModelsR1UUnits").Cells(1, 5)

    'Sets a pipe coating data. Data has to be passed as an array
    'where first column containts layerÅ]specific conductivity values
    'and a second column containts layer thicknesses.
    'Fails if UValueType is not calculated
    'UObj.SetCoatingData_SI

    If Range("ModelsR1AmbFluid").Cells(i, 1) = "Air" Then
    UObj.AmbientFluidType = 0
    Else
    UObj.AmbientFluidType = 1
    End If

ElseIf Range("ModelsR1UMode").Cells(i, 1) = "Bare (in Air)" Then
    UObj.UValueType = 3
ElseIf Range("ModelsR1UMode").Cells(i, 1) = "Bare (in Water)" Then
    UObj.UValueType = 4
ElseIf Range("ModelsR1UMode").Cells(i, 1) = "Insulated" Then
    UObj.UValueType = 1
ElseIf Range("ModelsR1UMode").Cells(i, 1) = "Coated" Then
    UObj.UValueType = 2
End If

'************************************************* RISER 2 *******************************************************************

Set FlowlObj = SingleBranchObj.ObjectProperties("Riser2")

If Range("ModelsR2Mode").Cells(i, 1) = "Simple" Then

    FlowlObj.UseDetailedProfile False

    SingleBranchObj.SetPropertyVal "Riser2", "HORIZONTAL LENGTH", Range("ModelsR2SimpleLength").Cells(i, 1), Range("ModelsR2SimpleUnits").Cells(1, 1)
    SingleBranchObj.SetPropertyVal "Riser2", "VERTICAL LENGTH", Range("ModelsR2SimpleElevation").Cells(i, 1), Range("ModelsR2SimpleUnits").Cells(1, 2)
    SingleBranchObj.SetPropertyVal "Riser2", "PIPE ID", Range("ModelsR2ID").Cells(i, 1), Range("ModelsR2SimpleUnits").Cells(1, 3)

    SingleBranchObj.SetPropertyVal "Riser2", "PIPE AMB TEMPERATURE", Range("ModelsR2AmbTemp").Cells(i, 1), Range("ModelsR2SimpleUnits").Cells(1, 6)
    SingleBranchObj.SetPropertyVal "Riser2", "PIPE ROUGHNESS", Range("ModelsR2Rough").Cells(i, 1), Range("ModelsR2SimpleUnits").Cells(1, 5)
    SingleBranchObj.SetPropertyVal "Riser2", "PIPE WT", Range("ModelsR2WT").Cells(i, 1), Range("ModelsR2SimpleUnits").Cells(1, 4)

ElseIf Range("ModelsR2Mode").Cells(i, 1) = "Complex" Then

    FlowlObj.UseDetailedProfile False
    SingleBranchObj.SetPropertyVal "Riser2", "PIPE ID", Range("ModelsR2ID").Cells(i, 1), Range("ModelsR2SimpleUnits").Cells(1, 3)
    SingleBranchObj.SetPropertyVal "Riser2", "PIPE ROUGHNESS", Range("ModelsR2Rough").Cells(i, 1), Range("ModelsR2SimpleUnits").Cells(1, 5)
    SingleBranchObj.SetPropertyVal "Riser2", "PIPE WT", Range("ModelsR2WT").Cells(i, 1), Range("ModelsR2SimpleUnits").Cells(1, 4)

    FlowlObj.UseDetailedProfile True        ' set the flowline to use detailed description
    FlowlObj.ClearDetailedProfile        ' clear all existing nodes

For M = 1 To Sheet3.Range("ProfilesList").Count + 1
        If M < Sheet3.Range("ProfilesList").Count + 1 Then
            If Sheet3.Range("ProfilesList").Cells(M, 1) = Sheet2.Range("ModelsR2ComplexGeo").Cells(i, 1) Then
                CurrentProfileRange2 = Sheet3.Range("ProfileTableReference").Offset(2, (M - 1) * 8).End(xlDown).Address
                CurrentProfileRange1 = Sheet3.Range("ProfileTableReference").Offset(2, (M - 1) * 8).Offset(0, 6).Address
                CurrentProfileRangeS = CurrentProfileRange2 & ":" & CurrentProfileRange1
                Set CurrentProfileRange = Sheet3.Range(CurrentProfileRangeS)
                    If CurrentProfileRange.Rows.Count > 1000 Then
                        MsgBox (Sheet2.Range("ModelsR2ComplexGeo").Cells(i, 1).Value & " Information not entered correctly")
                        GoTo LastLineOnErrors
                    End If
                GoTo ProfileDataSolvedR2:
            End If
            Else
                MsgBox ("Profile Not Present in List - Check Model Geometry Names")
                GoTo LastLineOnErrors
      End If
Next M

ProfileDataSolvedR2:

For k = 1 To CurrentProfileRange.Rows.Count 'all values in strict SI
        FlowlObj.AddProfileNode_SI CurrentProfileRange.Cells(k, 1), CurrentProfileRange.Cells(k, 2), CurrentProfileRange.Cells(k, 3), CurrentProfileRange.Cells(k, 4), CStr(CurrentProfileRange.Cells(k, 7))
Next k
     
Else
    'If Complexity value missing goes to next model - Should be unneccessary
    MsgBox ("Error - Missing Riser 2 Complexity Value")
    GoTo Lastline
End If

'U -Value
Set UObj = FlowlObj.HeatTransfer
If Range("ModelsR2UMode").Cells(i, 1) = "Input" Then

    UObj.UValueType = 5
    UObj.SetUValue Range("ModelsR2InputU").Cells(i, 1), Range("ModelsR2UUnits").Cells(1, 1)

ElseIf Range("ModelsR2UMode").Cells(i, 1) = "Calculated" Then

    UObj.UValueType = 0

    'UObj.SetPipeConductivity Range("ModelsCalcUPipeCond").Cells(i, 1), Range("ModelsR2UUnits").Cells(1, 4)
    UObj.SetAmbientFluidVelocity Range("ModelsR2FluidVel").Cells(i, 1), Range("ModelsR2UUnits").Cells(1, 5)

    'Sets a pipe coating data. Data has to be passed as an array
    'where first column containts layerÅ]specific conductivity values
    'and a second column containts layer thicknesses.
    'Fails if UValueType is not calculated
    'UObj.SetCoatingData_SI

    If Range("ModelsR2AmbFluid").Cells(i, 1) = "Air" Then
    UObj.AmbientFluidType = 0
    Else
    UObj.AmbientFluidType = 1
    End If

ElseIf Range("ModelsR2UMode").Cells(i, 1) = "Bare (in Air)" Then
    UObj.UValueType = 3
ElseIf Range("ModelsR2UMode").Cells(i, 1) = "Bare (in Water)" Then
    UObj.UValueType = 4
ElseIf Range("ModelsR2UMode").Cells(i, 1) = "Insulated" Then
    UObj.UValueType = 1
ElseIf Range("ModelsR2UMode").Cells(i, 1) = "Coated" Then
    UObj.UValueType = 2
End If

'*****************************************************************************************************************************
'*****************************************************************************************************************************

Dim q As Integer
q = i

'Call SystemAnalysisRun(q, SingleBranchObj)
'Final Save

Dim bOK As Boolean
bOK = SingleBranchObj.SaveModel(ModelSaveDirectory & Range("ModelsName").Cells(i, 1).Value & ".bps")

End If

Next i

Lastline:
LastLineOnErrors:

End Sub

Function CheckTableErrors()

Dim rng As Range
Set rng = Sheet2.Range("ModelsCellsChecker")

Dim v As Integer
Dim w As Integer
Dim str1 As String
Dim str2 As String

For v = 1 To rng.Rows.Count

If Range("ModelsActivate").Cells(v, 1) = "Y" Then

For w = 1 To 80
    
    If rng.Cells(v, w).FormatConditions.Count = 0 Then
        
    ElseIf rng.Cells(v, w).FormatConditions.Count = 1 Then
      str1 = rng.Cells(v, w).FormatConditions(1).Formula1
        If Evaluate(str1) = True Then
            MsgBox ("Missing Values in Model3 " & Range("ModelsNumber").Cells(v, 1) & ": " & rng.Cells(v, w).Offset(-(2 + v), 0))
            CheckTableErrors = "ERRORS"
            GoTo CheckerLastline
        End If

    ElseIf rng.Cells(v, w).FormatConditions.Count = 2 Then

        str1 = rng.Cells(v, w).FormatConditions(1).Formula1
        str2 = rng.Cells(v, w).FormatConditions(2).Formula1

        If Evaluate(str1) = False And Evaluate(str2) = True Then
          MsgBox ("Missing Values in Model " & Range("ModelsNumber").Cells(v, 1) & ": " & rng.Cells(v, w).Offset(-(2 + v), 0))
          CheckTableErrors = "ERRORS"
          GoTo CheckerLastline
        End If
    End If

Next w

End If

Next v

CheckerLastline:

End Function
Function arraymaker(rng As Range)

Dim xx, yy As Integer
xx = rng.Rows.Count - 1
yy = rng.Columns.Count - 1

Dim arraym() As Double
ReDim arraym(xx, yy)

For i = 0 To xx
For j = 0 To yy

arraym(j, i) = rng(i + 1, j + 1).Value
MsgBox (arraym(j, i))
Next j
Next i

For i = 0 To xx
    arraym(1, i) = arraym(1, i) / 1000
Next

MsgBox (arraym(0, 0))

arraymaker = arraym

End Function



Sub SystemAnalysisRun(i As Integer, SingleBranchObj As NET32COMLib.ISingleBranchModel)

'Sets Operation to System Analysis
SingleBranchObj.SetOperationType 0

Dim SysA As New PSOPSYSTEMSLib.ISystemsAnalysis

If Range("SysAnalysisCalculatedVariable1") = "Inlet Pressure" Then
    SysA.BoundaryConds.CalculatedVariable = 1
    If Range("SysAnalysisRateMode1") = "Liquid Rate" Then
        SysA.BoundaryConds.FluidType = 0
    ElseIf Range("SysAnalysisRateMode1") = "Gas Rate" Then
        SysA.BoundaryConds.FluidType = 1
    ElseIf Range("SysAnalysisRateMode1") = "Mass Rate" Then
        SysA.BoundaryConds.FluidType = 2
    End If

    SysA.BoundaryConds.FluidRate_SI = Range("SysAnalysisFluidRate1").Value
    SysA.BoundaryConds.OutletPressure_SI = Range("SysAnalysisOutletPressure1").Value

ElseIf Range("SysAnalysisCalculatedVariable1") = "Outlet Pressure" Then
    SysA.BoundaryConds.CalculatedVariable = 0
    If Range("SysAnalysisRateMode1") = "Liquid Rate" Then
        SysA.BoundaryConds.FluidType = 0
    ElseIf Range("SysAnalysisRateMode1") = "Gas Rate" Then
        SysA.BoundaryConds.FluidType = 1
    ElseIf Range("SysAnalysisRateMode1") = "Mass Rate" Then
        SysA.BoundaryConds.FluidType = 2
    End If

    SysA.BoundaryConds.FluidRate_SI = Range("SysAnalysisFluidRate1").Value

ElseIf Range("SysAnalysisCalculatedVariable1") = "Fluid Rate" Then
    SysA.BoundaryConds.CalculatedVariable = 2
    If Range("SysAnalysisRateMode1") = "Liquid Rate" Then
        SysA.BoundaryConds.FluidType = 0
    ElseIf Range("SysAnalysisRateMode1") = "Gas Rate" Then
        SysA.BoundaryConds.FluidType = 1
    ElseIf Range("SysAnalysisRateMode1") = "Mass Rate" Then
        SysA.BoundaryConds.FluidType = 2
    End If

    SysA.BoundaryConds.OutletPressure_SI = Range("SysAnalysisOutletPressure1").Value

Else
    MsgBox ("Missing System Analysis Data")
End If

'Sets System Analysis Permutation Mode

Dim PermuteMode As String
Dim PermuteModeNo As Long
Dim MaxSysAnalysisCol As Integer
Dim sensitivitysloop As Variant

PermuteMode = Range("SysAnalysisPermuteStepMode")

If PermuteMode = "Permuted against each other" Then
    PermuteModeNo = 0
    MaxSysAnalysisCol = 4
ElseIf PermuteMode = "Change in step with Sens Var 1" Then
    PermuteModeNo = 1
    MaxSysAnalysisCol = 10
ElseIf PermuteMode = "Change in step with x axis" Then
    PermuteModeNo = 2
    MaxSysAnalysisCol = 10
End If

SysA.PermuteStepMode = PermuteModeNo

For sensitivitysloop = -1 To MaxSysAnalysisCol - 1
           
'Next sesitivitysloop
'Set System Analysis Variables

'SetSensitivityData_SI (indx As Long, ObjectStr As String, VarStr As String, Values_SI, vbActive As Boolean, QuantClass As String)
'SetSensitivityData (indx As Long, ObjectStr As String, VarStr As String, Values_DefEng, vbActive As Boolean)

'Sets the sensitivity information.
'Indx: the sensitivity box indx. Pass -1 for the X axis variable, 0,1,2, Åc , (MaxSensitivityVars Å]1) for the corresponding sensitivity.
'ObjectStr: the object name
'VarStr: the sensitivity variable name
'Values_DefEng: an array of double values in default engineering units containing the list of sensitivity values
'vbActive: sets the sensitivity group active (True) or inactive (False)

'SysA.SetSensitivityData
'syst.SetSensitivityData_SI sensitivitysloop, "System Data", "Liquid Rate", sensvalue, True, "STD VOLUME LIQUID RATE"

SingleBranchObj.SetOperationInterface SysA
SingleBranchObj.SetOperationInterface SysA


End Sub

Sub GetSysAnalysisResults(i As Integer)


End Sub


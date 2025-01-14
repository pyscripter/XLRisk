Attribute VB_Name = "UserInterface"
Option Explicit

Public UserStopped As Boolean
Public gSimulation As ClsSimulation

Public Function SetUpXLRisk() As Worksheet
'  Creates a sheet named XLRisk that contains risk settings, inputs and outputs
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim CurrentWS As Worksheet
     
    Set WB = ActiveWorkBook
    Set CurrentWS = WB.ActiveSheet
        
    On Error Resume Next
    Set WS = WB.Worksheets("XLRisk")
    If Err <> 0 Then
        Application.ScreenUpdating = False
        Err.Clear
        Set WS = WB.Sheets.Add(, WB.Worksheets(WB.Worksheets.Count))
        With WS
            .Name = "XLRisk"
            .Range("A1").Name.Visible = True
            .Cells(1, 1) = "Simulation Settings"
            .Cells(1, 1).Font.Bold = True
            .Cells(2, 1) = "Seed"
            .Cells(2, 2) = 0
            .Names.Add Name:="Seed", RefersTo:=.Cells(2, 2)
            .Cells(3, 1) = "Update Screen"
            .Cells(3, 2) = False
            .Names.Add Name:="ScreenUpdate", RefersTo:=.Cells(3, 2)
            .Cells(4, 1) = "Iterations"
            .Cells(4, 2) = 1000
            .Names.Add Name:="Iterations", RefersTo:=.Cells(4, 2)
            .Cells(5, 1) = "Latin Hypercube"
            .Cells(5, 2) = True
            .Names.Add Name:="LatinHypercube", RefersTo:=.Cells(5, 2)
            .Cells(6, 1) = "Calculate data tables during simulation"
            .Cells(6, 2) = True
            .Names.Add Name:="CalcDataTables", RefersTo:=.Cells(6, 2)
        
            .Cells(10, 1) = "Macro to run before each iteration"
            .Cells(10, 2) = ""
            .Names.Add Name:="MacroBefore", RefersTo:=.Cells(10, 2)
            .Cells(11, 1) = "Macro to run after each iteration"
            .Cells(11, 2) = ""
            .Names.Add Name:="MacroAfter", RefersTo:=.Cells(11, 2)
            .Cells(12, 1) = "Macro to run after simulation"
            .Cells(12, 2) = ""
            .Names.Add Name:="MacroAfterSimulation", RefersTo:=.Cells(12, 2)
                        
            .Columns(1).AutoFit
            .Range("A2..A12").Font.Italic = True
    
            .Cells(1, 4) = "Simulation Inputs"
            .Cells(3, 4) = "Range"
            .Cells(3, 5) = "Formula"
            .Range("D1..E3").Font.Bold = True
    
            .Names.Add Name:="RiskInputs", RefersTo:=.Cells(3, 4)
      
            .Cells(1, 7) = "Simulation Outputs"
            .Cells(3, 7) = "Range"
            .Cells(3, 8) = "Name"
            .Range("G1..H3").Font.Bold = True
            .Names.Add Name:="RiskOutputs", RefersTo:=.Cells(3, 7)
        End With
        CurrentWS.Activate
        Application.ScreenUpdating = True
    End If
    On Error GoTo 0
    ShowRiskInputs WS
    Set SetUpXLRisk = WS
End Function

Public Function CreateOutputSheet() As Worksheet
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim CurrentWS As Worksheet
    Dim I As Integer
     
    Set WB = ActiveWorkBook
    Set CurrentWS = WB.ActiveSheet
    Err.Clear
    
    On Error Resume Next
    I = 1
    Do While True
        Set WS = WB.Worksheets("Risk Results " + CStr(I))
        If Err <> 0 Then
            Err.Clear
            On Error GoTo 0
            Set WS = WB.Sheets.Add(, WB.Worksheets(WB.Worksheets.Count))
            WS.Name = "Risk Results " + CStr(I)
            Set CreateOutputSheet = WS
            CurrentWS.Activate
            Exit Do
        End If
        I = I + 1
    Loop
End Function

Public Sub ShowRiskInputs(XLRiskSheet As Worksheet)
    ' Show RiskInputs in the XLRisk sheet
    Dim R As Range
    Dim Coll As New Collection
    Dim RiskInput As ClsRiskInput
    
    Set R = XLRiskSheet.Range("RiskInputs").CurrentRegion
    ' Clear Inputs if present
    If R.Rows.Count > 1 Then R.Resize(R.Rows.Count - 1).Offset(1).Clear
    
    Set R = XLRiskSheet.Range("RiskInputs")
    CollectRiskInputs Coll
    For Each RiskInput In Coll
        Set R = R.Offset(1)
        R = AddressWithSheet(RiskInput.Cell)
        R.Offset(0, 1) = Right(RiskInput.Cell.Formula, Len(RiskInput.Cell.Formula) - 1)
    Next RiskInput
    R.CurrentRegion.Columns.AutoFit
End Sub

Public Sub ShowOptions()
' Action for related command button
    Load XLRiskOptions
    XLRiskOptions.Show
End Sub

Public Sub ShowAboutBox()
' Action for related command button
    Load AboutBox
    AboutBox.Show
End Sub

Public Sub StopSim()
' Action for related command button
    UserStopped = True
End Sub

Public Sub AddOutput()
' Action for the AddOutput command button
    Dim Name As String
    Dim XLRisk As Worksheet
    Dim R As Range
    Dim Sel As Range
    
    If Not TypeOf Selection Is Range Then Exit Sub
    Set Sel = Selection
    
    Name = InputBox("Please provide a name for the risk output", "Add Output")
    If Name = vbNullString Then Exit Sub
        
    Set XLRisk = SetUpXLRisk
    Set R = XLRisk.Range("RiskOutputs").CurrentRegion
    Set R = R.Rows(R.Rows.Count).Offset(1) 'Offset the last row
    R.Cells(1, 1).Formula = "=AddressWithSheet(" & AddressWithSheet(Sel) & ")"
    R.Cells(1, 2) = Name
    XLRisk.Range("RiskOutputs").CurrentRegion.Columns.AutoFit
End Sub

Public Sub ShowOutputs()
' Action for the AddOutput menu command
    Dim XLRisk As Worksheet
    Set XLRisk = SetUpXLRisk
    XLRisk.Activate
    XLRisk.Range("RiskOutputs").CurrentRegion.Select
End Sub



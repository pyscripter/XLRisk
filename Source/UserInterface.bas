Attribute VB_Name = "UserInterface"
Option Explicit


Public Function SetUpXLRisk() As Worksheet
'  Creates a sheet named XLRisk that contains risk settings, inputs and outputs
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim CurrentWS As Worksheet
     
    On Error Resume Next
    Set WB = ActiveWorkBook
    Set CurrentWS = WB.ActiveSheet
    Err.Clear
        
    Set WS = WB.Worksheets("XLRisk")
    If Err <> 0 Then
        Application.ScreenUpdating = False
        Err.Clear
        Set WS = WB.Sheets.Add(, WB.Worksheets(WB.Worksheets.Count))
        With WS
            .Name = "XLRisk"
            '.Visible = xlSheetVisible 'xlVeryHidden
            .Range("A1").Name = "XLRiskSetup"
            .Range("A1").Name.Visible = True
            .Cells(1, 1) = "Simulation Settings"
            .Cells(1, 1).Font.Bold = True
            .Cells(2, 1) = "Seed"
            .Cells(2, 2) = 0
            .Cells(2, 2).Name = "Seed"
            .Cells(3, 1) = "Update Screen"
            .Cells(3, 2) = False
            .Cells(3, 2).Name = "ScreenUpdate"
            .Cells(4, 1) = "Iterations"
            .Cells(4, 2) = 1000
            .Cells(4, 2).Name = "Iterations"
        
            .Range("A1").Columns.AutoFit
            .Range("A2.A4").Font.Italic = True
    
            .Cells(1, 4) = "Simulation Inputs"
            .Cells(3, 4) = "Range"
            .Cells(3, 5) = "Formula"
            .Range("G1.H3").Font.Bold = True
    
            WB.Names.Add Name:="RiskInputs", RefersTo:="=XLRisk!$D$3"
      
            .Cells(1, 7) = "Simulation Outputs"
            .Cells(3, 7) = "Range"
            .Cells(3, 8) = "Name"
            .Range("D1.E3").Font.Bold = True
            WB.Names.Add Name:="RiskOutputs", RefersTo:="=XLRisk!$G$3"
        End With
        CurrentWS.Activate
        Application.ScreenUpdating = True
    End If
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
            'Application.ScreenUpdating = False
            Set WS = WB.Sheets.Add(, WB.Worksheets(WB.Worksheets.Count))
            WS.Name = "Risk Results " + CStr(I)
            Set CreateOutputSheet = WS
            'Application.ScreenUpdating = True
            CurrentWS.Activate
            Exit Do
        End If
        I = I + 1
    Loop
End Function

Sub ShowRiskInputs(XLRiskSheet As Worksheet)
    ' Show RiskInputs in the XLRisk sheet
    Dim R As Range
    Dim Coll As New Collection
    Dim Cell As Range
    
    Set R = XLRiskSheet.Range("RiskInputs").CurrentRegion
    ' Clear Inputs if present
    If R.Rows.Count > 1 Then R.Resize(R.Rows.Count - 1).Offset(1).Clear
    
    Set R = XLRiskSheet.Range("RiskInputs")
    CollectRiskInputs Coll
    For Each Cell In Coll
        Set R = R.Offset(1)
        R = AddressWithSheet(Cell)
        R.Offset(0, 1) = Right(Cell.Formula, Len(Cell.Formula) - 1)
    Next Cell
    R.CurrentRegion.Columns.AutoFit
End Sub

Sub ShowOptions()
' Action for related command button
    Load XLRiskOptions
    XLRiskOptions.Show
End Sub

Sub ShowAboutBox()
' Action for related command button
    Load AboutBox
    AboutBox.Show
End Sub

Sub StopSim()
' Action for related command button
    UserStopped = True
End Sub

Sub AddOutput()
' Action for the AddOutput command button
    Dim Name As String
    Dim XLRisk As Worksheet
    Dim R As Range
    
    Name = InputBox("Please provide a name for the risk output", "Add Output")
    Set XLRisk = SetUpXLRisk
    Set R = XLRisk.Range("RiskOutputs").CurrentRegion
    Set R = R.Rows(R.Rows.Count).Offset(1) 'Offset the last row
    R.Cells(1, 1).Value = "'" & AddressWithSheet(Selection)
    R.Cells(1, 2) = Name
    XLRisk.Range("RiskOutputs").CurrentRegion.Columns.AutoFit
End Sub

Sub ShowOutputs()
' Action for the AddOutput menu command
    Dim XLRisk As Worksheet
    Set XLRisk = SetUpXLRisk
    XLRisk.Activate
    XLRisk.Range("RiskOutputs").CurrentRegion.Select
End Sub



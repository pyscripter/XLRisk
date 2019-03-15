Attribute VB_Name = "UserInterface"
Option Explicit
Const ShowSamplesTag = 1001
Const RunSimTag = 1002
Const StopSimTag = 1003
Const ShowOptionsTag = 1004
Const AddOutputTag = 1005
Const IterationsTag = 1006
Const MenuShowSamplesTag = 1101
Const MenuRunSimTag = 1102
Const MenuStopSimTag = 1103
Const MenuShowOptionsTag = 1104
Const MenuAddOutputTag = 1105
Const MenuIterationsTag = 1106


Sub CreateToolBar()
    ' Creates the XLRisk Command Bar
    Dim cbToolBar As CommandBar
    Dim cbbButton As CommandBarButton
    Dim cbbComboBox As CommandBarComboBox
    Dim Label As Variant
    
    On Error GoTo ErrorHandle
    
    ' In case it exists
    RemoveToolBar
    
    'Make the toolbar
    Set cbToolBar = CommandBars.Add
    cbToolBar.Name = "XLRisk"
    
    'Now we add a button to the toolbar. FaceId is the button's icon,
    'OnAction is the macro to run, if the button is clicked, and
    'ToolTipText is the text that will show when the mouse hovers.

    Set cbbButton = cbToolBar.Controls.Add(msoControlButton)
    With cbbButton
        .BeginGroup = True
        .Caption = "Setup"
        .Style = msoButtonIconAndCaption
        .OnAction = "ShowOptions"
        .FaceId = 642
        .Tag = ShowOptionsTag
        .TooltipText = "Setup XLRisk"
    End With
            
    Set cbbButton = cbToolBar.Controls.Add(msoControlButton)
    With cbbButton
        .Caption = "Show Samples"
        .Style = msoButtonIconAndCaption
        .OnAction = "ShowSamples"
        .FaceId = 98
        .Tag = ShowSamplesTag
        .TooltipText = "Show random samples of distributions instead of means"
        If ProduceRandomSample Then
            .State = msoButtonDown
        Else
            .State = msoButtonUp
        End If
    End With
    
    Set cbbButton = cbToolBar.Controls.Add(msoControlButton)
    With cbbButton
        .BeginGroup = True
        .Caption = "Add Output"
        .Style = msoButtonIconAndCaption
        .OnAction = "AddOutput"
        .FaceId = 6954
        .Tag = AddOutputTag
        .TooltipText = "Add selected cells to simulation outputs"
    End With
    
    #If Not Mac Then
    Set cbbComboBox = cbToolBar.Controls.Add(msoControlComboBox)
    With cbbComboBox
        .Caption = "Iterations:"
        .Style = msoComboLabel
        .AddItem (100)
        .AddItem (1000)
        .AddItem (10000)
        .Tag = IterationsTag
        .OnAction = "SetIterations"
       .TooltipText = "Select number of iterations"
    End With
    #End If
    
    Set cbbButton = cbToolBar.Controls.Add(msoControlButton)
    With cbbButton
        .Caption = "Run Simulation"
        .Style = msoButtonIconAndCaption
        .OnAction = "Simulate"
        .FaceId = 2151
        .Tag = RunSimTag
        .TooltipText = "Run Simulation with current settings"
    End With
    
    Set cbbButton = cbToolBar.Controls.Add(msoControlButton)
    With cbbButton
        .Caption = "Stop Simulation"
        .Style = msoButtonIconAndCaption
        .OnAction = "StopSim"
        .FaceId = 51
        .Tag = StopSimTag
        .TooltipText = "Interrupt/Stop simulation"
        '.Enabled = False
    End With
    
    Set cbbButton = cbToolBar.Controls.Add(msoControlButton)
    With cbbButton
        .BeginGroup = True
        .Caption = "About.."
        .Style = msoButtonIcon
        .OnAction = "ShowAboutBox"
        .FaceId = 9325
        .TooltipText = "Show info about XLRisk"
    End With
    
    cbToolBar.Visible = True
    
BeforeExit:
    Set cbToolBar = Nothing
    Set cbbButton = Nothing
    Exit Sub
ErrorHandle:
    MsgBox Err.Description & " CreateToolBar", vbOKOnly + vbCritical, "Error"
    RemoveToolBar
    Resume BeforeExit
End Sub

Sub RemoveToolBar()
    'Removes the toolbar "Shortcuts".
    'If it doesn't exist we get an error,
    'and that is why we use On Error Resume Next.
    
    On Error Resume Next
    CommandBars("XLRisk").Delete
End Sub

Public Sub CreateMenu()
    Dim RiskMenu As CommandBarPopup
    Dim ToolsMenu As CommandBarPopup
    
    Dim Ctrl As CommandBarButton
    
    ' In case it exists
    RemoveMenu
    
    On Error Resume Next
    
    Set RiskMenu = CommandBars.FindControl(, , "XLRiskMenu")
    RiskMenu.Delete
    
    Set ToolsMenu = CommandBars.FindControl(, 30007) ' "Tools Menu
    Set RiskMenu = ToolsMenu.Controls.Add(msoControlPopup, , , , True)
    RiskMenu.Caption = "XLRisk"
    RiskMenu.Tag = "XLRiskMenu"
    
    Set Ctrl = RiskMenu.CommandBar.Controls.Add(msoControlButton)
    With Ctrl
        .Caption = "Setup"
        .TooltipText = "Setup XLRisk"
        .Caption = "Setup"
        .Style = msoButtonIconAndCaption
        .OnAction = "ShowOptions"
        .FaceId = 642
        .Tag = MenuShowOptionsTag
    End With
    Set Ctrl = RiskMenu.CommandBar.Controls.Add(msoControlButton)
    With Ctrl
        .Caption = "Show Samples"
        .Style = msoButtonIconAndCaption
        .OnAction = "ShowSamples"
        .FaceId = 98
        .Tag = MenuShowSamplesTag
        .TooltipText = "Show random samples of distributions instead of means"
        If ProduceRandomSample Then
            .State = msoButtonDown
        Else
            .State = msoButtonUp
        End If
    End With
    Set Ctrl = RiskMenu.CommandBar.Controls.Add(msoControlButton)
    With Ctrl
        .BeginGroup = True
        .Caption = "Add Output"
        .Style = msoButtonIconAndCaption
        .OnAction = "AddOutput"
        .FaceId = 6954
        .Tag = MenuAddOutputTag
        .TooltipText = "Add selected cells to simulation outputs"
    End With
    Set Ctrl = RiskMenu.CommandBar.Controls.Add(msoControlButton)
    With Ctrl
        .Caption = "Output Ranges.."
        .TooltipText = "Show the Simulation Output ranges"
        .OnAction = "ShowOutputs"
    End With
    Set Ctrl = RiskMenu.CommandBar.Controls.Add(msoControlButton)
    With Ctrl
        .BeginGroup = True
        .Caption = "Simulate"
        .TooltipText = "Run Simulation"
        .OnAction = "Simulate"
        .Style = msoButtonIconAndCaption
        .FaceId = 2151
        .Tag = MenuRunSimTag
        .TooltipText = "Run Simulation with current settings"
    End With
    Set Ctrl = RiskMenu.CommandBar.Controls.Add(msoControlButton)
    With Ctrl
        .Caption = "Stop Simulation"
        .Style = msoButtonIconAndCaption
        .OnAction = "StopSim"
        .FaceId = 51
        .Tag = StopSimTag
        .TooltipText = "Interrupt/Stop simulation"
        '.Enabled = False
    End With
    Set Ctrl = RiskMenu.CommandBar.Controls.Add(msoControlButton)
    With Ctrl
        .BeginGroup = True
        .Caption = "About.."
        .Style = msoButtonIconAndCaption
        .OnAction = "ShowAboutBox"
        .FaceId = 9325
        .TooltipText = "Show info about XLRisk"
    End With
End Sub

Public Sub RemoveMenu()
  Dim RiskMenu As CommandBarPopup
  
  On Error Resume Next
  Set RiskMenu = CommandBars.FindControl(, , "XLRiskMenu")
  RiskMenu.Delete
End Sub


Public Function SetUpXLRisk() As Worksheet
'  Creates a sheet named XLRisk that contains risk settings, inputs and outputs
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim CurrentWS As Worksheet
     
    On Error Resume Next
    Set WB = ActiveWorkbook
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
     
    Set WB = ActiveWorkbook
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
        R = "'" & Cell.Parent.Name & "'!" & Cell.Address
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

Sub RandomSampleChanged()
    Dim Btn As CommandBarButton
        
    On Error GoTo Recalc
    Set Btn = CommandBars("XLRisk").FindControl(Tag:=ShowSamplesTag)
    If ProduceRandomSample Then
        Btn.State = msoButtonDown
    Else
        Btn.State = msoButtonUp
    End If
    Set Btn = CommandBars("XLRiskMenu").FindControl(Tag:=MenuShowSamplesTag)
    If ProduceRandomSample Then
        Btn.State = msoButtonDown
    Else
        Btn.State = msoButtonUp
    End If
Recalc:
    Application.Calculate
End Sub

Sub ShowSamples()
' Action for the ShowSamples command button
    ProduceRandomSample = Not ProduceRandomSample
    RandomSampleChanged
End Sub

Sub SetIterations()
' Action for the AddOutput command button
    Dim XLRisk As Worksheet
    Dim cbbComboBox As CommandBarComboBox
    
    Set cbbComboBox = CommandBars("XLRisk").FindControl(Tag:=IterationsTag)
    
    Set XLRisk = SetUpXLRisk
    On Error GoTo InvalidNumber
    XLRisk.Range("Iterations") = CInt(cbbComboBox.text)
    Exit Sub
InvalidNumber:
    MsgBox ("Invalid Number")
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
    R.Cells(1, 1).Value = "''" & Selection.Parent.Name & "'!" & Selection.Address
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


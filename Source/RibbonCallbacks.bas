Attribute VB_Name = "RibbonCallbacks"
' Declarations generated by Custom UI Editor
Option Explicit

Private XLRiskRibbonUI As IRibbonUI
Private Running As Boolean

#If Win32 Then
  'Used to restore RibbonUI ID
  Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As LongPtr)
#End If

'Callback for customUI.onLoad
Sub OnRibbonLoad(ribbon As IRibbonUI)
    Dim StoreRibbonPointer As LongPtr
    
    Set XLRiskRibbonUI = ribbon
    
    #If Win32 Then
        'Store pointer to IRibbonUI in a Named Range within add-in file
        StoreRibbonPointer = ObjPtr(ribbon)
        ThisWorkbook.Names.Add Name:="RibbonID", RefersTo:=StoreRibbonPointer
    #End If
End Sub

Sub GetRibbon()
    Dim objRibbon As Object
    Dim lRibbonPointer As LongPtr
    On Error GoTo ErrorHandler
    #If Win32 Then
        If XLRiskRibbonUI Is Nothing Then
            lRibbonPointer = CLngPtr(Replace(ThisWorkbook.Names("RibbonID").RefersTo, "=", ""))
            CopyMemory objRibbon, lRibbonPointer, LenB(lRibbonPointer)
            Set XLRiskRibbonUI = objRibbon
            Set objRibbon = Nothing
        End If
    #End If
ErrorHandler:
    Exit Sub
End Sub

'Callback for BtnSetup onAction
Sub RibbonShowOptions(control As IRibbonControl)
    ShowOptions
    If XLRiskRibbonUI Is Nothing Then GetRibbon
    If Not XLRiskRibbonUI Is Nothing Then
        XLRiskRibbonUI.InvalidateControl "ComboIterations"
        XLRiskRibbonUI.InvalidateControl "BtnSamples"
    End If
End Sub

'Callback for BtnSetup getEnabled
Sub RibbonGetEnabled(control As IRibbonControl, ByRef returnedVal)
    If control.ID = "BtnStop" Then
        returnedVal = Running
    Else
        returnedVal = Not Running
    End If
End Sub

'Callback for BtnSamples onAction
Sub RibbonShowSamples(control As IRibbonControl, pressed As Boolean)
    ProduceRandomSample = Not ProduceRandomSample
    Application.CalculateFull
End Sub

'Callback for BtnSamples getPressed
Sub RibbonShowSamplesGetPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = ProduceRandomSample
End Sub

'Callback for BtnHelp onAction
Sub RibbonHelp(control As IRibbonControl)
    ActiveWorkBook.FollowHyperlink "https://github.com/pyscripter/XLRisk/wiki"
End Sub

'Callback for BtnAddOutput onAction
Sub RibbonAddOutput(control As IRibbonControl)
    AddOutput
End Sub

'Callback for ComboIterations getText
Sub RibbonIterationsText(control As IRibbonControl, ByRef returnedVal)
    Dim XLRisk As Worksheet
     
    On Error Resume Next
    Set XLRisk = ActiveWorkBook.Worksheets("XLRisk")
    If Err = 0 Then
        returnedVal = CStr(XLRisk.Range("Iterations"))
    Else
        'XLRisk has not been setup
        returnedVal = 1000  'default
    End If
End Sub

'Callback for ComboIterations onChange
Sub RibbonSetIterations(control As IRibbonControl, text As String)
' Action for the AddOutput command button
    Dim XLRisk As Worksheet
    
    Set XLRisk = SetUpXLRisk
    On Error GoTo InvalidNumber
    XLRisk.Range("Iterations") = CLng(text)
    Exit Sub
InvalidNumber:
    MsgBox ("Invalid Number")
End Sub

' This is an auxialiary routine to force the update of the enabled state of the buttons
Sub DelayedSimulate()
    Dim Simulation As ClsSimulation
    
    On Error GoTo CleanUp
    Set Simulation = New ClsSimulation
    'Set the Global Simulation object
    Set gSimulation = Simulation
    Simulation.Run
CleanUp:
    Running = False
    If XLRiskRibbonUI Is Nothing Then GetRibbon
    If Not XLRiskRibbonUI Is Nothing Then XLRiskRibbonUI.Invalidate
    Set gSimulation = Nothing
End Sub

'Callback for BtnRun onAction
Sub RibbonSimulate(control As IRibbonControl)
    Running = True
    If XLRiskRibbonUI Is Nothing Then GetRibbon
    If Not XLRiskRibbonUI Is Nothing Then XLRiskRibbonUI.Invalidate
    'Simulate after a delay to give Excel time to update the Ribbon
    Application.OnTime Now + TimeSerial(0, 0, 1), "DelayedSimulate"
End Sub

'Callback for BtnStop onAction
Sub RibbonStopSim(control As IRibbonControl, pressed As Boolean)
    StopSim
End Sub

'Callback for BtnAbout onAction
Sub RibbonShowAboutBox(control As IRibbonControl)
  ShowAboutBox
End Sub



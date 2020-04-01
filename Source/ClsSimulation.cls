VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsSimulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mWB As Workbook
Private mSeed As Double
Private mIterations As Long
Private mScreenUpdate As Boolean
Private mSimError As Boolean
Private mSimErrorMsg As String
' The following two-dimensional array will contain the values of Iter, Risk Input Samples, Risk Output results
Private mSimResults() As Double

Property Get ActiveWorkBook() As Workbook
    Set ActiveWorkBook = mWB
End Property

Private Sub Class_Initialize()
    Dim XLRisk As Worksheet
    Set mWB = ActiveWorkBook
    
    Set XLRisk = SetUpXLRisk
    mSeed = XLRisk.Range("Seed").Value
    mIterations = XLRisk.Range("Iterations").Value
    mScreenUpdate = XLRisk.Range("ScreenUpdate").Value
End Sub

Private Sub SimIteration(Iter As Integer, RiskInputs As Collection, RiskOutputs As Collection)
    Dim Cell As Range
    Dim RiskOutput As ClsRiskOutput
    Dim I As Integer
      
    On Error GoTo SSError
        
    'Recalculate
    Application.Calculate
      
    ' Inputs
    mSimResults(Iter, 1) = Iter
    I = 2
    For Each Cell In RiskInputs
        mSimResults(Iter, I) = Cell.Value
        I = I + 1
    Next Cell
    
    'Outputs
    For Each RiskOutput In RiskOutputs
        mSimResults(Iter, I) = RiskOutput.Cell.Value
        I = I + 1
    Next RiskOutput
    Exit Sub
SSError:
    mSimError = True
    mSimErrorMsg = "Error in simulation iteration"
End Sub

Public Sub Run()
    Dim AppCalculation As XlCalculation
    Dim RiskInputs As Collection
    Dim RiskOutputs As Collection
    Dim Iter As Integer
    Dim OldProduceRandomSample As Boolean
    Dim SimResults As ClsSimResults
    
    ' Save ProduceRandomSample
    OldProduceRandomSample = ProduceRandomSample
    
    ' Save Calculation Mode
    AppCalculation = Application.Calculation
    Application.Calculation = xlCalculationManual
    
    On Error GoTo RestoreExcel
        
    '  Stop Screen Updating?
    Application.ScreenUpdating = mScreenUpdate
    Application.Cursor = xlWait
    
    Set RiskInputs = New Collection
    CollectRiskInputs RiskInputs
    If RiskInputs.Count = 0 Then
        MsgBox "No risk inputs defined", Title:="XLRisk simulation error"
        GoTo RestoreExcel
    End If
    
    If OneRiskFunctionPerCell(RiskInputs) Then GoTo RestoreExcel
    
    Set RiskOutputs = New Collection
    CollectRiskOutputs RiskOutputs
    If RiskOutputs.Count = 0 Then
        MsgBox "No risk outputs defined", Title:="XLRisk simulation error"
        GoTo RestoreExcel
    End If
        
    'Perform simulation
    UserStopped = False
    ProduceRandomSample = True
    Application.CalculateFull
        
    'Randomize
    If mSeed <> 0 Then
        'https://stackoverflow.com/questions/16589180/visual-basic-random-number-with-seed-but-in-reverse
        Rnd (-1)
        Randomize (mSeed)
    Else
        Randomize
    End If
    
    ' Allocate storage space
    ReDim mSimResults(1 To mIterations, 1 To 1 + RiskInputs.Count + RiskOutputs.Count)
    
    For Iter = 1 To mIterations
        If mSimError Then
            mSimError = False
            MsgBox mSimErrorMsg
            Exit For
        End If
        SimIteration Iter, RiskInputs, RiskOutputs
        DoEvents
        'Check whether to Stop
        If UserStopped Then
            UserStopped = False
            MsgBox "The simulation was interrupted"
            Exit For
        End If
        Application.StatusBar = "Iteration: " & CStr(Iter) & "/" & CStr(mIterations)
    Next Iter
    
    Set SimResults = New ClsSimResults
    SimResults.Init RiskInputs, RiskOutputs, mIterations, mSimResults
    SimResults.ProduceResults
RestoreExcel:
    'Restore ProduceRandomSample
    ProduceRandomSample = OldProduceRandomSample
    
    'Restore Calculation Mode
    Application.Calculation = AppCalculation
    Application.CalculateFull
    
    ' Restore Status Bar
    Application.StatusBar = False
    
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
End Sub

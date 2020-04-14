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
Private mProcessingCorrmatInfo
Private mSeed As Double
Private mIterations As Long
Private mScreenUpdate As Boolean
Private mLatinHypercube As Boolean
Private mSimError As Boolean
Private mSimErrorMsg As String
' The following two-dimensional array will contain the values of Iter, Risk Input Samples, Risk Output results
Private mSimResults() As Variant
Private mRiskInputs As Collection
Private mRiskOutputs As Collection
Private mInputGroups As Collection
Private mCurrentIteration As Long

Private Const NoCorrelationGroupName As String = "No Correlation"

Property Get ActiveWorkBook() As Workbook
    Set ActiveWorkBook = mWB
End Property

Property Get ProcessingCorrmatInfo() As Boolean
    ProcessingCorrmatInfo = mProcessingCorrmatInfo
End Property

Private Sub Class_Initialize()
    Dim XLRisk As Worksheet
    Set mWB = Application.ActiveWorkBook
    
    Set XLRisk = SetUpXLRisk
    mSeed = XLRisk.Range("Seed").Value
    mIterations = XLRisk.Range("Iterations").Value
    mScreenUpdate = XLRisk.Range("ScreenUpdate").Value
    mLatinHypercube = XLRisk.Range("LatinHypercube").Value
    mCurrentIteration = 0
    mProcessingCorrmatInfo = False
End Sub

Private Sub SimIteration(Iter As Integer)
    Dim RiskInput As ClsRiskInput
    Dim RiskOutput As ClsRiskOutput
    Dim I As Integer
      
    On Error GoTo SSError
        
    'Recalculate
    Application.Calculate
      
    ' Inputs
    mSimResults(Iter, 1) = Iter
    I = 2
    For Each RiskInput In mRiskInputs
        mSimResults(Iter, I) = RiskInput.Cell.Value2
        I = I + 1
    Next RiskInput
    
    'Outputs
    For Each RiskOutput In mRiskOutputs
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
    Dim Iter As Integer
    Dim OldProduceRandomSample As Boolean
    Dim SimResults As ClsSimResults
    Dim InputGroup As ClsInputGroup
    Dim RiskInput As ClsRiskInput
    
    ' Save ProduceRandomSample
    OldProduceRandomSample = ProduceRandomSample
    
    ' Save Calculation Mode
    AppCalculation = Application.Calculation
    Application.Calculation = xlCalculationManual
    
    On Error GoTo RestoreExcel
        
    '  Stop Screen Updating?
    Application.ScreenUpdating = mScreenUpdate
    Application.Cursor = xlWait
    
    Set mRiskInputs = New Collection
    CollectRiskInputs mRiskInputs
    If mRiskInputs.Count = 0 Then
        MsgBox "No risk inputs defined", Title:="XLRisk simulation error"
        GoTo RestoreExcel
    End If
    
    If Not OneRiskFunctionPerCell(mRiskInputs) Then GoTo RestoreExcel
    
    Set mRiskOutputs = New Collection
    CollectRiskOutputs mRiskOutputs
    If mRiskOutputs.Count = 0 Then
        MsgBox "No risk outputs defined", Title:="XLRisk simulation error"
        GoTo RestoreExcel
    End If
        
    'Perform simulation
    UserStopped = False
    ProduceRandomSample = True
    
    'Setup Risk Input Groups
    Set mInputGroups = New Collection
    
    ' Do a full calculation and process correlation information
    mProcessingCorrmatInfo = True
    Application.CalculateFull
    mProcessingCorrmatInfo = False
        
    ' Create an InputGroup with uncorrelated inputs
    For Each RiskInput In mRiskInputs
        If RiskInput.CorrmatAddress = vbNullString Then
            If InputGroup Is Nothing Then
                Set InputGroup = New ClsInputGroup
                InputGroup.Init mIterations, mLatinHypercube
                mInputGroups.Add InputGroup, NoCorrelationGroupName
            End If
            RiskInput.CorrmatAddress = NoCorrelationGroupName
            RiskInput.Index = InputGroup.AddRiskInput(RiskInput)
        End If
    Next RiskInput
    
    'Check Input Groups
    For Each InputGroup In mInputGroups
        If Not InputGroup.IsValid Then GoTo RestoreExcel
        InputGroup.GenerateRndSamples
    Next InputGroup
    
    'Randomize
    If mSeed <> 0 Then
        'https://stackoverflow.com/questions/16589180/visual-basic-random-number-with-seed-but-in-reverse
        Rnd (-1)
        Randomize (mSeed)
    Else
        Randomize
    End If
    
    ' Allocate storage space
    ReDim mSimResults(1 To mIterations, 1 To 1 + mRiskInputs.Count + mRiskOutputs.Count)
    ' Set up Risk InputGroups
      
    For Iter = 1 To mIterations
        mCurrentIteration = Iter
        If mSimError Then
            mSimError = False
            MsgBox mSimErrorMsg
            Exit For
        End If
        SimIteration Iter
        DoEvents
        'Check whether to Stop
        If UserStopped Then
            UserStopped = False
            MsgBox "The simulation was interrupted"
            GoTo RestoreExcel
        End If
        Application.StatusBar = "Iteration: " & CStr(Iter) & "/" & CStr(mIterations)
    Next Iter
    mCurrentIteration = 0
    
    Set SimResults = New ClsSimResults
    SimResults.Init mRiskInputs, mRiskOutputs, mIterations, mSimResults
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

Public Function GetRndSample(Cell As Range) As Double
    Dim RiskInput As ClsRiskInput
    Dim InputGroup As ClsInputGroup
    If (mCurrentIteration < 1) Or Not (Cell.Worksheet.Parent Is mWB) Then
        GetRndSample = Rnd()
    Else
        Set RiskInput = mRiskInputs.Item(AddressWithSheet(Cell))
        Set InputGroup = mInputGroups.Item(RiskInput.CorrmatAddress)
        GetRndSample = InputGroup.GetRndSample(mCurrentIteration, RiskInput.Index)
    End If
End Function

Public Sub ProcessCorrmatInfo(Cell As Range, CorrmatRng As Range, Index As Long)
    Dim InputGroup As ClsInputGroup
    Dim RiskInput As ClsRiskInput
    Dim Address As String
    
    Address = AddressWithSheet(CorrmatRng)
    On Error GoTo InputNotFound
    Set RiskInput = mRiskInputs.Item(AddressWithSheet(Cell))
    On Error GoTo 0
    RiskInput.CorrmatAddress = Address
    RiskInput.Index = Index

    On Error GoTo InputGroupNotFound
    Set InputGroup = mInputGroups.Item(Address)
    On Error GoTo 0
    InputGroup.AddRiskInput RiskInput
    Exit Sub
InputGroupNotFound:
    On Error GoTo 0
    Set InputGroup = New ClsInputGroup
    InputGroup.Init mIterations, mLatinHypercube
    Set InputGroup.CorrmatRng = CorrmatRng
    mInputGroups.Add InputGroup, Address
    InputGroup.AddRiskInput RiskInput
    Exit Sub
InputNotFound:
    Err.Raise Number:=vbObjectError + 15, _
              Source:="ProcessCorrmatInfo", _
              Description:="Risk Input Not Found"
End Sub



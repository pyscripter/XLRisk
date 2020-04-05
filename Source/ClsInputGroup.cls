VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsInputGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'   In XLRISK input variables are grouped based on the specified correlation matrices.
'   For every matrix specified an InputGroup is created.  One extra uncorrelated group
'   is created if there are risk inputs not linked to a correlation matrix.
'   This class takes care of calculating random samples from a (0, 1] distribution that can
'   be used by the risk functions of the risk inputs as an input to the inversion sampling
'   method.
'   Dependending on the UseLatinHypercube option the generated samples can use the Latin
'   Hypercube sampling method which accelerates convergence.
'   The Iman-Conover method is used for calculating correlated inputs.

Private mRndSamples() As Double
Private mIterations As Long
Private mRiskInputs As Collection
Private mLatinHypercube As Boolean
'  A Collection object that is used as a Dictionary of indices
Private mCorrmatRng As Range

Property Set CorrmatRng(ARange As Range)
    Set mCorrmatRng = ARange
End Property

Public Sub Init(Iterations As Long, LatinHyperCube As Boolean)
    mIterations = Iterations
    mLatinHypercube = LatinHyperCube
    Set mRiskInputs = New Collection
End Sub

Function AddRiskInput(RiskInput As ClsRiskInput) As Long
    mRiskInputs.Add RiskInput, AddressWithSheet(RiskInput.Cell)
    AddRiskInput = mRiskInputs.Count
End Function

Public Sub GenerateRndSamples()
    Dim Iter As Long
    Dim Col As Long
    Dim U As Double
    Dim NVars As Long
    Dim I As Long
    Dim J As Long
    Dim Corrmat() As Double
    
    
    ReDim mRndSamples(1 To mIterations, 1 To mRiskInputs.Count)
    For Iter = 1 To mIterations
        For Col = 1 To mRiskInputs.Count
            '  We use Latin Hypercube when the inputs are correlated
            If mLatinHypercube Or Not (mCorrmatRng Is Nothing) Then
                If Iter = 1 Then
                    Do
                        U = Rnd()
                    Loop Until U <> 0
                    mRndSamples(Iter, Col) = U / mIterations
                Else
                    mRndSamples(Iter, Col) = (Rnd() + Iter - 1) / mIterations
                End If
            Else
                mRndSamples(Iter, Col) = Rnd()
            End If
        Next Col
    Next Iter
    If (mCorrmatRng Is Nothing) And mLatinHypercube Then
        For Col = 1 To mRiskInputs.Count
            Shuffle mRndSamples, Col
        Next Col
    ElseIf Not (mCorrmatRng Is Nothing) Then
        '   Complete the correlation matrix and apply the Iman Conover Method
        NVars = mCorrmatRng.Rows.Count
        ReDim Corrmat(1 To NVars, 1 To NVars)
        
        ' Fill the array with values from the range.
        For I = 1 To NVars
            For J = 1 To NVars
                Corrmat(I, J) = mCorrmatRng(I, J).Value
            Next
        Next
        
        ' Fill in the correlation array if it's missing the lower or upper part.
        For I = 1 To NVars
            For J = 1 To NVars
                If I <> J Then
                    If Corrmat(I, J) = 0 Then
                        Corrmat(I, J) = Corrmat(J, I)
                    ElseIf Corrmat(J, I) = 0 Then
                        Corrmat(J, I) = Corrmat(I, J)
                    End If
                End If
            Next
        Next
        mRndSamples = ImanConover(mRndSamples, Corrmat)
    End If
End Sub

Function IsValid() As Boolean
    Dim Size As Long
    Dim IsSet() As Boolean
    Dim RiskInput As ClsRiskInput

    IsValid = True
    If mCorrmatRng Is Nothing Then Exit Function
        
    IsValid = RiskIsValidCorrmat(mCorrmatRng, , True)
    If Not IsValid Then Exit Function
    
    Size = mCorrmatRng.Rows.Count
    If mRiskInputs.Count <> Size Then
        IsValid = False
        MsgBox "The risk input functions have not been properly linked to the" _
            & "correlation matrix at " & AddressWithSheet(mCorrmatRng), vbExclamation
        Exit Function
    End If
    
    ReDim IsSet(1 To Size)
    For Each RiskInput In mRiskInputs
        If IsSet(RiskInput.Index) Then
            IsValid = False
            MsgBox "The risk input functions have not been properly linked to the" _
                & "correlation matrix at " & AddressWithSheet(mCorrmatRng), vbExclamation
            Exit Function
        End If
        IsSet(RiskInput.Index) = True
    Next RiskInput
End Function

Public Function GetRndSample(Iter As Long, Index As Long) As Double
    GetRndSample = mRndSamples(Iter, Index)
End Function


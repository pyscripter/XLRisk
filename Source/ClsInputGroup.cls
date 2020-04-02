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

Private mIterations As Long
Private mRiskInputs As Collection
Private mLatinHypercube As Boolean
'  A Collection object that is used as a Dictionary of indices
Private mIndices As Collection
Private mRndSamples() As Double

Public Sub Init(Iterations As Long, RiskInputs As Collection, LatinHyperCube As Boolean)
    Dim Cell As Range
    Dim Index As Variant
    mIterations = Iterations
    mLatinHypercube = LatinHyperCube
    Set mRiskInputs = RiskInputs
    Set mIndices = New Collection
    Index = 0
    For Each Cell In RiskInputs
        Index = Index + 1
        mIndices.Add Index, AddressWithSheet(Cell)
    Next Cell
    GenerateRndSamples
End Sub

Private Sub GenerateRndSamples()
    Dim Iter As Long
    Dim Col As Long
    Dim U As Double
    ReDim mRndSamples(1 To mIterations, 1 To mIndices.Count)
    For Iter = 1 To mIterations
        For Col = 1 To mRiskInputs.Count
            If mLatinHypercube Then
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
    If mLatinHypercube Then
        For Col = 1 To mRiskInputs.Count
            Shuffle mRndSamples, Col
        Next Col
    End If
End Sub

Public Function GetRndSample(Iter As Long, Address As String) As Double
    Dim Index As Long
    Index = mIndices.Item(Address)
    GetRndSample = mRndSamples(Iter, Index)
End Function

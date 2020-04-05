VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsRiskInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCell As Range
Private mCorrmatAddress As String
Private mIndex As Long

Property Get Cell() As Range
    Set Cell = mCell
End Property

Sub Init(Cell As Range)
    Set mCell = Cell
End Sub

Property Get CorrmatAddress() As String
    CorrmatAddress = mCorrmatAddress
End Property

Property Let CorrmatAddress(Address As String)
    mCorrmatAddress = Address
End Property

Property Get Index() As Long
    Index = mIndex
End Property

Property Let Index(AIndex As Long)
    mIndex = AIndex
End Property


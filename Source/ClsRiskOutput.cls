VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsRiskOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mName As String
Private mCell As Range

Property Get Name() As String
    Name = mName
End Property

Property Get Cell() As Range
    Set Cell = mCell
End Property

Sub Init(AName As String, ACell As Range)
    mName = AName
    Set mCell = ACell
End Sub


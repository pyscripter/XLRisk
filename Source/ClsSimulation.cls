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

Public Sub Run()

End Sub

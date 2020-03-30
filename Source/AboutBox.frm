VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AboutBox 
   Caption         =   "About XLRisk"
   ClientHeight    =   2730
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   4350
   OleObjectBlob   =   "AboutBox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AboutBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const XLRiskVersion = "0.80"

Private Sub btnOK_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    lblVersionNumber.Caption = XLRiskVersion
End Sub

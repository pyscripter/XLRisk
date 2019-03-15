VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_AddinInstall()
#If Mac Then
    CreateToolBar
#End If
  FunctionDescriptions
End Sub

Private Sub Workbook_AddinUninstall()
#If Mac Then
    RemoveToolBar
#End If
End Sub


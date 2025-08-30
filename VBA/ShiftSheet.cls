VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShiftSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub RefreshShiftButton_Click()
  Call MacroFunc.ShiftOpRefresh
End Sub

Private Sub PrintButton_Click()
  Call WebFunc.UploadTextViaFTP("10.68.89.233", "sato", "pass", "label.ipl", ShiftSheet.Range("PrintLabelCell").Text)
End Sub

Private Sub RefreshSQLButton_Click()
  Call MacroFunc.SQLRefresh
End Sub

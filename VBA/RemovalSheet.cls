VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RemovalSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub ClearButton_Click()
  Call MacroFunc.ClearResponseCells
End Sub

Private Sub RefreshButton_Click()
  Call MacroFunc.SQLRefresh
End Sub

Private Sub SQLSendButton_Click()
  Call SendSQLCommand("aimwrnmes01", "MES_COMMON", Range("SQLQueryCell").Text, "SQLResponseCell")
End Sub

Private Sub SQLSendButtonDT_Click()
  Call SendSQLCommand("aimwrnnyssql2", "MES_COMMON", Range("SQLQueryCell_DT").Text, "SQLResponseCell_DT")
End Sub

Private Sub SQLSendButtonWS_Click()
  Call SendSQLCommand("aimwrnnyssql", "MES_COMMON", Range("SQLQueryCell_WS").Text, "SQLResponseCell_WS")
End Sub

Private Sub StaffRemovalButton_Click()
  Dim r As Range
  For Each r In RemovalSheet.Range("RemoveOpTable[Name]").Cells
    Call MacroFunc.DeleteStaffRow(r.Value)
  Next r
End Sub

Private Sub VBAButton_Click()
  Call MacroFunc.OpenVBAToProc("RemovalSheet", "SQLSendButton_Click")
End Sub

Private Sub Worksheet_Calculate()
  SQLSendButton.Enabled = StartsWith(Range("SQLQueryCell").Text, "update")
  SQLSendButtonDT.Enabled = StartsWith(Range("SQLQueryCell_DT").Text, "update")
  SQLSendButtonWS.Enabled = StartsWith(Range("SQLQueryCell_WS").Text, "update")
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
  Call Worksheet_Calculate
End Sub

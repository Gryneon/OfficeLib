Attribute VB_Name = "MacroFunc"
Option Explicit
Option Base 1

'`Macro Function Library
'Version 1.0.1

'History
' 1.0.1 - Added ErrKey, ActiveErr, ClearFormat, RefreshTables

'Current

Public Sub RefreshQueryTables()
  Dim Sheet As Worksheet, QTable As QueryTable
  For Each Sheet In ThisWorkbook.Sheets
    For Each QTable In Sheet.QueryTables
      QTable.Refresh
    Next
  Next
End Sub

Sub ErrKey()
  Selection.Style = "ErrKey"
End Sub

Sub ActiveErr()
  Selection.Style = "ActError"
End Sub

Sub ClearFormat()
  Selection.Style = "Normal"
End Sub

Sub RefreshTables()

  Dim Sheet As Worksheet
  Dim Query As QueryTable
  
  For Each Sheet In ActiveWorkbook.Worksheets
  
    For Each Query In Sheet.QueryTables
      Query.Refresh BackgroundQuery:=False
    Next Query
    
  Next Sheet
  
End Sub

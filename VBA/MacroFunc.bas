Attribute VB_Name = "MacroFunc"
Option Explicit
Option Base 1

'`Macro Function Library
'Version 1.0.0

'Current

Public Sub RefreshQueryTables()
  Dim Sheet As Worksheet, QTable As QueryTable
  For Each Sheet In ThisWorkbook.Sheets
    For Each QTable In Sheet.QueryTables
      QTable.Refresh
    Next
  Next
End Sub


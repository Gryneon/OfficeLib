Attribute VB_Name = "LkpFunc"
Option Explicit
Option Compare Text
Option Base 1

'`Lookup Function Library
'Version 1.1.3

'History
' 1.1.2 - Added XLCellLookup function
' 1.1.3 - Added XLTableRow, XLFindRow, XLFindTableRow function
'Current

Public Function XLIntersect(ByVal R1, ByVal R2)
  Set XLIntersect = Intersect(R1, R2)
End Function

Public Function XLEntireRow(ByVal cell)
  Set XLEntireRow = cell.EntireRow
End Function

Public Function XLCellLookup(ByVal RowCell, ByVal ColCell)
  Set XLCellLookup = Intersect(RowCell.EntireRow, ColCell.EntireColumn)
End Function

Public Function XLEntireColumn(ByVal cell As Range) As Range
  Set XLEntireColumn = cell.EntireColumn
End Function

Public Function XLTableRow(ByVal cell As Range, ByVal table As String) As Range
  Set XLTableRow = Intersect(cell, Range(table))
End Function

Public Function XLFindRow(ByVal findText As String, ByVal col As Range) As Range
  Set XLFindRow = col.Find(findText, LookIn:=xlValues, MatchCase:=True)
End Function

Public Function XLFindTableRow(findText As String, col As Range, Optional table As ListObject = Nothing) As Range
  If table Is Nothing Then Set table = col.ListObject
  Set XLFindTableRow = Intersect(XLFindRow(findText, col), table.DataBodyRange)
End Function
 
Public Function GLookup(ByRef table, ByVal RVal, ByVal row, ByVal CVal, ByVal col)
  On Error GoTo Invalid

  Dim r As Range: Set r = r.Find(RVal, LookIn:=row).EntireColumn
  Dim C As Range: Set C = C.Find(CVal, LookIn:=col).EntireRow
  
  Set GLookup = Intersect(r, C)
 
Exit Function
Invalid:
  ErrorMsg
End Function

Public Function RowNum(ByRef HeaderCell)
  Application.Volatile
  Dim ThisCell As Range: Set ThisCell = Application.Caller
  RowNum = ThisCell.row - HeaderCell.row
End Function

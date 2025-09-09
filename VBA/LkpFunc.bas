Attribute VB_Name = "LkpFunc"
Option Explicit
Option Compare Text
Option Base 1

'`Lookup Function Library
'Version 1.1.4

'History
' 1.1.2 - Added XLCellLookup function
' 1.1.3 - Added XLTableRow, XLFindRow, XLFindTableRow function
' 1.1.4 - Updated GLookup
'Current

Public Function XLIntersect(R1, R2)
  Set XLIntersect = Intersect(R1, R2)
End Function

Public Function XLEntireRow(cell)
  Set XLEntireRow = cell.EntireRow
End Function

Public Function XLCellLookup(RowCell, ColCell)
  Set XLCellLookup = Intersect(RowCell.EntireRow, ColCell.EntireColumn)
End Function

Public Function XLEntireColumn(cell As Range) As Range
  Set XLEntireColumn = cell.EntireColumn
End Function

Public Function XLTableRow(cell As Range, ByVal table As String) As Range
  Set XLTableRow = Intersect(cell, Range(table))
End Function

Public Function XLFindRow(ByVal findText As String, col As Range) As Range
  Set XLFindRow = col.Find(findText, LookIn:=xlValues, MatchCase:=True)
End Function

Public Function XLFindTableRow(findText As String, col As Range, Optional table As ListObject = Nothing) As Range
  If table Is Nothing Then Set table = col.ListObject
  Set XLFindTableRow = Intersect(XLFindRow(findText, col), table.DataBodyRange)
End Function
 
Public Function GLookup(ByVal RVal As Variant, KeyColumn As Range, ByVal CVal As Variant, ByVal HeaderRow As Range) As Variant
  On Error GoTo NotFound

  Dim RCell As Range: Set RCell = KeyColumn.Find(RVal, LookIn:=xlValues, LookAt:=xlWhole)
  Dim CCell As Range: Set CCell = HeaderRow.Find(CVal, LookIn:=xlValues, LookAt:=xlWhole)
  
  If RCell Is Nothing Or CCell Is Nothing Then GoTo NotFound
  
  Dim R As Range: Set R = RCell.EntireRow
  Dim C As Range: Set C = CCell.EntireColumn
  
  GLookup = Intersect(R, C).Value
 
Exit Function
NotFound:
  GLookup = CVErr(xlErrNA)
End Function

Public Function RowNum(ByRef HeaderCell)
  Application.Volatile
  Dim ThisCell As Range: Set ThisCell = Application.Caller
  RowNum = ThisCell.row - HeaderCell.row
End Function

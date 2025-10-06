Attribute VB_Name = "LkpFunc"
Option Explicit
Option Compare Text
Option Base 1

'`Lookup Function Library
'Version 1.1.5

'History
' 1.1.2 - Added XLCellLookup function
' 1.1.3 - Added XLTableRow, XLFindRow, XLFindTableRow function
' 1.1.4 - Added GetValueOf function
` 1.1.5 - Updated GLookup
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

Public Function XLEntireColumn(ByVal cell As range) As range
  Set XLEntireColumn = cell.EntireColumn
End Function

Public Function XLTableRow(ByVal cell As range, ByVal table As String) As range
  Set XLTableRow = Intersect(cell, range(table))
End Function

Public Function XLFindRow(ByVal findText As String, ByVal col As range) As range
  Set XLFindRow = col.Find(findText, LookIn:=xlValues, MatchCase:=True)
End Function

Public Function XLFindTableRow(findText As String, col As range, Optional table As ListObject = Nothing) As range
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

Public Function GetValueOf(ByVal name As String)
  Dim rng As range
  Set rng = range(name)
  GetValueOf = rng.Value
End Function

Public Function RowNum(ByRef HeaderCell)
  Application.Volatile
  Dim ThisCell As range: Set ThisCell = Application.Caller
  RowNum = ThisCell.row - HeaderCell.row
End Function

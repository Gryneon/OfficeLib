Attribute VB_Name = "LkpFunc"
Option Explicit
Option Compare Text
Option Base 1

'`Lookup Function Library
'Version 1.1.2

'History
' 1.1.2 - Added XLCellLookup function

'Current

Public Function XLIntersect(ByVal R1, ByVal R2)
  XLIntersect = Intersect(R1, R2)
End Function

Public Function XLEntireRow(ByVal cell)
  XLEntireRow = cell.EntireRow
End Function

Public Function XLCellLookup(ByVal RowCell, ByVal ColCell)
  XLCellLookup = Intersect(RowCell.EntireRow, ColCell.EntireColumn)
End Function

Public Function XLEntireColumn(ByVal cell)
  XLEntireColumn = cell.EntireColumn
End Function

Public Function GLookup(ByRef Table, ByVal RVal, ByVal Row, ByVal CVal, ByVal Col)
  On Error GoTo Invalid

  Dim r As Range: Set r = r.Find(RVal, LookIn:=Row).EntireColumn
  Dim C As Range: Set C = C.Find(CVal, LookIn:=Col).EntireRow
  
  GLookup = Intersect(r, C)
 
Exit Function
Invalid:
  ErrorMsg
End Function

Public Function RowNum(ByRef HeaderCell)
  Application.Volatile
  Dim ThisCell As Range: Set ThisCell = Application.Caller
  RowNum = ThisCell.Row - HeaderCell.Row
End Function

Attribute VB_Name = "MacroFunc"
Option Explicit

Public Sub SQLRefresh()
  Call SQLSheet.ListObjects("SQLOperatorTable").Refresh
  Call DTSheet.ListObjects("SQLOperators_DT").Refresh
  Call WSSheet.ListObjects("SQLOperators_WS").Refresh
  Call IDPoolSheet.ListObjects("AllowedIDs").Refresh
End Sub

Public Sub AddOperators()
  Call IDPoolSheet.ListObjects("AllowedIDs").Refresh
  
  Dim col As Range
  Set col = Range("CalcAddOpTable").Cells
  
  Dim Index As Range
  For Each Index In col
    If StartsWith(Index.Text, "*") Then
      Dim Dept As Range
      Dim DTex As String
      Set Dept = Index.Offset(0, -11)
      DTex = Dept.Text
      'lookup table name in table table
      'add row to table with text in DTex
      'refresh table to add line
      'copy cells to clipboard
    End If
  Next Index
  
  
End Sub

Public Sub ShiftOpRefresh()
  Range("OperatorsByDept").ListObject.Refresh
End Sub

Public Sub FilterOps()
  
  Dim Active As Boolean, Op As String, UseActive As Boolean, UseDept As Boolean
  Active = Range("FilterActive")
  UseActive = Len(Range("FilterActive").Text) > 0
  UseDept = Len(Range("FilterDeptCode")) > 0
  
  Op = IIf(Active, "=", "<>")
  With Range("SQLOperatorTable")
  
    If Not UseActive And Not UseDept Then
      .AutoFilter field:=3, Criteria1:="<>"
      .AutoFilter field:=4, Operator:=xlOr, Criteria1:="=", Criteria2:="<>"
    ElseIf Not UseActive And UseDept Then
      .AutoFilter field:=3, Criteria1:=Range("FilterDeptCode")
      .AutoFilter field:=4, Operator:=xlOr, Criteria1:="=", Criteria2:="<>"
    ElseIf UseActive And Not UseDept Then
      .AutoFilter field:=3, Criteria1:="<>"
      .AutoFilter field:=4, Criteria1:=Op
    Else
      .AutoFilter field:=3, Criteria1:=Range("FilterDeptCode")
      .AutoFilter field:=4, Criteria1:=Op
    End If
  
  End With
  
End Sub

Sub DeleteStaffRow(ByVal Operator As String)
  Dim table As ListObject
  Dim col As Range
  Dim match As Range
  
  Set table = ShiftSheet.ListObjects("EmployeesByDept")
  Set col = ShiftSheet.Range("EmployeesByDept[Name]")
  Set match = col.Find(What:=Operator, LookAt:=xlWhole, MatchCase:=False)
  
  If Not match Is Nothing Then
    Call Intersect(match.EntireRow, table.DataBodyRange).Delete(xlShiftUp)
  End If
  
End Sub

Sub ClearResponseCells()
  RemovalSheet.Range("SQLResponseCell").Value = ""
  RemovalSheet.Range("SQLResponseCell_DT").Value = ""
  RemovalSheet.Range("SQLResponseCell_WS").Value = ""
End Sub

Sub OpenVBAToProc(module As String, proc As String)
  Dim mdl As CodeModule
  Dim lineNum As Long
  
  Set mdl = ThisWorkbook.VBProject.VBComponents(module).CodeModule
  lineNum = mdl.ProcStartLine(proc, vbext_pk_Proc)
  
  With Application.VBE
    .MainWindow.Visible = True
    Call .ActiveVBProject.VBComponents(module).Activate
    Call mdl.CodePane.SetSelection(lineNum, 1, lineNum, 1)
    Call .MainWindow.SetFocus
  End With
  
End Sub

Attribute VB_Name = "StyleFunc"
Option Explicit
Option Compare Text
Option Base 1

'`Style Function Library
'Version 1.0.8

'History
' 1.0.8 - Fixed Selection.Range Error
' 1.0.7 - Added SelectColumn

'Current

Private Const DefStyle = "Normal"
Private Const DefTitleText = "Added Title"
Private ActiveFontSet As FontSet

Public Type FontSet
  Body As String
  Mono As String
  Cond As String
  Head As String
  BSize As Integer
  HSize As Integer
  TSize As Integer
  SetFont As Boolean
  SetForm As Boolean
  ChgDef As Boolean
End Type

Public Sub SelectColumn()

  Dim Ref As Range: Set Ref = Selection.Range
  Set Ref = Intersect(Ref.ListObject.DataBodyRange, Ref.EntireColumn)
  Ref.Select

End Sub

Private Sub TableColumn(typ As String, Ref As Range, Optional Body As String = "Cell", Optional Head As String = "Hd")
  Dim Bd: Set Bd = Intersect(Ref.ListObject.DataBodyRange, Ref.EntireColumn)
  Dim Hd: Set Hd = Intersect(Ref.ListObject.HeaderRowRange, Ref.EntireColumn)
  Bd.style = typ + Body
  Hd.style = typ + Head
End Sub

Public Sub LookupColumn()
  TableColumn "Lkp", Selection
End Sub

Public Sub CalcColumn()
  TableColumn "Calc", Selection
End Sub

Public Sub DeacColumn()
  TableColumn "Deac", Selection
End Sub

Public Sub InputColumn()
  TableColumn "Inp", Selection
End Sub

Public Sub InternalColumn()
  TableColumn "Int", Selection
End Sub

Public Sub ErrorColumn()
  TableColumn "Err", Selection
End Sub

Public Sub QueryColumn()
  TableColumn "Que", Selection
End Sub

Public Sub FixColumn()
  Dim Sel As style, STyp As String, HTyp As String, BTyp As String
  Set Sel = Selection.style
  
  If Sel Like "Act*" Or Sel = DefStyle Then Exit Sub
  If Sel Like "Calc*" Or Sel Like "Deac*" Then
    STyp = Strings.Left(Sel, 4)
  Else
    STyp = Strings.Left(Sel, 3)
  End If
  
  If Sel Like "*HdKey" Then
    HTyp = "HdKey"
    BTyp = "Key"
  ElseIf Sel Like "*Hd" Then
    HTyp = "Hd"
    BTyp = "Cell"
  ElseIf Sel Like "*Cell" Or Sel Like "*Date" Then
    HTyp = "Hd"
    BTyp = Right(Sel, 4)
  ElseIf Sel Like "*Val" Then
    HTyp = "Hd"
    BTyp = "Val"
  End If
  
  TableColumn STyp, Selection, BTyp, HTyp
  
End Sub

Public Sub AddTitle()
  With Selection.Resize(1).Offset(-1, 0)
    .Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromLeftOrAbove
    .style = "BoxTitle"
    .Merge
    .Text = DefTitleText
  End With
End Sub

'Leave Percent, Normal, Hyperlink, Followed Hyperlink
Private Sub RemoveAllDefStyles()
  Dim Item As style
  For Each Item In ActiveWorkbook.Styles: With Item
    If .Name Like "*Accent*" Or _
       .Name Like "Heading*" Or _
       .Name Like "*put" Or _
       .Name Like "Curr*" Or _
       (.Name Like "* *" And Not .Name Like "*Hyperlink*") Or _
       .Name Like "Comm*" Or _
       .Name = "Title" Or _
       .Name = "Total" Or .Name = "Good" Or .Name = "Title" Or .Name = "Bad" Or .Name = "Neutral" _
          Then .Delete
  End With: Next Item
End Sub

Private Sub LoadFontSet()

End Sub

Public Sub UpdateStyles()
  RemoveAllDefStyles
  LoadFontSet

  'All Styles except Normal
  Dim Item As style: For Each Item In ActiveWorkbook.Styles: With Item
    Dim sw As New StyleWrapper
    Set sw.StyleObj = Item
    If StartsWith(.Name, "Act") Then
      .Font.Size = ActiveFontSet.BSize
      .Font.Name = ActiveFontSet.Body
    ElseIf EndsWith(.Name, "Title") Then
      .Font.Size = ActiveFontSet.TSize
      .Font.Name = ActiveFontSet.Head
    ElseIf EndsWith(.Name, Array("Hd", "HdKey", "Head")) Then
      .Font.Size = ActiveFontSet.HSize
      .Font.Name = ActiveFontSet.Head
    ElseIf EndsWith(.Name, Array("Val", "Date")) Then
      .Font.Size = ActiveFontSet.BSize
      .Font.Name = ActiveFontSet.Mono
    ElseIf .Name <> DefStyle Then
      .Font.Size = ActiveFontSet.BSize
      .Font.Name = ActiveFontSet.Body
    End If
    
    'Box and Title Styles
    If StartsWith(.Name, "Box") Then
      .IncludeAlignment = True
    Else
      .IncludeAlignment = False
    End If
    
    'Normal Style
    If .Name = DefStyle And ActiveFontSet.ChgDef Then
      .Font.Name = ActiveFontSet.Body
      .Font.Size = ActiveFontSet.BSize
    End If
    
    'Font Setter Styles
    If StartsWith(.Name, "x") Then
      .Font.Name = Range("FontTable[" & Right(.Name, 4) & "]")
    Else
      .IncludeFont = ActiveFontSet.SetFont
    End If
    
    'Format Setter Styles
    If EndsWith(.Name, Array("Date", "Percent")) Then
      .IncludeNumber = True
    Else
      .IncludeNumber = ActiveFontSet.SetForm
    End If
    
  End With: Next Item
  
  Dim TItem As TableStyle, TElement As TableStyleElement
  For Each TItem In ActiveWorkbook.TableStyles
    Dim TName As String: TName = TItem.Name
    Set TElement = TItem.TableStyleElements.Item(xlWholeTable)
    'Dim ListSheet.Range ("TableStyleTable[StyleName]")
  Next TItem
  
End Sub

Public Sub StyleWriter()

  Dim WriteStyle As style: Set WriteStyle = ActiveWorkbook.Styles("ActWarn")
  Dim Sel As Range:        Set Sel = Selection
  Dim Cols As New Collection
  Dim Vals As New Collection
  
  With Cols
    .add "Name"
    .add "BGColor"
    .add "FTColor"
    .add "PTColor"
    .add "Pattern"
    .add "Italics"
    .add "HasLock"
  End With
  
  Dim x: For x = 1 To Cols.count
    Sel = Cols(x)
    Set Sel = Sel.Offset(0, 1)
  Next x
  
  Set Sel = Sel.Offset(1, -Cols.count)
  
  For Each WriteStyle In ActiveWorkbook.Styles
    
    With WriteStyle
    
      Vals.add .Name
      Vals.add .Interior.ColorIndex
      Vals.add .Font.ColorIndex
      Vals.add .Interior.PatternColorIndex
      Vals.add .Interior.Pattern
      Vals.add .Font.Italic
      Vals.add .Locked
    
    End With
    
    For x = 1 To Cols.count
      Sel = Vals(x)
      Set Sel = Sel.Offset(0, 1)
    Next x
    
    Set Sel = Sel.Offset(1, -Cols.count)
    Set Vals = New Collection
    
  Next WriteStyle

End Sub

Public Sub ReformCell()

  With Selection
    .UnMerge
    .BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
    .Merge
    .style = .style
  End With

End Sub

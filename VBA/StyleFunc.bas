Attribute VB_Name = "StyleFunc"
Option Explicit
Option Compare Text
Option Base 1

'`Style Function Library
'Version 1.1.0

'History
' 1.0.7 - Added SelectColumn'
' 1.0.8 - Added Full Style Application
' 1.0.9 - Added MSize (Monospace Size)
' 1.1.0 - Added SelectColumn Error Proofing

'Current

Private Const DefStyle = "Normal"
Private Const DefTitleText = "Added Title"
Private ActiveFontSet As New ActiveFonts

Public Sub SelectColumn()
  Dim Ref As Range: Set Ref = Selection.Range
  
  If Ref.ListObject Is Nothing Then Exit Sub
  
  Set Ref = Intersect(Ref.ListObject.DataBodyRange, Ref.EntireColumn)
  Ref.Select
End Sub

Private Sub TableColumn(ByVal Typ As String, ByRef Ref As Range, Optional ByVal Body As String = "Cell", Optional ByVal Head As String = "Hd")
  Dim Bd: Set Bd = Intersect(Ref.ListObject.DataBodyRange, Ref.EntireColumn)
  Dim Hd: Set Hd = Intersect(Ref.ListObject.HeaderRowRange, Ref.EntireColumn)
  Bd.Style = Typ + Body
  Hd.Style = Typ + Head
End Sub

Public Sub LookupColumn()
  TableColumn "Lkp", Selection.Range
End Sub

Public Sub CalcColumn()
  TableColumn "Calc", Selection.Range
End Sub

Public Sub DeacColumn()
  TableColumn "Deac", Selection.Range
End Sub

Public Sub InputColumn()
  TableColumn "Inp", Selection.Range
End Sub

Public Sub InternalColumn()
  TableColumn "Int", Selection.Range
End Sub

Public Sub ErrorColumn()
  TableColumn "Err", Selection.Range
End Sub

Public Sub QueryColumn()
  TableColumn "Que", Selection.Range
End Sub

Public Sub FixColumn()
  Dim Sel As Style, STyp As String, HTyp As String, BTyp As String
  Set Sel = Selection.Style
  
  If Sel Like "Act*" Or Sel = DefStyle Then Exit Sub
  If Sel Like "Calc*" Or Sel Like "Deac*" Then
    STyp = Left(Sel, 4)
  Else
    STyp = Left(Sel, 3)
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
  
  TableColumn STyp, Selection.Range, BTyp, HTyp
  
End Sub

Public Sub AddTitle()
  With Selection.Resize(1).Offset(-1, 0)
    .Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromLeftOrAbove
    .Style = "BoxTitle"
    .Merge
    .Text = DefTitleText
  End With
End Sub

'Leave Percent, Normal
Private Sub RemoveAllDefStyles()
  Dim Item As Style
  For Each Item In ActiveWorkbook.Styles: With Item
    If .BuiltIn = True And .Name <> "Normal" And .Name <> "Percent" Then Item.Delete
  End With: Next Item
End Sub

Private Sub LoadFontColorPatterns()

  Dim cnote As Variant, lockval As Variant
  Dim cell As Range, column As Range, font As Range, noter As Range, locked As Range
  Dim Name As String, note As String
  Dim notes() As String
  Dim sty As Style
  Dim count As Integer, x As Integer
  
  Set column = KeySheet.Range("KeyTable[StyleName]")
  count = column.count
  
  For x = 1 To count
  
    Set cell = KeySheet.Range("B" & (x + 2))
    Set noter = KeySheet.Range("H" & (x + 2))
    Set font = KeySheet.Range("J" & (x + 2))
    Set locked = KeySheet.Range("I" & (x + 2))
    Name = cell.Value
    note = noter.Value
    notes = Split(note, ",")
    lockval = locked.Value
    
    Set sty = ActiveWorkbook.Styles(Name)
    
    sty.locked = (lockval = True)
    sty.Interior.Color = cell.Interior.Color
    sty.font.Name = font.Value
    sty.font.Color = cell.font.Color
    
    sty.IncludeFont = True
    sty.IncludeBorder = True
    sty.IncludePatterns = True
    sty.IncludeProtection = True
    sty.IncludeNumber = False
    sty.IncludeAlignment = False
    
    If lockval = "IGNORE" Then sty.IncludeProtection = False
    
    'Add special notes to style
    For Each cnote In notes
      cnote = Trim(cnote)
      Select Case cnote
      
        Case "Pattern Set"
          sty.Interior.Pattern = cell.Interior.Pattern
          sty.Interior.PatternColor = cell.Interior.PatternColor
        Case "P Set"
          sty.Interior.Pattern = cell.Interior.Pattern
          sty.Interior.PatternColor = cell.Interior.PatternColor
        
        Case "Txt Fmt"
          sty.NumberFormat = "@"
          sty.IncludeNumber = True
          
        Case "Date Format"
          sty.NumberFormat = "mm-dd-yy"
          sty.IncludeNumber = True
          
        Case "Italics"
          sty.font.Italic = True
        Case "Ital"
          sty.font.Italic = True
          
        Case "Cent HV"
          sty.VerticalAlignment = xlVAlignCenter
          sty.HorizontalAlignment = xlVAlignCenter
          sty.IncludeAlignment = True
        Case "Centered HV"
          sty.VerticalAlignment = xlVAlignCenter
          sty.HorizontalAlignment = xlVAlignCenter
          sty.IncludeAlignment = True
          
        Case "16 Pt"
          sty.font.Size = ActiveFontSet.TSize
        
        Case "Normal Set"
          sty.IncludeBorder = False
          sty.Interior.Color = xlNone
          
      End Select
    Next cnote
    
  Next x

End Sub

Public Sub UpdateStyles()
  RemoveAllDefStyles
  'All Styles
  Dim Item As Style
  For Each Item In ActiveWorkbook.Styles: With Item
    If EndsWith(.Name, "Title") Then
      .font.Size = ActiveFontSet.TSize
    ElseIf EndsWith(.Name, Array("Hd", "HdKey", "Head")) Then
      .font.Size = ActiveFontSet.HSize
    ElseIf StartsWith(.Name, "Act") Or EndsWith(.Name, Array("Val", "Date")) Or .Name = DefStyle Then
      .font.Size = ActiveFontSet.BSize
    ElseIf .Name = "xMono" Or EndsWith(.Name, "Val") Or EndsWith(.Name, "Date") Then
      .font.Size = ActiveFontSet.MSize
    End If
    
    'Font Setter Styles
    If StartsWith(.Name, "x") Then
      .font.Name = Range("FontTable[" & Right(.Name, 4) & "]")
    End If
    
  End With: Next Item
  
  Dim TItem As TableStyle, TElement As TableStyleElement
  For Each TItem In ActiveWorkbook.TableStyles
    Dim TName As String: TName = TItem.Name
    Set TElement = TItem.TableStyleElements.Item(xlWholeTable)
    'Dim ListSheet.Range ("TableStyleTable[StyleName]")
  Next TItem
  
  LoadFontColorPatterns
  
End Sub

Public Sub StyleWriter()

  Dim WriteStyle As Style: Set WriteStyle = ActiveWorkbook.Styles("ActWarn")
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
    .add "NumStyle"
  End With
  
  Dim x: For x = 1 To Cols.count
    Sel = Cols(x)
    Set Sel = Sel.Offset(0, 1)
  Next x
  
  Set Sel = Sel.Offset(1, -Cols.count)
  
  For Each WriteStyle In ActiveWorkbook.Styles
    
    With WriteStyle
      Vals.add .Name
      Vals.add .Interior.Color
      Vals.add .font.Color
      Vals.add .Interior.PatternColor
      Vals.add .Interior.Pattern
      Vals.add .font.Italic
      Vals.add .NumberFormat
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
    .Style = .Style
  End With
End Sub

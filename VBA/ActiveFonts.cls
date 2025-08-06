VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActiveFonts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Option Base 1

'`ActiveFonts Class Library
'Version 1.0.0

'History
' 1.0.0 - Initial Release

'Current

Property Get TSize() As Integer
  TSize = ListSheet.Range("SizeOverrideTable[Cell Title]").Value
End Property

Property Get HSize() As Integer
  HSize = ListSheet.Range("SizeOverrideTable[Headers]").Value
End Property

Property Get BSize() As Integer
  BSize = ListSheet.Range("SizeOverrideTable[Cells]").Value
End Property

Property Get HeadFont() As String
  HeadFont = ListSheet.Range("FontTable[Head]").Value
End Property

Property Get BodyFont() As String
  BodyFont = ListSheet.Range("FontTable[Body]").Value
End Property

Property Get MonoFont() As String
  MonoFont = ListSheet.Range("FontTable[Mono]").Value
End Property

Property Get CondFont() As String
  CondFont = ListSheet.Range("FontTable[Cond]").Value
End Property

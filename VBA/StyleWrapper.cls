VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StyleWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Option Base 1

'`StyleWrapper Class Library
'Version 1.0.0

'History
' 1.0.0 - Initial Release

'Current

Private obj As Style

Property Set StyleObj(ByRef sty As Style)
  Set obj = sty
End Property

Property Get Name() As String
  Name = obj.Name
End Property

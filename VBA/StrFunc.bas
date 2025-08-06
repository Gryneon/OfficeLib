Attribute VB_Name = "StrFunc"
Option Explicit
Option Compare Text
Option Base 1

'`String Function Library
'Version 1.1.1

'Imports
'Microsoft Scripting Runtime

'History
' 1.0.7 - Fixed Contains check to look for proper InStr return
' 1.0.8 - Functions now accept an array of strings, as well as a normal string
'         Removed ContainsAny, StartsWithAny, EndsWithAny, ReplaceMany
'         Added Substitute that can accept an array, collection, dictionary, or strings
' 1.0.9 - Syntax corrections
' 1.1.0 - Syntax corrections
' 1.1.1 - Substitute now works, and is named Substitute2

'Current

Private Const NoParam = "@@@@@"

'Text Analysis

Public Function Contains(ByVal checktext As String, fortext) As Boolean
  If IsArray(fortext) Then
    Dim Result: Contains = False
    Dim Item: For Each Item In fortext
      If InStr(checktext, Item) > 0 Then
        Contains = True
        Exit Function
      End If
    Next Item
  Else
    Contains = InStr(checktext, fortext) > 0
  End If
End Function

Public Function StartsWith(ByVal checktext As String, fortext) As Boolean
  If IsArray(fortext) Then
    Dim Result: StartsWith = False
    Dim Item: For Each Item In fortext
      If InStr(checktext, Item) = 1 Then
        StartsWith = True
        Exit Function
      End If
    Next Item
  Else
    StartsWith = InStr(checktext, fortext) = 1
  End If
End Function

Private Function ew(ByVal C As String, ByVal f As String) As Boolean
  ew = InStr(C, f) = (Len(C) - Len(f) + 1) And InStr(C, f) > 0
End Function

Public Function EndsWith(ByVal checktext As String, fortext) As Boolean
  If IsArray(fortext) Then
    Dim Result: EndsWith = False
    Dim Item: For Each Item In fortext
      If ew(checktext, Item) Then
        EndsWith = True
        Exit Function
      End If
    Next Item
  Else
    EndsWith = ew(checktext, fortext)
  End If
End Function

'Find & Substitute Variants

Public Function Substitute2(ByVal Source As String, ByRef Find, Optional ByRef Rep) As String
  On Error GoTo Invalid
  Dim i As Integer, rng As Range
  Substitute2 = Source
  Set rng = Find
  For i = 1 To rng.count
    Substitute2 = Replace(Substitute2, Find(i), Rep(i))
  Next i
Exit Function
Invalid:
  Substitute2 = CVErr(xlErrValue)
End Function

'IsNumeric Variants

Public Function IsEmpty2(ByVal checkvalue) As Boolean
  IsEmpty2 = IsEmpty(checkvalue) Or IsNull(checkvalue) Or Len(checkvalue) = 0
End Function

Public Function IsNum(ByVal checkvalue) As Boolean
  IsNum = IsNumeric(checkvalue) Or IsDate(checkvalue)
End Function

Public Function IsText(ByVal checkvalue) As Boolean
  IsText = Not IsNumeric(checkvalue) And Not IsEmpty2(checkvalue) And Not IsDate(checkvalue)
End Function

'IfError Variants

Public Function IfText(ByVal checkvalue, ByVal valueiftext, Optional ByVal trimvalue As Boolean = False) As Variant
  If trimvalue Then checkvalue = Trim(checkvalue)
  
  IfText = IIf(Not IsText(checkvalue), checkvalue, valueiftext)
End Function

Public Function IfNum(ByVal checkvalue, ByVal valueifnum, Optional ByVal trimvalue As Boolean = True) As Variant
  If trimvalue Then checkvalue = Trim(checkvalue)
  
  IfNum = IIf(IsNum(checkvalue) Or IsEmpty2(checkvalue), valueifnum, checkvalue)
End Function

Public Function IfEmpty(ByVal checkvalue, ByVal valueifempty, Optional ByVal trimvalue As Boolean = True) As Variant
  If trimvalue Then checkvalue = Trim(checkvalue)

  IfEmpty = IIf(IsEmpty2(checkvalue), valueifempty, checkvalue)
End Function

Public Function IfError2(ByVal checkvalue, ByVal valueiferror)
  IfError2 = IIf(IsError(checkvalue), valueiferror, checkvalue)
End Function

'Make Plural if needed

Public Function Plural(ByVal initialtext As String, ByVal num As Double, Optional ByVal appendtext As String = "s") As String
  Plural = CStr(num) & " " & initialtext & IIf(num <> 1, appendtext, "")
End Function

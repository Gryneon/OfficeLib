Attribute VB_Name = "ComFunc"
Option Explicit
Option Compare Text
Option Base 1

'Common Function Library
'Version 1.0.1

'Current

Public Function Max2(ByVal x, ByVal y)
  Max2 = IIf(x > y, x, y)
End Function

Public Function Min2(ByVal x, ByVal y)
  Min2 = IIf(x < y, x, y)
End Function

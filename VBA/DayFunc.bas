Attribute VB_Name = "DayFunc"
Option Explicit
Option Compare Text
Option Base 1

'`Date Function Library
'Version 1.1.5

'History
' 1.1.2 - Added Optional Base7 Parameter to DayStr
' 1.1.3 - Added NextMonthOn Function
' 1.1.4 - Added NextWeekOn Function, Removed Base7 Parameter from DayStr
' 1.1.5 - Added NextIncOn, IncStart, IncCode Function

'Current

Public Const DTToday As Integer = -2
Public Const vbSaturday2 As Integer = 0

Public Function IsThisWeek(ByVal datenumber, Optional ByVal startday As Integer = vbMonday) As Boolean

  If datenumber < 0 Then GoTo Invalid

  IsThisWeek = WeekStart() = WeekStart(datenumber)
Exit Function
  
Invalid:
  IsThisWeek = CVErr(xlErrValue)
End Function


Public Function DayCode(Optional ByVal datenumber = DTToday, Optional ByVal base7 As Boolean = False) As Integer

  If datenumber = DTToday Then datenumber = Date
  If datenumber < 0 Then GoTo Invalid
  
  DayCode = IIf(base7 And Int(datenumber) Mod 7 = vbSaturday2, vbSaturday, Int(datenumber) Mod 7)
Exit Function
  
Invalid:
  DayCode = CVErr(xlErrValue)
End Function

Public Function IncCode(ByVal inc, Optional ByVal datenumber = DTToday) As Integer

  If datenumber = DTToday Then datenumber = Date
  If datenumber < 0 Then GoTo Invalid
  
  IncCode = Lng(datenumber) Mod inc
Exit Function
  
Invalid:
  IncCode = CVErr(xlErrValue)
End Function

Public Function WeekStart(Optional ByVal datenumber = DTToday, Optional ByVal startday As Integer = vbMonday) As Date
  
  If datenumber = DTToday Then datenumber = Date
  If datenumber < 0 Then GoTo Invalid
  
  WeekStart = (Int(datenumber / 7) * 7) + startday
Exit Function
  
Invalid:
  WeekStart = CVErr(xlErrValue)
End Function

Public Function IncStart(ByVal inc As Integer, Optional ByVal datenumber = DTToday) As Date
  
  If datenumber = DTToday Then datenumber = Date
  If datenumber < 0 Then GoTo Invalid
  
  IncStart = (Lng(datenumber / inc) * inc)
Exit Function
  
Invalid:
  IncStart = CVErr(xlErrValue)
End Function

Public Function WeekFrom(ByVal datenumber, Optional ByVal todatenumber = DTToday, Optional ByVal startday = vbMonday, Optional ByVal base1index As Boolean = False) As Integer

  If todatenumber = DTToday Then todatenumber = Date
  If todatenumber < 0 Then GoTo Invalid
  
  WeekFrom = Int((datenumber - startday) / 7) - Int((todatenumber - startday) / 7) - CInt(base1index)
Exit Function
  
Invalid:
  WeekFrom = CVErr(xlErrValue)
End Function

Public Function DayStr(Optional ByVal datenumber = DTToday) As String
  
  If datenumber = DTToday Then datenumber = Date
  If datenumber < 0 Then GoTo Invalid
 
  DayStr = WeekdayName(DayCode(datenumber, True))
Exit Function
  
Invalid:
  DayStr = CVErr(xlErrValue)
End Function

Public Function YearStart(Optional ByVal datenumber = DTToday) As Date

  If datenumber = DTToday Then datenumber = Date
  If datenumber < 0 Then GoTo Invalid
  
  YearStart = DateSerial(Year(datenumber), 1, 1)
Exit Function
  
Invalid:
  YearStart = CVErr(xlErrValue)
End Function

Public Function NextMonthOn(ByVal targetday, Optional ByVal count = 1, Optional ByVal datenumber = DTToday)

  If datenumber = DTToday Then datenumber = Date
  If datenumber < 0 Then GoTo Invalid
  
  If day(datenumber) <= targetday Then
    NextMonthOn = DateSerial(Year(datenumber), Month(datenumber) + count - 1, targetday)
  Else
    NextMonthOn = DateSerial(Year(datenumber), Month(datenumber) + count, targetday)
  End If
Exit Function
  
Invalid:
  NextMonthOn = CVErr(xlErrValue)
End Function

Public Function NextWeekOn(ByVal targetday, Optional ByVal count = 1, Optional ByVal datenumber = DTToday)

  If datenumber = DTToday Then datenumber = Date
  If datenumber < 0 Then GoTo Invalid
  
  Dim thisweek As Long

  thisweek = CLng(WeekStart(datenumber))
  
  If DayCode(datenumber, False) <= targetday Then
    NextWeekOn = thisweek + targetday - 2 + (count - 1) * 7
  Else
    NextWeekOn = thisweek + targetday - 2 + count * 7
  End If
Exit Function
  
Invalid:
  NextWeekOn = CVErr(xlErrValue)
End Function

Public Function NextIncOn(ByVal inc, ByVal targetinc, Optional ByVal count = 1, Optional ByVal datenumber = DTToday)

  If datenumber = DTToday Then datenumber = Date
  If datenumber < 0 Then GoTo Invalid
  
  Dim thisinc As Long

  thisinc = CLng(IncStart(inc, datenumber))
  
  If IncCode(datenumber, False) <= targetinc Then
    NextIncOn = thisinc + targetinc + (count - 1) * inc
  Else
    NextIncOn = thisinc + targetinc + count * inc
  End If
Exit Function
  
Invalid:
  NextIncOn = CVErr(xlErrValue)
End Function

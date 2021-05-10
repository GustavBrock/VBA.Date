Attribute VB_Name = "VDateTime"
Option Explicit
'
' VDateTime
' Version 1.0.0
'
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.Date
'
' Functions to replace the native VBA.DateTime functions in queries
' where one or more arguments may be Null or invalid, and the
' native functions would fail.
' For such cases, these functions will not fail but return Null.
'
' License: MIT (http://opensource.org/licenses/mit-license.php)
'
' Required references:
'   None
'
' Required modules:
'   None
'

' Returns the date from Date1 added Number of Intervals.
' Returns Null if any argument is Null or invalid.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateAdd( _
    ByVal Interval As Variant, _
    ByVal Number As Variant, _
    ByVal Date1 As Variant) _
    As Variant
    
    Dim ResultDate  As Variant
    
    ResultDate = Null
    
    On Error Resume Next
    ResultDate = DateAdd(Interval, Number, Date1)
    On Error GoTo 0
        
    VDateAdd = ResultDate
    
End Function

' Returns the difference between Date1 and Date2 as an
' integer count of the Interval passed.
' Returns Null if any required argument is Null or invalid.
' The optional arguments, however, must be valid if passed.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateDiff( _
    ByVal Interval As Variant, _
    ByVal Date1 As Variant, _
    ByVal Date2 As Variant, _
    Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbSunday, _
    Optional ByVal FirstWeekOfYear As VbFirstWeekOfYear = vbFirstJan1) _
    As Variant
    
    Dim Result      As Variant
    
    Result = Null
    
    On Error Resume Next
    Result = DateDiff(Interval, Date1, Date2, FirstDayOfWeek, FirstWeekOfYear)
    On Error GoTo 0
        
    VDateDiff = Result
    
End Function

' Returns the date or time part of Date1 as an integer count of the Interval passed.
' Returns Null if any required argument is Null or invalid.
' The optional arguments, however, must be valid if passed.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDatePart( _
    ByVal Interval As Variant, _
    ByVal Date1 As Variant, _
    Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbSunday, _
    Optional ByVal FirstWeekOfYear As VbFirstWeekOfYear = vbFirstJan1) _
    As Variant
    
    Dim Result      As Variant
    
    Result = Null
    
    On Error Resume Next
    Result = DatePart(Interval, Date1, FirstDayOfWeek, FirstWeekOfYear)
    On Error GoTo 0
        
    VDatePart = Result
    
End Function

' Returns a date value from its year, month, and day part.
' Returns Null if any argument is Null or invalid.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateSerial( _
    ByVal Year As Variant, _
    ByVal Month As Variant, _
    ByVal Day As Variant) _
    As Variant
    
    Dim ResultDate  As Variant
    
    ResultDate = Null
    
    On Error Resume Next
    ResultDate = DateSerial(Year, Month, Day)
    On Error GoTo 0
        
    VDateSerial = ResultDate
    
End Function

' Returns a date value from a string expression.
' Returns Null if the argument is Null or invalid.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateValue( _
    ByVal Date1 As Variant) _
    As Variant
    
    Dim ResultDate  As Variant
    
    ResultDate = Null
    
    On Error Resume Next
    ResultDate = DateValue(Date1)
    On Error GoTo 0
        
    VDateValue = ResultDate
    
End Function

' Returns the day part of Date1 as an integer.
' Returns Null if the argument is Null or invalid.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDay( _
    ByVal Date1 As Variant) _
    As Variant
    
    Dim Result      As Variant
    
    Result = Null
    
    On Error Resume Next
    Result = Day(Date1)
    On Error GoTo 0
        
    VDay = Result
    
End Function

' Returns the hour part of Date1 as an integer.
' Returns Null if the argument is Null or invalid.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VHour( _
    ByVal Time1 As Variant) _
    As Variant
    
    Dim Result      As Variant
    
    Result = Null
    
    On Error Resume Next
    Result = Hour(Time1)
    On Error GoTo 0
        
    VHour = Result
    
End Function

' Returns the minute part of Date1 as an integer.
' Returns Null if the argument is Null or invalid.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VMinute( _
    ByVal Time1 As Variant) _
    As Variant
    
    Dim Result      As Variant
    
    Result = Null
    
    On Error Resume Next
    Result = Minute(Time1)
    On Error GoTo 0
        
    VMinute = Result
    
End Function

' Returns the month part of Date1 as an integer.
' Returns Null if the argument is Null or invalid.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VMonth( _
    ByVal Date1 As Variant) _
    As Variant
    
    Dim Result      As Variant
    
    Result = Null
    
    On Error Resume Next
    Result = Month(Date1)
    On Error GoTo 0
        
    VMonth = Result
    
End Function

' Returns the second part of Date1 as an integer.
' Returns Null if the argument is Null or invalid.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VSecond( _
    ByVal Time1 As Variant) _
    As Variant
    
    Dim Result      As Variant
    
    Result = Null
    
    On Error Resume Next
    Result = Second(Time1)
    On Error GoTo 0
        
    VSecond = Result
    
End Function

' Returns a date (time) value from its hour, minute, and second part.
' Returns Null if any argument is Null or invalid.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VTimeSerial( _
    ByVal Hour As Variant, _
    ByVal Minute As Variant, _
    ByVal Second As Variant) _
    As Variant
    
    Dim ResultDate  As Variant
    
    ResultDate = Null
    
    On Error Resume Next
    ResultDate = TimeSerial(Hour, Minute, Second)
    On Error GoTo 0
        
    VTimeSerial = ResultDate
    
End Function

' Returns a time value from a string expression.
' Returns Null if the argument is Null or invalid.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VTimeValue( _
    ByVal Time1 As Variant) _
    As Variant
    
    Dim ResultTime  As Variant
    
    ResultTime = Null
    
    On Error Resume Next
    ResultTime = TimeValue(Time1)
    On Error GoTo 0
        
    VTimeValue = ResultTime
    
End Function

' Returns weekday of Date1 as an integer.
' Returns Null if any required argument is Null or invalid.
' The optional argument, however, must be valid if passed.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VWeekday( _
    ByVal Date1 As Variant, _
    Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbSunday) _
    As Variant
    
    Dim Result      As Variant
    
    Result = Null
    
    On Error Resume Next
    Result = Weekday(Date1, FirstDayOfWeek)
    On Error GoTo 0
        
    VWeekday = Result
    
End Function

' Returns the year part of Date1 as an integer.
' Returns Null if the argument is Null or invalid.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VYear( _
    ByVal Date1 As Variant) _
    As Variant
    
    Dim Result      As Variant
    
    Result = Null
    
    On Error Resume Next
    Result = Year(Date1)
    On Error GoTo 0
        
    VYear = Result
    
End Function


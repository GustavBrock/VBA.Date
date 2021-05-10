Attribute VB_Name = "VDateBase"
Option Explicit
'
' VDateBase
' Version 1.4.1
'
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.Date
'
' V-versions of basic functions for the entire project.
'
' License: MIT (http://opensource.org/licenses/mit-license.php)
'
' Required references:
'   None
'
' Required modules:
'   DateBase
'

' Returns the count of months of a valid value which is a
' value that can be converted to DtInterval.
' Optionally, also returns the count for an extended value.
'
' Returns 0 (zero) if an invalid value is passed.
'
' Examples:
'   Months = VIntervalMonths("u", True)
'   Months -> 1200
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VIntervalMonths( _
    ByVal Value As Variant, _
    Optional Extended As Boolean) _
    As Integer

    Dim Months      As Integer
    
    If VIsIntervalSetting(Value, Extended) Then
        Months = IntervalMonths(Value, Extended)
    End If

    VIntervalMonths = Months

End Function

' Returns the interval setting from a value of DtInterval for use as
' the Interval parameter of DateAdd, DateDiff, and DatePart.
' Optionally, returns custom (extended) values for Interval in
' DateAddExt, DateDiffExt, and DatePartExt.
'
' Returns Null if an invalid value is passed.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VIntervalSetting( _
    ByVal Interval As Variant, _
    Optional ByVal Extended As Boolean) _
    As Variant

    Dim Symbol  As Variant
    
    Symbol = Null
    
    If VIsInterval(Interval, Extended) Then
        Symbol = IntervalSetting(Interval, Extended)
    End If

    VIntervalSetting = Symbol

End Function

' Returns the DtInterval of a valid value which is a
' value that can be converted to DtInterval.
' Optionally, also validates an extended value.
'
' The case of Value will be ignored.
'
' Returns -1 if an invalid value is passed.
'
' 2021-01-06. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VIntervalValue( _
    ByVal Value As Variant, _
    Optional ByVal Extended As Boolean) _
    As DtInterval
    
    Const ErrorValue    As Long = -1
    
    Dim Interval    As DtInterval
    
    If VIsIntervalSetting(Value, Extended) Then
        Interval = IntervalValue(Value, Extended)
    Else
        Interval = ErrorValue
    End If
   
    VIntervalValue = Interval
    
End Function

' Returns True if Interval is passed a valid value
' of DtInterval.
' Optionally, also returns True for an extended value.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VIsInterval( _
    ByVal Interval As Variant, _
    Optional ByVal Extended As Boolean) _
    As Boolean
    
    Dim Result  As Boolean
    
    If Not IsEmpty(Interval) Then
        If IsNumeric(Interval) Then
            Result = IsInterval(Interval, Extended)
        End If
    End If
    
    VIsInterval = Result
    
End Function

' Returns True if Interval is passed a valid value
' of DtInterval for intervals of one day or higher.
' Optionally, also returns True for an extended value.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VIsIntervalDate( _
    ByVal Interval As Variant, _
    Optional ByVal Extended As Boolean) _
    As Boolean
    
    Dim Result  As Boolean
    
    If Not IsEmpty(Interval) Then
        If IsNumeric(Interval) Then
            Result = IsIntervalDate(Interval, Extended)
        End If
    End If
    
    VIsIntervalDate = Result
    
End Function

' Returns True if the passed Value is a valid setting for
' parameter Interval in DateAdd, DateDiff, and DatePart.
' Optionally, validates custom (extended) values accepted
' by DateAddExt, DateDiffExt, and DatePartExt.
'
' Returns False if any parameter is Null or invalid.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VIsIntervalSetting( _
    ByVal Value As Variant, _
    Optional ByVal Extended As Boolean) _
    As Boolean
    
    Dim Result  As Boolean
    
    Result = IsIntervalSetting(Nz(Value), Extended)
    
    VIsIntervalSetting = Result

End Function

' Returns True if Interval is passed a valid value
' of DtInterval for intervals of less than one day.
' Optionally, also returns True for an extended value.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VIsIntervalTime( _
    ByVal Interval As Variant, _
    Optional ByVal Extended As Boolean) _
    As Boolean
    
    Dim Result  As Boolean
    
    If Not IsEmpty(Interval) Then
        If IsNumeric(Interval) Then
            Result = IsIntervalTime(Interval, Extended)
        End If
    End If
    
    VIsIntervalTime = Result
    
End Function


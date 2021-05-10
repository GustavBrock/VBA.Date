Attribute VB_Name = "VDateSpan"
Option Explicit
'
' VDateSpan
' Version 1.6.3
'
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.Date
'
' Functions for converting between all sorts of date and time systems.
'
' License: MIT (http://opensource.org/licenses/mit-license.php)
'
' Required references:
'   None
'
' Required modules:
'   DateBase
'   DateCalc
'   DateCore
'   DateFind
'   DateMsec
'   DateSpan
'

' Calculates the .beats for the "Swatch Internet Time" from
' a date/time value.
' A such .beat is 1/1000 of a day or 1 minute 26.4 seconds,
' thus the count of .beats is between 0 and 999.
' Returns Null if Date1 is Null or invalid.
'
' The result is by default rounded by +/- half a .beat to
' the nearest integer .beat.
' Optionally, by passing parameter RoundSeconds as False,
' deciseconds will be respected: 4, 8, 2, 6, or 0.
'
' If .beats and times are converted back and forth using the
' functions VBeat and VDateBeat, parameter RoundSeconds must be
' either True or False both ways or inconsistent results will
' be returned.
'
' Beats are counted from Midnight of the Swatch timezone BMT,
' Biel Meantime, which equals the UTC+01.00 timezone.
' Thus, if the local timezone is another, the passed value
' must first be converted to timezone UTC+01.00.
'
' Reference:
'   https://www.swatch.com/en_us/internet-time/
'
' Examples
'   RoundSeconds = True:
'   00:00:00     ->   0
'   00:00:43     ->   0
'   00:00:43.200 ->   0
'   00:00:43.201 ->   1
'   00:00:44     ->   1
'   11:57:50     -> 498
'   11:57:51     -> 499
'   11:59:16     -> 499
'   11:59:17     -> 500
'   12:00:00     -> 500
'   23:57:50     -> 998
'   23:57:51     -> 999
'   23:59:16     -> 999
'   23:59:16.799 -> 999
'   23:59:16.800 ->   0
'   23:59:17     ->   0
'
'   RoundSeconds = False:
'   00:00:00.000 ->   0
'   00:01:26     ->   0
'   00:01:26.399 ->   0
'   00:01:26.400 ->   1
'   00:01:27     ->   1
'   11:59:59     -> 499
'   12:00:00     -> 500
'   23:58:33     -> 998
'   23:58:34     -> 999
'   23:59:59     -> 999
'   23:59:59.999 -> 999
'
' 2018-01-25. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VBeat( _
    ByVal Date1 As Variant, _
    Optional ByVal RoundSeconds As Boolean = True) _
    As Variant
    
    Dim Beats   As Variant
    
    If IsDate(Date1) Then
        Beats = Beat(CDate(Date1), RoundSeconds)
    Else
        Beats = Null
    End If
    
    VBeat = Beats
    
End Function

' Calculates the time from a count of .beats of the
' "Swatch Internet Time".
' A such .beat is 1/1000 of a day or 1 minute 26.4 seconds,
' thus the count of .beats is between 0 and 999.
' Returns Null if Beats is Null or beyond a Long.
'
' The result is by default rounded to the second.
' Optionally, by passing parameter RoundSeconds as False,
' deciseconds will be preserved: 4, 8, 2, 6, or 0.
'
' If .beats and times are converted back and forth using the
' functions VBeat and VDateBeat, parameter RoundSeconds must be
' either True or False both ways or inconsistent results will
' be returned.
'
' Beats are counted from Midnight of the Swatch timezone BMT,
' "Biel Meantime", which equals the UTC+01.00 timezone.
' Thus, if the local timezone is another, the passed value
' must first be converted to timezone UTC+01.00.
'
' Reference:
'   https://www.swatch.com/en_us/internet-time/
'
' Examples.
'   RoundSeconds = True:
'     0 -> 00:00:00
'     1 -> 00:01:26
'     2 -> 00:02:53
'   499 -> 11:58:34
'   500 -> 12:00:00
'   501 -> 12:01:26
'   998 -> 23:57:07
'   999 -> 23:58:34
'  1000 -> 00:00:00
'
'   RoundSeconds = False:
'     0 -> 00:00:00.000
'     1 -> 00:01:26.400
'     2 -> 00:02:52.800
'   998 -> 23:57:07.200
'   999 -> 23:58:33.800
'
' 2018-01-25. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateBeat( _
    ByVal Beats As Variant, _
    Optional ByVal RoundSeconds As Boolean = True) _
    As Variant
    
    Dim TimeValue   As Variant
    
    TimeValue = Null
    
    On Error GoTo Exit_VDateBeat
    
    If IsNumeric(Beats) Then
        TimeValue = DateBeat(CLng(Beats), RoundSeconds)
    End If
    
Exit_VDateBeat:
    
    VDateBeat = TimeValue

End Function

' Returns the ordinal day for a specified date.
' This is sometimes (wrongly) named the Julian day.
' Date1 can be any expression for a valid Date value of VBA.
' Returns Null for invalid expressions.
'
' Examples:
'    100-01-01 ->     1
'   1899-12-30 ->   364
'   1980-03-01 ->    61
'   1981-03-01 ->    60
'   2000-12-31 ->   366
'   9999-12-31 ->   365
'  "1981-04-31"->  Null
'
' 2021-04-02. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VOrdinalDay( _
    ByVal Date1 As Variant) _
    As Variant
    
    Dim Result  As Variant
    
    If IsDate(Date1) Then
        Result = DatePart("y", Date1)
    Else
        Result = Null
    End If
    
    VOrdinalDay = Result
    
End Function


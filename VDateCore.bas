Attribute VB_Name = "VDateCore"
Option Explicit
'
' DateCore
' Version 1.4.0
'
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.Date
'
' Functions for all sorts of calculations related to date and time.
'
' License: MIT (http://opensource.org/licenses/mit-license.php)
'
' Required references:
'   None
'
' Required modules:
'   DateBase
'   DateCalc
'   DateFind
'   DateMsec
'   DateCore
'

' Returns the century of a date as the first year of the century.
' Returns Null if Date1 is Null or invalid.
' Value is 100 for the first century and 9900 for the last.
'
' 2016-04-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VCentury( _
    ByVal Date1 As Variant) _
    As Variant
    
    Dim Result      As Variant
    
    If IsDateExt(Date1) Then
        Result = Century(CDate(Date1))
    Else
        Result = Null
    End If
    
    VCentury = Result

End Function

' An extended replacement for DateAdd, that can handle any
' value set for Interval and Number that DateDiff can return.
' Returns Null if Date1 is Null or parameters are invalid.
'
' The maximum number is given by seconds of the full range of Date:
'
'   SecondsMax = DateDiff("s", #1/1/100#, #12/31/9999 11:59:59 PM#)
'   SecondsMax = 312413759999
'
' The maximum number DateAdd can accept is 2 ^ 31 - 1 or:
'
'   MaxNumber = 2147483647
'
' For larger numbers, intervals are added in a loop, each adding
' the maximum number. Maximum loop count is:
'
'   Max. loops = 312413759999 / 2147483647 = 145
'
' Inside the loop, Date- and TimeSerial are used to clear bit errors
' that would escalate to an error of one second for the outer day of
' the range of Date when adding seconds beyond the equivalent total
' interval of "Days of the range of Date" minus one - in other words:
'
'   Adding seconds from times of 100-01-01 resulting in values
'   later than 9999-12-30.
'   Subtracting seconds from times of 9999-12-31 resulting in values
'   earlier than 100-01-02.
'
' Will also handle additions of all non-native periods.
'
' 2016-02-15. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateAddExt( _
    ByVal Interval As Variant, _
    ByVal Number As Variant, _
    ByVal Date1 As Variant) _
    As Variant
    
    Dim Result  As Variant
    
    Result = Null
    
    If IsDateExt(Date1) And IsNumeric(Number) Then
        Interval = Trim(Nz(Interval))
        If IsIntervalSetting(Interval, True) Then
            On Error Resume Next
            Result = DateAddExt(Interval, CDbl(Number), CDate(Date1))
            On Error GoTo 0
        End If
    End If
    
    VDateAddExt = Result

End Function

' Returns the difference between two dates.
' Will also return the difference in extended
' intervals of half, third, and sixth years.
' Returns Null if Date1 or Date2 is Null or parameters are invalid.
'
' Note, that optional parameters for week settings
' are ignored, as weeks always are handled according
' to ISO 8601.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateDiffExt( _
    ByVal Interval As Variant, _
    ByVal Date1 As Variant, _
    ByVal Date2 As Variant, _
    Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbSunday, _
    Optional ByVal FirstWeekOfYear As VbFirstWeekOfYear = vbFirstJan1) _
    As Variant
    
    Dim Result          As Variant
    
    Result = Null
    
    If IsDateExt(Date1) And IsDateExt(Date2) Then
        Interval = Trim(Nz(Interval))
        If IsIntervalSetting(Interval, True) Then
            On Error Resume Next
            Result = DateDiffExt(Interval, CDate(Date1), CDate(Date2), FirstDayOfWeek, FirstWeekOfYear)
            On Error GoTo 0
        End If
    End If
    
    VDateDiffExt = Result

End Function

' Returns a date that always will fall in the first
' ISO 8601 week of a year.
' If Year is not passed, the current year is used.
' Returns Null if Date1 is Null or parameters are invalid.
'
' 2015-12-04. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateFirstWeekYear( _
    Optional ByVal Year As Variant = 0) _
    As Variant

    Dim Result  As Variant
    
    If VIsYear(Year) Then
        Result = DateFirstWeekYear(CInt(Year))
    Else
        Result = Null
    End If
    
    VDateFirstWeekYear = Result

End Function

' Converts a timespan value to a date value.
' Returns Null if Value is Null or an invalid value.
' Useful only for result date values prior to 1899-12-30 as
' these have a negative numeric value.
'
' 2015-12-15. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateFromTimespan( _
    ByVal Value As Variant) _
    As Variant
  
    Dim Result  As Variant
    
    If IsDateExt(Value) Then
        Result = DateFromTimespan(CDate(Value))
    Else
        Result = Null
    End If
    
    VDateFromTimespan = Result
    
End Function

' Returns the first or earliest date and/or time of an interval of the date
' and/or time passed with an offset specified by Number.
' Optionally, milliseconds may be included.
' Returns Null if Date1 is Null or parameters are invalid.
'
' For intervals of a year down to a week, a time part is stripped.
' For intervals af a day down to millisecond, the time part is rounded down.
' Se the in-line comments for full details.
'
' Any date/time value is accepted and the return value will be correct within
' one millisecond as long as it will be within the range of Date. If not, an
' error is raised.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateIntervalPrimo( _
    ByVal Interval As Variant, _
    ByVal Number As Variant, _
    ByVal Date1 As Variant, _
    Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbSunday, _
    Optional AcceptMilliseconds As Boolean) _
    As Variant
    
    Dim Result  As Variant
    
    Result = Null
    
    If IsDateExt(Date1) And IsNumeric(Number) Then
        Interval = Trim(Nz(Interval))
        If IsIntervalSetting(Interval, True) Then
            On Error Resume Next
            Result = DateIntervalPrimo(Interval, CDbl(Number), CDate(Date1), FirstDayOfWeek, AcceptMilliseconds)
            On Error GoTo 0
        End If
    End If
    
    VDateIntervalPrimo = Result
    
End Function

' Returns the last or latest date and/or time of an interval of the date
' and/or time passed with an offset specified by Number.
' Optionally, milliseconds may be included.
'
' For intervals of a year down to a week, a time part is stripped.
' For intervals af a day down to millisecond, the time part is rounded down.
' Se the in-line comments for full details.
'
' Any date/time value is accepted and the return value will be correct within
' one millisecond as long as it will be within the range of Date. If not, an
' error is raised.
'
' Returns Null if Date1 is Null or parameters are invalid.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateIntervalUltimo( _
    ByVal Interval As Variant, _
    ByVal Number As Variant, _
    ByVal Date1 As Variant, _
    Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbSunday, _
    Optional AcceptMilliseconds As Boolean) _
    As Variant
    
    Dim Result  As Variant
    
    Result = Null
    
    If IsDateExt(Date1) And IsNumeric(Number) Then
        Interval = Trim(Nz(Interval))
        If IsIntervalSetting(Interval, True) Then
            On Error Resume Next
            Result = DateIntervalUltimo(Interval, CDbl(Number), CDate(Date1), FirstDayOfWeek, AcceptMilliseconds)
            On Error GoTo 0
        End If
    End If
    
    VDateIntervalUltimo = Result
    
End Function

' Returns a date that always will fall in the last
' ISO 8601 week of a year.
' If Year is not passed, the current year is used.
' Returns Null if Date1 is Null or parameters are invalid.
'
' 2015-12-04. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateLastWeekYear( _
    Optional ByVal Year As Variant = 0) _
    As Variant

    Dim Result  As Variant
    
    If VIsYear(Year) Then
        Result = DateLastWeekYear(CInt(Year))
    Else
        Result = Null
    End If
    
    VDateLastWeekYear = Result

End Function

' Returns the date part of a date.
' Will also return the extended period types of a date.
' Will return the correct ISO week number with parameters:
'
'    FirstDayOfWeek = vbMonday
'    FirstWeekOfYear = vbFirstFourDays
'
' equivalent to:
'
'    FirstDayOfWeek = 2
'    FirstWeekOfYear = 2
'
' Returns Null if Date1 is Null or parameters are invalid.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDatePartExt( _
    ByVal Interval As Variant, _
    ByVal Date1 As Variant, _
    Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbSunday, _
    Optional ByVal FirstWeekOfYear As VbFirstWeekOfYear = vbFirstJan1) _
    As Variant

    Dim Result  As Variant
    
    Result = Null
    
    If IsDateExt(Date1) Then
        Interval = Trim(Nz(Interval))
        If IsIntervalSetting(Interval, True) Then
            On Error Resume Next
            Result = DatePartExt(Interval, CDate(Date1), FirstDayOfWeek, FirstWeekOfYear)
            On Error GoTo 0
        End If
    End If
    
    VDatePartExt = Result

End Function

' Returns a date value from its year, month, day,
' hour, minute, and second part.
' Except for year, default values are used for parameters omitted.
' Returns Null if any parameter is Null or invalid.
'
' Will accept any combination of parameters that can build a
' date/time within the range of Date, including dates before
' 1899-12-30.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateTimeSerial( _
    ByVal Year As Variant, _
    Optional ByVal Month As Variant = 1, _
    Optional ByVal Day As Variant = 1, _
    Optional ByVal Hour As Variant = 0, _
    Optional ByVal Minute As Variant = 0, _
    Optional ByVal Second As Variant = 0) _
    As Variant
    
    Dim ResultDate  As Variant
    
    ResultDate = Null
    
    On Error Resume Next
    ResultDate = DateTimeSerial(Year, Month, Day, Hour, Minute, Second)
    On Error GoTo 0
        
    VDateTimeSerial = ResultDate
    
End Function

' Converts a date value to a timespan value.
' Returns Null if Value is Null or an invalid value.
' Useful only for date values prior to 1899-12-30 as
' these have a negative numeric value.
'
' 2015-12-15. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateToTimespan( _
    ByVal Value As Variant) _
    As Variant

    Dim Result  As Variant
    
    If IsDateExt(Value) Then
        Result = DateToTimespan(CDate(Value))
    Else
        Result = Null
    End If
  
    VDateToTimespan = Result

End Function

' Returns the day of the month like the native VBA.Day.
' However, the ultimo date(s) will always be returned as day 30.
' Returns Null if Date1 is Null or invalid.
'
' Examples:
'   VDay30(#2000-01-29#) -> 29
'   VDay30(#2000-01-30#) -> 30
'   VDay30(#2000-01-31#) -> 30
'   VDay30(#2000-02-27#) -> 27
'   VDay30(#2000-02-28#) -> 30
'   VDay30(#2000-02-29#) -> 30
'   VDay30(#2000-03-29#) -> 29
'   VDay30(#2000-03-31#) -> 30
'   VDay30(Null)         -> Null
'
' 2019-01-26. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function VDay30( _
    ByVal Date1 As Variant) _
    As Variant
    
    Dim Result      As Variant
    
    If IsDateExt(Date1) Then
        Result = Day30(CDate(Date1))
    Else
        Result = Null
    End If
    
    VDay30 = Result

End Function

' Returns the decade of a date as the first year of the decade.
' Returns Null if Date1 is Null or invalid.
' Value is 100 for the first decade and 9990 for the last.
'
' 2016-04-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDecade( _
    ByVal Date1 As Variant) _
    As Variant
    
    Dim Result      As Variant
    
    If IsDateExt(Date1) Then
        Result = Decade(CDate(Date1))
    Else
        Result = Null
    End If
    
    VDecade = Result

End Function

' Returns the fortnight of a date based on the ISO 8601 week number.
' Returns Null if Date1 is Null or invalid.
' The related ISO year is returned by ref.
'
' 2016-01-06. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VFortnight( _
    ByVal Date1 As Variant, _
    Optional ByRef IsoYear As Variant) _
    As Variant

    Dim Year    As Integer
    Dim Result  As Variant
    
    If IsDateExt(Date1) Then
        If VIsYear(IsoYear) Then
            Year = CInt(IsoYear)
        Else
            Year = VBA.Year(CDate(Date1))
        End If
        Result = Fortnight(CDate(Date1), Year)
        IsoYear = Year
    Else
        Result = Null
        IsoYear = Null
    End If
    
    VFortnight = Result
    
End Function

' Returns the weekday within a fortnight based on an ISO 8601 week.
' Parameter FirstDayOfWeek is ignored as the first day must be Monday.
'
' Return values for the first seven days, Monday-Sunday, are 1 to 7.
' Return values for the next seven days, Monday-Sunday, are 8 to 14.
' Note that a fortnight of 27 will only have days of the first week.
' Returns Null if Date1 is Null or invalid.
'
' 2016-02-06. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VFortnightday( _
    ByVal Date1 As Variant, _
    Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbMonday) _
    As Variant
    
    Dim Result  As Variant
    
    If IsDateExt(Date1) Then
        FirstDayOfWeek = vbMonday
        Result = Fortnightday(CDate(Date1), FirstDayOfWeek)
    Else
        Result = Null
    End If
    
    VFortnightday = Result
    
End Function

' Returns the count of a fortnightday between two dates.
' Returns Null if Date1 or Date2 is Null or invalid.
'
' 2016-02-14. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VFortnightdayCount( _
    ByVal Date1 As Variant, _
    ByVal Date2 As Variant) _
    As Variant
    
    Dim Result  As Variant
    
    If IsDateExt(Date1) And IsDateExt(Date2) Then
        Result = FortnightdayCount(CDate(Date1), CDate(Date2))
    Else
        Result = Null
    End If
    
    VFortnightdayCount = Result
    

End Function

' Returns the count of fortnights based on the ISO 8601 week count of a year.
' If Year is not passed, the current year is used.
' Returns Null if Year is Null or an invalid value.
'
' 2016-02-14. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VFortnightsOfYear( _
    Optional ByVal Year As Variant = 0) _
    As Variant

    Dim Result  As Variant

    If VIsYear(Year) Then
        Result = FortnightsOfYear(CInt(Year))
    Else
        Result = Null
    End If
    
    VFortnightsOfYear = Result

End Function

' Returns True if Century can be a century.
' Returns Null if Century is Null or an invalid value.
'
' 2016-04-11. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VIsCentury( _
    ByVal Century As Variant) _
    As Variant

    Dim Result  As Variant
    
    Result = Null

    On Error Resume Next
    Result = IsCentury(Century)
    On Error GoTo 0
    
    VIsCentury = Result
    
End Function

' Returns True if Decade can be a decade.
' Returns Null if Decade is Null or an invalid value.
'
' 2016-04-11. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VIsDecade( _
    ByVal Decade As Variant) _
    As Variant

    Dim Result  As Variant
    
    Result = Null

    On Error Resume Next
    Result = IsDecade(Decade)
    On Error GoTo 0
    
    VIsDecade = Result
    
End Function

' Returns True if Fortnight can be a fortnight of Year.
' If Year is not specified, the current year is used.
' Returns Null if Fortnight or Year is Null or an invalid value.
'
' 2016-02-14. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VIsFortnight( _
    ByVal Fortnight As Variant, _
    Optional ByVal Year As Variant = 0) _
    As Variant

    Dim Result  As Variant
    
    Result = Null

    On Error Resume Next
    Result = IsFortnight(Fortnight, Year)
    On Error GoTo 0
    
    VIsFortnight = Result
    
End Function

' Returns True if Year is a leap year.
' If Year is not passed, the current year is used.
' Returns Null if Year is Null or an invalid value.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VIsLeapYear( _
    Optional ByVal Year As Variant = 0) _
    As Variant

    Dim LeapYear    As Variant
    
    LeapYear = Null

    On Error Resume Next
    LeapYear = IsLeapYear(Year)
    On Error GoTo 0
    
    VIsLeapYear = LeapYear
    
End Function

' Returns True if Millenium can be a millenium.
' Returns Null if Century is Null or an invalid value.
'
' 2016-04-11. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VIsMillenium( _
    ByVal Millenium As Variant) _
    As Variant

    Dim Result  As Variant
    
    Result = Null

    On Error Resume Next
    Result = IsMillenium(Millenium)
    On Error GoTo 0
    
    VIsMillenium = Result
    
End Function

' Returns True if Month can be a month.
' Returns Null if Month is Null or an invalid value.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VIsMonth( _
    ByVal Month As Variant) _
    As Variant

    Dim Result              As Variant
    
    Result = Null

    On Error Resume Next
    Result = IsMonth(Month)
    On Error GoTo 0
    
    VIsMonth = Result
    
End Function

' Returns True if Quarter can be a quarter.
' Returns Null if Quarter is Null or an invalid value.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VIsQuarter( _
    ByVal Quarter As Variant) _
    As Variant

    Dim Result              As Variant
    
    Result = Null

    On Error Resume Next
    Result = IsQuarter(Quarter)
    On Error GoTo 0
    
    VIsQuarter = Result
    
End Function

' Returns True if Semimonth can be a semimonth.
' Returns Null if Semimonth is Null or an invalid value.
'
' 2016-02-15. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VIsSemimonth( _
    ByVal Semimonth As Variant) _
    As Variant

    Dim Result              As Variant
    
    Result = Null

    On Error Resume Next
    Result = IsSemimonth(Semimonth)
    On Error GoTo 0
    
    VIsSemimonth = Result
    
End Function

' Returns True if Semiyear can be a semiyear.
' Returns Null if Semiyear is Null or an invalid value.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VIsSemiyear( _
    ByVal Semiyear As Variant) _
    As Variant

    Dim Result              As Variant
    
    Result = Null

    On Error Resume Next
    Result = IsSemiyear(Semiyear)
    On Error GoTo 0
    
    VIsSemiyear = Result
    
End Function

' Returns True if Sextayear can be a sextayear.
' Returns Null if Sextayear is Null or an invalid value.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VIsSextayear( _
    ByVal Sextayear As Variant) _
    As Variant

    Dim Result              As Variant
    
    Result = Null

    On Error Resume Next
    Result = IsSextayear(Sextayear)
    On Error GoTo 0
    
    VIsSextayear = Result
    
End Function

' Returns True if Tertiayear can be a tertiayear.
' Returns Null if Tertiayear is Null or an invalid value.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VIsTertiayear( _
    ByVal Tertiayear As Variant) _
    As Variant

    Dim Result              As Variant
    
    Result = Null

    On Error Resume Next
    Result = IsTertiayear(Tertiayear)
    On Error GoTo 0
    
    VIsTertiayear = Result
    
End Function

' Returns True if IsoWeek can be an ISO 8601 week of IsoYear.
' If Year is not specified, the current year is used.
' Returns Null if Week or Year is Null or an invalid value.
'
' 2016-01-06. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VIsWeek( _
    ByVal IsoWeek As Variant, _
    Optional ByVal IsoYear As Variant = 0) _
    As Variant

    Dim Result  As Variant
    
    Result = Null
    
    On Error Resume Next
    Result = IsWeek(IsoWeek, IsoYear)
    On Error GoTo 0
    
    VIsWeek = Result
    
End Function

' Returns True if Year can be a year.
' Returns Null if Year is Null or an invalid value.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VIsYear( _
    ByVal Year As Variant) _
    As Variant

    Dim Result              As Variant
    
    Result = Null

    On Error Resume Next
    Result = IsYear(Year)
    On Error GoTo 0
    
    VIsYear = Result
    
End Function

' Returns the millenium of a date as the first year of the millenium.
' Returns Null if Date1 is Null or invalid.
' Value is 0 for the first millenium and 9000 for the last.
'
' 2016-04-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VMillenium( _
    ByVal Date1 As Variant) _
    As Variant
    
    Dim Result      As Variant
    
    If IsDateExt(Date1) Then
        Result = Millenium(CDate(Date1))
    Else
        Result = Null
    End If
    
    VMillenium = Result

End Function

' Returns the quarter of a date.
' Value is 1 for the first quarter of the year,
' 4 for the last.
' Returns Null if Date1 is Null or invalid.
'
' 2015-12-21. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VQuarter( _
    ByVal Date1 As Variant) _
    As Variant
    
    Dim Result      As Variant
    
    If IsDateExt(Date1) Then
        Result = Quarter(CDate(Date1))
    Else
        Result = Null
    End If
    
    VQuarter = Result

End Function

' Returns the second of a date.
' Returns Null if Time1 is Null or invalid.
'
' Will return correct reading of seconds of the last day of Date (9999-12-31).
' See function DateTest.TheLastSeconds for full information.
'
' 2016-02-14. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VSecondExt( _
    ByVal Time1 As Variant) _
    As Variant
    
    Dim Result      As Variant
    
    If IsDateExt(Time1) Then
        Result = SecondExt(CDate(Time1))
    Else
        Result = Null
    End If
    
    VSecondExt = Result

End Function

' Returns the semimonth of a date.
' Returns Null if Date1 is Null or invalid.
'
' 2016-02-14. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VSemimonth( _
    ByVal Date1 As Variant) _
    As Variant
    
    Dim Result      As Variant
    
    If IsDateExt(Date1) Then
        Result = Semimonth(CDate(Date1))
    Else
        Result = Null
    End If
    
    VSemimonth = Result

End Function

' Returns the semiyear of a date.
' Value is 1 for the first semiyear of a year,
' 2 for the second.
' Returns Null if Date1 is Null or invalid.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VSemiyear( _
    ByVal Date1 As Variant) _
    As Variant
    
    Dim Result      As Variant
    
    If IsDateExt(Date1) Then
        Result = Semiyear(CDate(Date1))
    Else
        Result = Null
    End If
    
    VSemiyear = Result

End Function

' Returns the sextayear of a date.
' Value is 1 for the first sextayear of a year,
' 2 for the second.
' Returns Null if Date1 is Null or invalid.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VSextayear( _
    ByVal Date1 As Variant) _
    As Variant
    
    Dim Result      As Variant
    
    If IsDateExt(Date1) Then
        Result = Sextayear(CDate(Date1))
    Else
        Result = Null
    End If
    
    VSextayear = Result

End Function

' Returns the tertiayear of a date.
' Value is 1 for the first tertiayear of a year,
' 2 for the second.
' Returns Null if Date1 is Null or invalid.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VTertiayear( _
    ByVal Date1 As Variant) _
    As Variant
    
    Dim Result      As Variant
    
    If IsDateExt(Date1) Then
        Result = Tertiayear(CDate(Date1))
    Else
        Result = Null
    End If
    
    VTertiayear = Result

End Function

' Returns an invalid numeric negative time Date1
' as its positive equivalent.
' Returns Null if Date1 is Null or invalid.
'
' 2015-12-15. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VTimeEmend( _
    ByVal Date1 As Variant) _
    As Variant
  
    Dim ResultDate  As Variant
    
    If IsDateExt(Date1) Then
        ' Emend the time Date1.
        ResultDate = TimeEmend(CDate(Date1))
    Else
        ResultDate = Null
    End If
    
    VTimeEmend = ResultDate
  
End Function

' Returns the ISO 8601 week of a date.
' The related ISO year is returned by ref.
' Returns Null if Date1 is Null or invalid.
'
' 2016-02-06. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VWeek( _
    ByVal Date1 As Variant, _
    Optional ByRef IsoYear As Variant) _
    As Variant

    Dim Year    As Integer
    Dim Result  As Variant
    
    If IsDateExt(Date1) Then
        If VIsYear(IsoYear) Then
            Year = CInt(IsoYear)
        Else
            Year = VBA.Year(CDate(Date1))
        End If
        Result = Week(CDate(Date1), Year)
        IsoYear = Year
    Else
        Result = Null
        IsoYear = Null
    End If
    
    VWeek = Result
    
End Function

' Returns the ISO 8601 week count of a year.
' If Year is not passed, the current year is used.
' Returns Null if Year is Null or an invalid value.
'
' 2015-12-04. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VWeeksOfYear( _
    Optional ByVal Year As Variant) _
    As Variant

    Dim Weeks   As Variant
    
    Weeks = Null

    On Error Resume Next
    Weeks = WeeksOfYear(Year)
    On Error GoTo 0
    
    VWeeksOfYear = Weeks

End Function

' Returns the year of a date.
' Returns Null if Date1 is Null or invalid.
'
' 2015-12-21. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VYear( _
    ByVal Date1 As Variant) _
    As Variant
    
    Dim Result      As Variant
    
    If IsDateExt(Date1) Then
        Result = Year(CDate(Date1))
    Else
        Result = Null
    End If
    
    VYear = Result

End Function


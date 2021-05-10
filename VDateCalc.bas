Attribute VB_Name = "VDateCalc"
Option Explicit
'
' VDateCalc
' Version 1.4.1
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
'   DateCore
'   DateFind
'   DateMsec
'

' Returns the difference in full years from DateOfBirth to current date,
' optionally to another date.
' Returns zero if AnotherDate is earlier than DateOfBirth.
' Returns Null if DateOfBirth is Null.
'
' Calculates correctly for:
'   leap years
'   dates of 29. February
'   date/time values with embedded time values
'   any date/time value of data type Date.
'
' DateAdd() is used for check for month end of February as it correctly
' returns Feb. 28th when adding a count of years to dates of Feb. 29th
' when the resulting year is a common year.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VAge( _
    ByVal DateOfBirth As Variant, _
    Optional ByVal AnotherDate As Variant) _
    As Variant
    
    Dim Years   As Variant
      
    If IsDateExt(DateOfBirth) Then
        Years = Age(CDate(DateOfBirth), AnotherDate)
    Else
        Years = Null
    End If
    
    VAge = Years
  
End Function

' Returns the age as it will be at the 30th of April of the current year.
'
' Calculates correctly for:
'   leap years
'   dates of 29. February
'   date/time values with embedded time values
'   any date/time value of data type Date
'
' DateAdd() is used for check for month end of February as it correctly
' returns Feb. 28th when adding a count of years to dates of Feb. 29th
' when the resulting year is a common year.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VAgeLeague( _
    ByVal DateOfBirth As Variant) _
    As Variant
    
    Dim Years   As Variant
      
    If IsDateExt(DateOfBirth) Then
        Years = AgeLeague(CDate(DateOfBirth))
    Else
        Years = Null
    End If
    
    VAgeLeague = Years
  
End Function

' Returns the difference in full Months from DateOfBirth to current date,
' optionally to another date.
' Returns zero if AnotherDate is earlier than DateOfBirth.
' Returns Null if DateOfBirth is Null.
'
' Calculates correctly for:
'   leap Months
'   dates of 29. February
'   date/time values with embedded time values
'   any date/time value of data type Date
'
' DateAdd() is used for check for month end of February as it correctly
' returns Feb. 28th when adding a count of Months to dates of Feb. 29th
' when the resulting year is a common year.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VAgeMonths( _
    ByVal DateOfBirth As Variant, _
    Optional ByVal AnotherDate As Variant) _
    As Variant
    
    Dim Months  As Variant
      
    If IsDateExt(DateOfBirth) Then
        Months = AgeMonths(CDate(DateOfBirth), AnotherDate)
    Else
        Months = Null
    End If
    
    VAgeMonths = Months
  
End Function

' Returns the rounded up difference in full years from DateOfBirth to
' current date, optionally to another date.
' Returns zero if AnotherDate is earlier than DateOfBirth.
' Returns Null if DateOfBirth is Null.
'
' Calculates correctly for:
'   leap years
'   dates of 29. February
'   date/time values with embedded time values
'   any date/time value of data type Date.
'
' DateAdd() is used for check for month end of February as it correctly
' returns Feb. 28th when adding a count of years to dates of Feb. 29th
' when the resulting year is a common year.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VAgeRoundUp( _
    ByVal DateOfBirth As Variant, _
    Optional ByVal AnotherDate As Variant) _
    As Variant
    
    Dim Years   As Variant
      
    If IsDateExt(DateOfBirth) Then
        Years = AgeRoundUp(CDate(DateOfBirth), AnotherDate)
    Else
        Years = Null
    End If
    
    VAgeRoundUp = Years
  
End Function

' Returns the difference in full centuries between Date1 and Date2.
' Returns Null if either Date1 or Date2 is Null.
'
' Calculates correctly for:
'   negative differences
'   leap years
'   dates of 29. February
'   date/time values with embedded time values
'   any date/time value of data type Date
'
' Optionally returns negative counts rounded down to provide a
' linear sequence of century counts.
' For a given Date1, if Date2 is decreased stepwise one century from
' returning a positive count to returning a negative count, one or two
' occurrences of count zero will be returned.
' If LinearSequence is False, the sequence will be:
'   3, 2, 1, 0,  0, -1, -2
' If LinearSequence is True, the sequence will be:
'   3, 2, 1, 0, -1, -2, -3
'
' If LinearSequence is False, reversing Date1 and Date2 will return
' results of same absolute Value, only the sign will change.
' This behaviour mimics that of Fix().
' If LinearSequence is True, reversing Date1 and Date2 will return
' results where the negative count is offset by -1.
' This behaviour mimics that of Int().
'
' DateAdd() is used for check for month end of February as it correctly
' returns Feb. 28th when adding a count of centuries to dates of Feb. 29th
' when the resulting year is a common year.
'
' 2016-04-11. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VCenturies( _
    ByVal Date1 As Variant, _
    ByVal Date2 As Variant, _
    Optional ByVal LinearSequence As Variant = Null) _
    As Variant
    
    Dim CenturyCount    As Variant
    
    If IsDateExt(Date1) And IsDateExt(Date2) Then
        If IsNumeric(LinearSequence) Then
            LinearSequence = CBool(LinearSequence)
        Else
            LinearSequence = False
        End If
        CenturyCount = Centuries(CDate(Date1), CDate(Date2), LinearSequence)
    Else
        CenturyCount = Null
    End If
    VCenturies = CenturyCount
  
End Function

' Calculates Easter Sunday for year 1583 to 4099.
' Returns the date of Easter Sunday for the passed year.
' Easter Sunday is the Sunday following the Paschal Full Moon
' (PFM) date for the year.
'
' Argument Year must be a year between 1583 and 4099.
' Values outside this range will return non-verified results.
' Null is returned for non-numeric values passed.
'
' This algorithm is an arithmetic interpretation of the three step
' Easter Dating Method developed by Ron Mallen 1985, as a vast
' improvement on the method described in the Common Prayer Book.
' Because this algorithm is a direct translation of the
' official tables, it can be easily proved to be 100% correct.
'
'   Main source:
'       Astronomical Society of South Australia Inc.
'       https://www.assa.org.au/resources/more-articles/easter-dating-method/
'
'       Research by Ronald W. Mallen, Adelaide, Australia.
'       Simplified Easter Dating Method produced by Ronald W. Mallen.
'       Programming algorithm by Greg Mallen.
'
'   Other sources:
'       https://en.wikipedia.org/wiki/Metonic_cycle
'       https://en.wikipedia.org/wiki/Golden_number_(time)
'       https://en.wikipedia.org/wiki/Ecclesiastical_full_moon
'       https://mathlair.allfunandgames.ca/easter.php
'
' 2020-11-14. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateEasterSunday( _
    ByVal Year As Variant) _
    As Variant
    
    Dim EasterSunday    As Variant
    
    EasterSunday = Null
    
    If IsNumeric(Year) Then
        If IsYear(Year) Then
            EasterSunday = DateEasterSunday(Year)
        End If
    End If
    
    VDateEasterSunday = EasterSunday

End Function

' Returns the difference in full months between Date1 and Date2.
' Returns Null if either Date1 or Date2 is Null.
'
' Calculates correctly for:
'   negative differences
'   leap years
'   dates of 29. February
'   date/time values with embedded time values
'   any date/time value of data type Date
'
' Optionally returns negative counts rounded down to provide a
' linear sequence of month counts.
' For a given Date1, if Date2 is decreased stepwise one month from
' returning a positive count to returning a negative count, one or two
' occurrences of count zero will be returned.
' If LinearSequence is False, the sequence will be:
'   3, 2, 1, 0,  0, -1, -2
' If LinearSequence is True, the sequence will be:
'   3, 2, 1, 0, -1, -2, -3
'
' If LinearSequence is False, reversing Date1 and Date2 will return
' results of same absolute Value, only the sign will change.
' This behaviour mimics that of Fix().
' If LinearSequence is True, reversing Date1 and Date2 will return
' results where the negative count is offset by -1.
' This behaviour mimics that of Int().
'
' DateAdd() is, when adding a count of months to dates of 31th (29th),
' used for check for month end as it correctly returns the 30th (28th)
' when the resulting month has 30 or less days.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateParts( _
    ByVal Interval As Variant, _
    ByVal Date1 As Variant, _
    ByVal Date2 As Variant, _
    Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbSunday, _
    Optional ByVal FirstWeekOfYear As VbFirstWeekOfYear = vbFirstJan1, _
    Optional ByVal LinearSequence As Variant = Null) _
    As Variant
    
    Dim Parts   As Variant
    
    Parts = Null
    
    Interval = Trim(Nz(Interval))
    If IsIntervalSetting(Interval, True) Then
        If IsDateExt(Date1) And IsDateExt(Date2) Then
            If IsNumeric(LinearSequence) Then
                LinearSequence = CBool(LinearSequence)
            Else
                LinearSequence = False
            End If
            Parts = DateParts(Interval, CDate(Date1), CDate(Date2), FirstDayOfWeek, FirstWeekOfYear, LinearSequence)
        End If
    End If
    
    VDateParts = Parts

End Function

' Returns the count of days of the month of Date1.
' Returns Null if Date1 is Null.
'
' 2016-02-14. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDaysInMonth( _
    ByVal Date1 As Variant) _
    As Variant
  
    Dim Result  As Variant
  
    If IsDateExt(Date1) Then
        Result = DaysInMonth(CDate(Date1))
    Else
        Result = Null
    End If
  
    VDaysInMonth = Result
  
End Function

' Returns the count of days of the year of Date1.
' Returns Null if Date1 is Null.
'
' 2016-02-14. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDaysInYear( _
    ByVal Date1 As Variant) _
    As Variant
  
    Dim Result  As Variant
  
    If IsDateExt(Date1) Then
        Result = DaysInYear(CDate(Date1))
    Else
        Result = Null
    End If
  
    VDaysInYear = Result
  
End Function

' Returns the difference in full decades between Date1 and Date2.
' Returns Null if either Date1 or Date2 is Null.
'
' Calculates correctly for:
'   negative differences
'   leap years
'   dates of 29. February
'   date/time values with embedded time values
'   any date/time value of data type Date
'
' Optionally returns negative counts rounded down to provide a
' linear sequence of decade counts.
' For a given Date1, if Date2 is decreased stepwise one decade from
' returning a positive count to returning a negative count, one or two
' occurrences of count zero will be returned.
' If LinearSequence is False, the sequence will be:
'   3, 2, 1, 0,  0, -1, -2
' If LinearSequence is True, the sequence will be:
'   3, 2, 1, 0, -1, -2, -3
'
' If LinearSequence is False, reversing Date1 and Date2 will return
' results of same absolute Value, only the sign will change.
' This behaviour mimics that of Fix().
' If LinearSequence is True, reversing Date1 and Date2 will return
' results where the negative count is offset by -1.
' This behaviour mimics that of Int().
'
' DateAdd() is used for check for month end of February as it correctly
' returns Feb. 28th when adding a count of decades to dates of Feb. 29th
' when the resulting year is a common year.
'
' 2016-04-11. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDecades( _
    ByVal Date1 As Variant, _
    ByVal Date2 As Variant, _
    Optional ByVal LinearSequence As Variant = Null) _
    As Variant
    
    Dim DecadeCount As Variant
    
    If IsDateExt(Date1) And IsDateExt(Date2) Then
        If IsNumeric(LinearSequence) Then
            LinearSequence = CBool(LinearSequence)
        Else
            LinearSequence = False
        End If
        DecadeCount = Decades(CDate(Date1), CDate(Date2), LinearSequence)
    Else
        DecadeCount = Null
    End If
    VDecades = DecadeCount
  
End Function

' Returns the count of fortnights based on the ISO 8601 week count of a year of a date.
' Returns Null if Date1 is Null.
'
' 2016-02-14. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VFortnightsOfYearOfDate( _
    ByVal Date1 As Variant) _
    As Variant

    Dim Fortnight   As Variant

    If IsDateExt(Date1) Then
        Fortnight = FortnightsOfYearOfDate(CDate(Date1))
    Else
        Fortnight = Null
    End If
    
    VFortnightsOfYearOfDate = Fortnight

End Function

' Returns the count of fortnights based on the ISO 8601 week count of a year span.
' Returns Null if Year is Null or an invalid value.
'
' 2016-02-14. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VFortnightsOfYears( _
    ByVal Year1 As Variant, _
    ByVal Year2 As Variant) _
    As Variant

    Dim Result  As Variant
    
    Result = Null

    On Error Resume Next
    Result = FortnightsOfYears(Year1, Year2)
    On Error GoTo 0
    
    VFortnightsOfYears = Result

End Function

' Returns True if Date1 is of a leap year.
' Returns Null if Date1 is Null.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VIsDateLeapYear( _
    ByVal Date1 As Variant) _
    As Variant

    Dim LeapYear    As Variant

    If IsDateExt(Date1) Then
        LeapYear = IsLeapYear(Year(CDate(Date1)))
    Else
        LeapYear = Null
    End If
    
    VIsDateLeapYear = LeapYear
    
End Function

' Returns True if Date1 is the first day of the month.
' Returns Null if Date1 is Null.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VIsDatePrimoMonth( _
    ByVal Date1 As Variant) _
    As Variant

    Dim Primo   As Variant
    
    If IsDateExt(Date1) Then
        Primo = IsDatePrimoMonth(CDate(Date1))
    Else
        Primo = Null
    End If
    
    VIsDatePrimoMonth = Primo
    
End Function

' Returns True if Date1 is the first day of the quarter.
' Returns Null if Date1 is Null.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VIsDatePrimoQuarter( _
    ByVal Date1 As Variant) _
    As Variant

    Dim Primo       As Variant

    If IsDateExt(Date1) Then
        Primo = IsDatePrimoQuarter(CDate(Date1))
    Else
        Primo = Null
    End If
    
    VIsDatePrimoQuarter = Primo
    
End Function

' Returns True if Date1 is the first day of the week.
' Returns Null if Date1 is Null.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VIsDatePrimoWeek( _
    ByVal Date1 As Variant, _
    Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbSunday) _
    As Variant

    Dim Primo   As Variant

    If IsDateExt(Date1) Then
        Primo = IsDatePrimoWeek(CDate(Date1), FirstDayOfWeek)
    Else
        Primo = Null
    End If
    
    VIsDatePrimoWeek = Primo
    
End Function

' Returns True if Date1 is the first day of the year.
' Returns Null if Date1 is Null.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VIsDatePrimoYear( _
    ByVal Date1 As Variant) _
    As Variant

    Dim Primo   As Variant

    If IsDateExt(Date1) Then
        Primo = IsDatePrimoYear(CDate(Date1))
    Else
        Primo = Null
    End If
    
    VIsDatePrimoYear = Primo
    
End Function

' Returns True if Date1 is the last day of the month.
' If Include2830 is True, also February 28th of leap years
' and the 30th of any month will be regarded as ultimo.
' Returns Null if Date1 is Null.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VIsDateUltimoMonth( _
    ByVal Date1 As Variant, _
    Optional ByVal Include2830 As Boolean) _
    As Variant

    Dim Ultimo  As Variant
    
    If IsDateExt(Date1) Then
        Ultimo = IsDateUltimoMonth(CDate(Date1), Include2830)
    Else
        Ultimo = Null
    End If
    
    VIsDateUltimoMonth = Ultimo
    
End Function

' Returns True if Date1 is the last day of the quarter.
' Returns Null if Date1 is Null.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VIsDateUltimoQuarter( _
    ByVal Date1 As Variant) _
    As Variant

    Dim Ultimo  As Variant

    If IsDateExt(Date1) Then
        Ultimo = IsDateUltimoQuarter(CDate(Date1))
    Else
        Ultimo = Null
    End If
    
    VIsDateUltimoQuarter = Ultimo
    
End Function

' Returns True if Date1 is the last day of the week.
' Returns Null if Date1 is Null.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VIsDateUltimoWeek( _
    ByVal Date1 As Variant, _
    Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbSunday) _
    As Variant

    Dim Ultimo  As Variant

    If IsDateExt(Date1) Then
        Ultimo = IsDateUltimoWeek(CDate(Date1), FirstDayOfWeek)
    Else
        Ultimo = Null
    End If
    
    VIsDateUltimoWeek = Ultimo
    
End Function

' Returns True if Date1 is the last day of the year.
' Returns Null if Date1 is Null.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VIsDateUltimoYear( _
    ByVal Date1 As Variant) _
    As Variant

    Dim Ultimo  As Variant

    If IsDateExt(Date1) Then
        Ultimo = IsDateUltimoYear(CDate(Date1))
    Else
        Ultimo = Null
    End If
    
    VIsDateUltimoYear = Ultimo
    
End Function

' Returns True if the passed date is a weekend day ("off day") as
' specified by parameter WeekendType.
' Returns Null for if Date1 is passed Null or an invalid date expression.
'
' Default check is for the days of a long (Western) weekend, Saturday and Sunday.
'
'   2016-09-27. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VIsDateWeekend( _
    ByVal Date1 As Variant, _
    Optional ByVal WeekendType As DtWeekendType = DtWeekendType.dtLongWeekend) _
    As Variant
    
    Dim Result  As Variant

    If IsDate(Date1) Then
        Result = IsDateWeekend(CDate(Date1), WeekendType)
    Else
        Result = Null
    End If
    
    VIsDateWeekend = Result
    
End Function

' Returns the difference in full milleniums between Date1 and Date2.
' Returns Null if either Date1 or Date2 is Null.
'
' Calculates correctly for:
'   negative differences
'   leap years
'   dates of 29. February
'   date/time values with embedded time values
'   any date/time value of data type Date
'
' Optionally returns negative counts rounded down to provide a
' linear sequence of millenium counts.
' For a given Date1, if Date2 is decreased stepwise one millenium from
' returning a positive count to returning a negative count, one or two
' occurrences of count zero will be returned.
' If LinearSequence is False, the sequence will be:
'   3, 2, 1, 0,  0, -1, -2
' If LinearSequence is True, the sequence will be:
'   3, 2, 1, 0, -1, -2, -3
'
' If LinearSequence is False, reversing Date1 and Date2 will return
' results of same absolute Value, only the sign will change.
' This behaviour mimics that of Fix().
' If LinearSequence is True, reversing Date1 and Date2 will return
' results where the negative count is offset by -1.
' This behaviour mimics that of Int().
'
' DateAdd() is used for check for month end of February as it correctly
' returns Feb. 28th when adding a count of milleniums to dates of Feb. 29th
' when the resulting year is a common year.
'
' 2016-04-11. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VMilleniums( _
    ByVal Date1 As Variant, _
    ByVal Date2 As Variant, _
    Optional ByVal LinearSequence As Variant = Null) _
    As Variant
    
    Dim MilleniumCount  As Variant
    
    If IsDateExt(Date1) And IsDateExt(Date2) Then
        If IsNumeric(LinearSequence) Then
            LinearSequence = CBool(LinearSequence)
        Else
            LinearSequence = False
        End If
        MilleniumCount = Milleniums(CDate(Date1), CDate(Date2), LinearSequence)
    Else
        MilleniumCount = Null
    End If
    VMilleniums = MilleniumCount
  
End Function

' Returns the difference in full months between Date1 and Date2.
' Returns Null if either Date1 or Date2 is Null.
'
' Calculates correctly for:
'   negative differences
'   leap years
'   dates of 29. February
'   date/time values with embedded time values
'   any date/time value of data type Date
'
' Optionally returns negative counts rounded down to provide a
' linear sequence of month counts.
' For a given Date1, if Date2 is decreased stepwise one month from
' returning a positive count to returning a negative count, one or two
' occurrences of count zero will be returned.
' If LinearSequence is False, the sequence will be:
'   3, 2, 1, 0,  0, -1, -2
' If LinearSequence is True, the sequence will be:
'   3, 2, 1, 0, -1, -2, -3
'
' If LinearSequence is False, reversing Date1 and Date2 will return
' results of same absolute Value, only the sign will change.
' This behaviour mimics that of Fix().
' If LinearSequence is True, reversing Date1 and Date2 will return
' results where the negative count is offset by -1.
' This behaviour mimics that of Int().
'
' DateAdd() is, when adding a count of months to dates of 31th (29th),
' used for check for month end as it correctly returns the 30th (28th)
' when the resulting month has 30 or less days.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VMonths( _
    ByVal Date1 As Variant, _
    ByVal Date2 As Variant, _
    Optional ByVal LinearSequence As Variant = Null) _
    As Variant
    
    Dim MonthCount  As Variant
    
    If IsDateExt(Date1) And IsDateExt(Date2) Then
        If IsNumeric(LinearSequence) Then
            LinearSequence = CBool(LinearSequence)
        Else
            LinearSequence = False
        End If
        MonthCount = Months(CDate(Date1), CDate(Date2), LinearSequence)
    Else
        MonthCount = Null
    End If
    
    VMonths = MonthCount

End Function

' Returns the difference in full quarters between Date1 and Date2.
' Returns Null if either Date1 or Date2 is Null.
'
' Calculates correctly for:
'   negative differences
'   leap years
'   dates of 29. February
'   date/time values with embedded time values
'   any date/time value of data type Date
'
' Optionally returns negative counts rounded down to provide a
' linear sequence of quarter counts.
' For a given Date1, if Date2 is decreased stepwise one quarter from
' returning a positive count to returning a negative count, one or two
' occurrences of count zero will be returned.
' If LinearSequence is False, the sequence will be:
'   3, 2, 1, 0,  0, -1, -2
' If LinearSequence is True, the sequence will be:
'   3, 2, 1, 0, -1, -2, -3
'
' If LinearSequence is False, reversing Date1 and Date2 will return
' results of same absolute Value, only the sign will change.
' This behaviour mimics that of Fix().
' If LinearSequence is True, reversing Date1 and Date2 will return
' results where the negative count is offset by -1.
' This behaviour mimics that of Int().
'
' DateAdd() is, when adding a count of quarters to dates of 31th (29th),
' used for check for quarter end as it correctly returns the 30th (28th)
' when the resulting quarter has 30 or less days.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VQuarters( _
    ByVal Date1 As Variant, _
    ByVal Date2 As Variant, _
    Optional ByVal LinearSequence As Variant = Null) _
    As Variant
    
    Dim QuarterCount  As Variant
    
    If IsDateExt(Date1) And IsDateExt(Date2) Then
        If IsNumeric(LinearSequence) Then
            LinearSequence = CBool(LinearSequence)
        Else
            LinearSequence = False
        End If
        QuarterCount = Quarters(CDate(Date1), CDate(Date2), LinearSequence)
    Else
        QuarterCount = Null
    End If
    
    VQuarters = QuarterCount

End Function

' Returns the difference in full semimonths between Date1 and Date2.
' Returns Null if either Date1 or Date2 is Null.
'
' Calculates correctly for:
'   negative differences
'   leap years
'   dates of 29. February
'   date/time values with embedded time values
'   any date/time value of data type Date
'
' Optionally returns negative counts rounded down to provide a
' linear sequence of semimonth counts.
' For a given Date1, if Date2 is decreased stepwise one semimonth from
' returning a positive count to returning a negative count, one or two
' occurrences of count zero will be returned.
' If LinearSequence is False, the sequence will be:
'   3, 2, 1, 0,  0, -1, -2
' If LinearSequence is True, the sequence will be:
'   3, 2, 1, 0, -1, -2, -3
'
' If LinearSequence is False, reversing Date1 and Date2 will return
' results of same absolute Value, only the sign will change.
' This behaviour mimics that of Fix().
' If LinearSequence is True, reversing Date1 and Date2 will return
' results where the negative count is offset by -1.
' This behaviour mimics that of Int().
'
' DateAdd() is, when adding a count of semimonths to dates of 31th (29th),
' used for check for month end as it correctly returns the 30th (28th)
' when the resulting month of the semimonth has 30 or less days.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VSemimonths( _
    ByVal Date1 As Variant, _
    ByVal Date2 As Variant, _
    Optional ByVal LinearSequence As Variant = Null) _
    As Variant
    
    Dim SemimonthCount  As Variant
    
    If IsDateExt(Date1) And IsDateExt(Date2) Then
        If IsNumeric(LinearSequence) Then
            LinearSequence = CBool(LinearSequence)
        Else
            LinearSequence = False
        End If
        SemimonthCount = Semimonths(CDate(Date1), CDate(Date2), LinearSequence)
    Else
        SemimonthCount = Null
    End If
    
    VSemimonths = SemimonthCount

End Function

' Returns the difference in full semiyears between Date1 and Date2.
' Returns Null if either Date1 or Date2 is Null.
'
' Calculates correctly for:
'   negative differences
'   leap years
'   dates of 29. February
'   date/time values with embedded time values
'   any date/time value of data type Date
'
' Optionally returns negative counts rounded down to provide a
' linear sequence of semiyear counts.
' For a given Date1, if Date2 is decreased stepwise one semiyear from
' returning a positive count to returning a negative count, one or two
' occurrences of count zero will be returned.
' If LinearSequence is False, the sequence will be:
'   3, 2, 1, 0,  0, -1, -2
' If LinearSequence is True, the sequence will be:
'   3, 2, 1, 0, -1, -2, -3
'
' If LinearSequence is False, reversing Date1 and Date2 will return
' results of same absolute Value, only the sign will change.
' This behaviour mimics that of Fix().
' If LinearSequence is True, reversing Date1 and Date2 will return
' results where the negative count is offset by -1.
' This behaviour mimics that of Int().
'
' DateAdd() is, when adding a count of semiyears to dates of 31th (29th),
' used for check for semiyear end as it correctly returns the 30th (28th)
' when the resulting semiyear has 30 or less days.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VSemiyears( _
    ByVal Date1 As Variant, _
    ByVal Date2 As Variant, _
    Optional ByVal LinearSequence As Variant = Null) _
    As Variant
    
    Dim SemiyearCount       As Variant
    
    If IsDateExt(Date1) And IsDateExt(Date2) Then
        If IsNumeric(LinearSequence) Then
            LinearSequence = CBool(LinearSequence)
        Else
            LinearSequence = False
        End If
        SemiyearCount = Semiyears(CDate(Date1), CDate(Date2), LinearSequence)
    Else
        SemiyearCount = Null
    End If
    
    VSemiyears = SemiyearCount

End Function

' Returns the difference in full sextayears between Date1 and Date2.
' Returns Null if either Date1 or Date2 is Null.
'
' Calculates correctly for:
'   negative differences
'   leap years
'   dates of 29. February
'   date/time values with embedded time values
'   any date/time value of data type Date
'
' Optionally returns negative counts rounded down to provide a
' linear sequence of sextayear counts.
' For a given Date1, if Date2 is decreased stepwise one sextayear from
' returning a positive count to returning a negative count, one or two
' occurrences of count zero will be returned.
' If LinearSequence is False, the sequence will be:
'   3, 2, 1, 0,  0, -1, -2
' If LinearSequence is True, the sequence will be:
'   3, 2, 1, 0, -1, -2, -3
'
' If LinearSequence is False, reversing Date1 and Date2 will return
' results of same absolute Value, only the sign will change.
' This behaviour mimics that of Fix().
' If LinearSequence is True, reversing Date1 and Date2 will return
' results where the negative count is offset by -1.
' This behaviour mimics that of Int().
'
' DateAdd() is, when adding a count of sextayears to dates of 31th (29th),
' used for check for sextayear end as it correctly returns the 30th (28th)
' when the resulting sextayear has 30 or less days.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VSextayears( _
    ByVal Date1 As Variant, _
    ByVal Date2 As Variant, _
    Optional ByVal LinearSequence As Variant = Null) _
    As Variant
    
    Dim SextayearCount      As Variant
    
    If IsDateExt(Date1) And IsDateExt(Date2) Then
        If IsNumeric(LinearSequence) Then
            LinearSequence = CBool(LinearSequence)
        Else
            LinearSequence = False
        End If
        SextayearCount = Sextayears(CDate(Date1), CDate(Date2), LinearSequence)
    Else
        SextayearCount = Null
    End If
    
    VSextayears = SextayearCount

End Function

' Returns the difference in full tertiayears between Date1 and Date2.
' Returns Null if either Date1 or Date2 is Null.
'
' Calculates correctly for:
'   negative differences
'   leap years
'   dates of 29. February
'   date/time values with embedded time values
'   any date/time value of data type Date
'
' Optionally returns negative counts rounded down to provide a
' linear sequence of tertiayear counts.
' For a given Date1, if Date2 is decreased stepwise one tertiayear from
' returning a positive count to returning a negative count, one or two
' occurrences of count zero will be returned.
' If LinearSequence is False, the sequence will be:
'   3, 2, 1, 0,  0, -1, -2
' If LinearSequence is True, the sequence will be:
'   3, 2, 1, 0, -1, -2, -3
'
' If LinearSequence is False, reversing Date1 and Date2 will return
' results of same absolute Value, only the sign will change.
' This behaviour mimics that of Fix().
' If LinearSequence is True, reversing Date1 and Date2 will return
' results where the negative count is offset by -1.
' This behaviour mimics that of Int().
'
' DateAdd() is, when adding a count of tertiayears to dates of 31th (29th),
' used for check for tertiayear end as it correctly returns the 30th (28th)
' when the resulting tertiayear has 30 or less days.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VTertiayears( _
    ByVal Date1 As Variant, _
    ByVal Date2 As Variant, _
    Optional ByVal LinearSequence As Variant = Null) _
    As Variant
    
    Dim TertiayearCount      As Variant
    
    If IsDateExt(Date1) And IsDateExt(Date2) Then
        If IsNumeric(LinearSequence) Then
            LinearSequence = CBool(LinearSequence)
        Else
            LinearSequence = False
        End If
        TertiayearCount = Tertiayears(CDate(Date1), CDate(Date2), LinearSequence)
    Else
        TertiayearCount = Null
    End If
    
    VTertiayears = TertiayearCount

End Function

' Converts a time value - or a time part of a date value - to 12-hour time,
' effectively the time value as to the AM/PM format without the AM/PM label.
' Returns Null if Expression is Null or not a valid date/time.
'
' Examples:
'   VTimeToAm(#21:56:07#) -> #09:56:07#
'   VTimeToAm(#02:34:51#) -> #02:34:51#
'   VTimeToAm(Null) -> Null
'
'   2017-09-27. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VTimeToAm( _
    ByVal Expression As Variant) _
    As Variant
    
    Dim Result  As Variant
    
    If IsDateExt(Expression) Then
        Result = TimeToAm(CDate(Expression))
    Else
        Result = Null
    End If
    
    VTimeToAm = Result
    
End Function

' Returns the count of occurrences of the weekday of Date1 from Date1 to Date2 not including Date1.
' Returns Null if Date1 or Date2 or both are Null or an invalid value.
'
' 2016-02-14. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VWeekdayCount( _
    ByVal Date1 As Variant, _
    ByVal Date2 As Variant) _
    As Variant
    
    Dim Result  As Variant
    
    Result = Null
    
    On Error Resume Next
    Result = WeekdayCount(Date1, Date2)
    On Error GoTo 0

    VWeekdayCount = Result

End Function

' Returns the count of occurrences of a weekday in a month.
' If DayOfWeek is not passed, the weekday of the passed date is used.
' Returns Null if DateOfMonth or DayOfWeek is Null or an invalid value.
'
' Results:
'   If the weekday exists between the 29th and ultimo of the month,
'   the count is five, else four.
'
' 2017-01-21. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VWeekdayCountOfMonth( _
    ByVal DateOfMonth As Variant, _
    Optional ByVal DayOfWeek As Variant = VbDayOfWeek.vbUseSystemDayOfWeek) _
    As Variant
    
    Dim WeekdayCount    As Variant
    
    WeekdayCount = Null
    
    On Error Resume Next
    WeekdayCount = WeekdayCountOfMonth(DateOfMonth, DayOfWeek)
    On Error GoTo 0

    VWeekdayCountOfMonth = WeekdayCount

End Function

' Returns the count of occurrences of a weekday in a year.
' If Year is not passed, the current year is used.
' If DayOfWeek is not passed, the weekday of the current date is used.
' Returns Null if Year or DayOfWeek is Null or an invalid value.
'
' Results:
'   If primo of the year (1. January) falls on DayOfWeek,
'   the count is always 53.
'   If primo + 1 day of the year (2. January) falls on DayOfWeek,
'   the count is 53 in leap years.
'   In any other case, the count is 52.
'
' 2017-01-09. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VWeekdayCountOfYear( _
    Optional ByVal Year As Variant = 0, _
    Optional ByVal DayOfWeek As Variant = VbDayOfWeek.vbUseSystemDayOfWeek) _
    As Variant
    
    Dim WeekdayCount    As Variant
    
    WeekdayCount = Null
    
    On Error Resume Next
    WeekdayCount = WeekdayCountOfYear(Year, DayOfWeek)
    On Error GoTo 0
    
    VWeekdayCountOfYear = WeekdayCount

End Function

' Returns the signed count of a weekday between Date1 and Date2 not including Date1.
' Returns Null if Date1 or Date2 or both are Null or an invalid value.
'
' 2016-02-14. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VWeekdayDiff( _
    ByVal Date1 As Variant, _
    ByVal Date2 As Variant, _
    Optional ByVal DayOfWeek As VbDayOfWeek = vbUseSystemDayOfWeek) _
    As Variant
    
    Dim Result  As Variant
    
    Result = Null
    
    On Error Resume Next
    Result = WeekdayDiff(Date1, Date2, DayOfWeek)
    On Error GoTo 0
    
    VWeekdayDiff = Result
      
End Function

' Calculates the occurrence of the weekday of Date1 of the month of Date1.
' Returns this as an integer between 1 and 5.
' Returns Null if Date1 is Null or an invalid value.
'
' 2015-09-12, Cactus Data ApS, CPH.
'
Public Function VWeekdayOccurrenceOfMonth( _
    ByVal Date1 As Variant) _
    As Variant
  
    Dim WeekdayCount    As Variant
  
    If IsDateExt(Date1) Then
        WeekdayCount = WeekdayOccurrenceOfMonth(CDate(Date1))
    Else
        WeekdayCount = Null
    End If
  
    VWeekdayOccurrenceOfMonth = WeekdayCount
  
End Function

' Returns the ISO 8601 week count of years of two dates.
' Returns Null if Date1 or Date2 or both are Null or an invalid value.
'
' 2016-02-14. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VWeeks( _
    ByVal Date1 As Variant, _
    ByVal Date2 As Variant) _
    As Variant
    
    Dim Result  As Variant
    
    Result = Null
    
    On Error Resume Next
    Result = Weeks(Date1, Date2)
    On Error GoTo 0

    VWeeks = Result
    
End Function

' Returns the ISO 8601 week count of a year of a date.
' Returns Null if Date1 is Null.
'
' 2015-12-04. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VWeeksOfYearOfDate( _
    ByVal Date1 As Variant) _
    As Variant

    Dim Week    As Variant

    If IsDateExt(Date1) Then
        Week = WeeksOfYearOfDate(CDate(Date1))
    Else
        Week = Null
    End If
    
    VWeeksOfYearOfDate = Week

End Function

' Returns the ISO 8601 week count of a year span.
' Returns Null if Year1 or Year2 or both are Null or an invalid value.
'
' 2016-02-14. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VWeeksOfYears( _
    ByVal Year1 As Variant, _
    ByVal Year2 As Variant) _
    As Variant

    Dim Result  As Variant

    Result = Null
    
    On Error Resume Next
    Result = WeeksOfYears(Year1, Year2)
    On Error GoTo 0
    
    VWeeksOfYears = Result

End Function

' Returns the difference in full years between Date1 and Date2.
' Returns Null if either Date1 or Date2 is Null.
'
' Calculates correctly for:
'   negative differences
'   leap years
'   dates of 29. February
'   date/time values with embedded time values
'   any date/time value of data type Date
'
' Optionally returns negative counts rounded down to provide a
' linear sequence of year counts.
' For a given Date1, if Date2 is decreased stepwise one year from
' returning a positive count to returning a negative count, one or two
' occurrences of count zero will be returned.
' If LinearSequence is False, the sequence will be:
'   3, 2, 1, 0,  0, -1, -2
' If LinearSequence is True, the sequence will be:
'   3, 2, 1, 0, -1, -2, -3
'
' If LinearSequence is False, reversing Date1 and Date2 will return
' results of same absolute Value, only the sign will change.
' This behaviour mimics that of Fix().
' If LinearSequence is True, reversing Date1 and Date2 will return
' results where the negative count is offset by -1.
' This behaviour mimics that of Int().
'
' DateAdd() is used for check for month end of February as it correctly
' returns Feb. 28th when adding a count of years to dates of Feb. 29th
' when the resulting year is a common year.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VYears( _
    ByVal Date1 As Variant, _
    ByVal Date2 As Variant, _
    Optional ByVal LinearSequence As Variant = Null) _
    As Variant
    
    Dim YearCount   As Variant
    
    If IsDateExt(Date1) And IsDateExt(Date2) Then
        If IsNumeric(LinearSequence) Then
            LinearSequence = CBool(LinearSequence)
        Else
            LinearSequence = False
        End If
        YearCount = Years(CDate(Date1), CDate(Date2), LinearSequence)
    Else
        YearCount = Null
    End If
    
    VYears = YearCount

End Function


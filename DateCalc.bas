Attribute VB_Name = "DateCalc"
Option Explicit
'
' DateCalc
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
'   DateCore
'   DateFind
'   DateMsec
'

' Returns the difference in full years from DateOfBirth to current date,
' optionally to another date.
' Returns zero if AnotherDate is earlier than DateOfBirth.
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
Public Function Age( _
    ByVal DateOfBirth As Date, _
    Optional ByVal AnotherDate As Variant) _
    As Integer
    
    Dim ThisDate    As Date
    Dim Years       As Integer
      
    If IsDateExt(AnotherDate) Then
        ThisDate = CDate(AnotherDate)
    Else
        ThisDate = Date
    End If
    
    ' Find difference in calendar years.
    Years = DateDiff(IntervalSetting(DtInterval.dtYear), DateOfBirth, ThisDate)
    If Years > 0 Then
        ' Decrease by 1 if current date is earlier than birthday of current year
        ' using DateDiff to ignore a time portion of DateOfBirth.
        If DateDiff("d", ThisDate, DateAdd(IntervalSetting(DtInterval.dtYear), Years, DateOfBirth)) > 0 Then
            Years = Years - 1
        End If
    ElseIf Years < 0 Then
        Years = 0
    End If
    
    Age = Years
  
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
Public Function AgeLeague( _
    ByVal DateOfBirth As Date) _
    As Integer

    Const LeagueMonth   As Integer = 4
    Const LeagueDay     As Integer = 30
    
    Dim Age             As Integer
    Dim League          As Date
    
    League = DateSerial(Year(Date), LeagueMonth, LeagueDay)
    Age = Years(DateOfBirth, League)
    
    AgeLeague = Age
    
End Function

' Returns the difference in full months from DateOfBirth to current date,
' optionally to another date.
' Returns zero if AnotherDate is earlier than DateOfBirth.
'
' Calculates correctly for:
'   leap Months
'   dates of 29. February
'   date/time values with embedded time values
'   any date/time value of data type Date
'
' DateAdd() is, when adding a count of months to dates of 31th (29th),
' used for check for month end as it correctly returns the 30th (28th)
' when the resulting month has 30 or less days.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function AgeMonths( _
    ByVal DateOfBirth As Date, _
    Optional ByVal AnotherDate As Variant) _
    As Long
    
    Dim ThisDate    As Date
    Dim Months      As Long
      
    If IsDateExt(AnotherDate) Then
        ThisDate = CDate(AnotherDate)
    Else
        ThisDate = Date
    End If
    
    ' Find difference in calendar Months.
    Months = DateDiff("m", DateOfBirth, ThisDate)
    If Months > 0 Then
        ' Decrease by 1 if current date is earlier than birthday of current year
        ' using DateDiff to ignore a time portion of DateOfBirth.
        If DateDiff("d", ThisDate, DateAdd("m", Months, DateOfBirth)) > 0 Then
            Months = Months - 1
        End If
    ElseIf Months < 0 Then
        Months = 0
    End If
    
    AgeMonths = Months
  
End Function

' Returns the difference in full months from DateOfBirth to current date,
' optionally to another date.
' Returns by reference the difference in days.
' Returns zero if AnotherDate is earlier than DateOfBirth.
'
' Calculates correctly for:
'   leap Months
'   dates of 29. February
'   date/time values with embedded time values
'   any date/time value of data type Date
'
' DateAdd() is, when adding a count of months to dates of 31th (29th),
' used for check for month end as it correctly returns the 30th (28th)
' when the resulting month has 30 or less days.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function AgeMonthsDays( _
    ByVal DateOfBirth As Date, _
    Optional ByVal AnotherDate As Variant, _
    Optional ByRef Days As Integer) _
    As Long
    
    Dim ThisDate    As Date
    Dim Months      As Long
      
    If IsDateExt(AnotherDate) Then
        ThisDate = CDate(AnotherDate)
    Else
        ThisDate = Date
    End If
    
    ' Find difference in calendar Months.
    Months = DateDiff("m", DateOfBirth, ThisDate)
    If Months < 0 Then
        Months = 0
    Else
        If Months > 0 Then
            ' Decrease by 1 if current date is earlier than birthday of current year
            ' using DateDiff to ignore a time portion of DateOfBirth.
            If DateDiff("d", ThisDate, DateAdd("m", Months, DateOfBirth)) > 0 Then
                Months = Months - 1
            End If
        End If
        ' Find difference in days.
        Days = DateDiff("d", DateAdd("m", Months, DateOfBirth), ThisDate)
    End If
        
    AgeMonthsDays = Months
  
End Function

' Returns the rounded up difference in full years from DateOfBirth to
' current date, optionally to another date.
' Returns zero if AnotherDate is earlier than DateOfBirth.
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
Public Function AgeRoundUp( _
    ByVal DateOfBirth As Date, _
    Optional ByVal AnotherDate As Variant) _
    As Integer
    
    Dim ThisDate    As Date
    Dim Years       As Integer
      
    If IsDateExt(AnotherDate) Then
        ThisDate = CDate(AnotherDate)
    Else
        ThisDate = Date
    End If
    
    ' Find difference in calendar years.
    Years = DateDiff(IntervalSetting(DtInterval.dtYear), DateOfBirth, ThisDate)
    If Years >= 0 Then
        ' Increase by 1 if current date is earlier than birthday of current year
        ' using DateDiff to ignore a time portion of DateOfBirth.
        If DateDiff("d", ThisDate, DateAdd(IntervalSetting(DtInterval.dtYear), Years, DateOfBirth)) < 0 Then
            Years = Years + 1
        End If
    ElseIf Years < 0 Then
        Years = 0
    End If
    
    AgeRoundUp = Years
  
End Function

' Returns the difference in full centuries between Date1 and Date2.
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
Public Function Centuries( _
    ByVal Date1 As Date, _
    ByVal Date2 As Date, _
    Optional ByVal LinearSequence As Boolean) _
    As Long

    Dim CenturyCount    As Long
    
    CenturyCount = DateParts(IntervalSetting(DtInterval.dtCentury, True), Date1, Date2, LinearSequence)
    
    ' Return count of centuries as count of full century date parts.
    Centuries = CenturyCount
  
End Function

' Adds a decimal count of months to Date1.
'
' Will, on purpose, raise error 5 in case of excessive arguments.
'
' Examples:
'
'   Date1           n  n / 31   Result
'   --------------------------------------
'   2020-12-31      0  0.00     2020-12-31
'   2020-12-31      1  0.03     2021-01-01
'   2020-12-31      2  0.06     2021-01-02
'   2020-12-31      3  0.10     2021-01-03
'   2020-12-31      4  0.13     2021-01-04
'   2020-12-31      5  0.16     2021-01-05
'   2020-12-31      6  0.19     2021-01-06
'   2020-12-31      7  0.23     2021-01-07
'   2020-12-31      8  0.26     2021-01-08
'   2020-12-31      9  0.29     2021-01-09
'   2020-12-31     10  0.32     2021-01-10
'   2020-12-31     11  0.35     2021-01-11
'   2020-12-31     12  0.39     2021-01-12
'   2020-12-31     13  0.42     2021-01-13
'   2020-12-31     14  0.45     2021-01-14
'   2020-12-31     15  0.48     2021-01-15
'   2020-12-31     16  0.52     2021-01-16
'   2020-12-31     17  0.55     2021-01-17
'   2020-12-31     18  0.58     2021-01-18
'   2020-12-31     19  0.61     2021-01-19
'   2020-12-31     20  0.65     2021-01-20
'   2020-12-31     21  0.68     2021-01-21
'   2020-12-31     22  0.71     2021-01-22
'   2020-12-31     23  0.74     2021-01-23
'   2020-12-31     24  0.77     2021-01-24
'   2020-12-31     25  0.81     2021-01-25
'   2020-12-31     26  0.84     2021-01-26
'   2020-12-31     27  0.87     2021-01-27
'   2020-12-31     28  0.90     2021-01-28
'   2020-12-31     29  0.94     2021-01-29
'   2020-12-31     30  0.97     2021-01-30
'   2020-12-31     31  1.00     2021-01-31
'
'   DateAddMonths(#2020/01/31#, 0.99)   -> 2020-02-29
'   DateAddMonths(#2020/01/31#, 1)      -> 2020-02-29
'   DateAddMonths(#2020/02/01#, 0.99)   -> 2020-03-01
'   DateAddMonths(#2020/02/01#, 1)      -> 2020-03-02
'   DateAddMonths(#2020/02/29#, 0.25)   -> 2020-03-06
'   DateAddMonths(#2020/02/29#, 0.5)    -> 2020-03-13
'   DateAddMonths(#2020/02/29#, 0.75)   -> 2020-03-21
'   DateAddMonths(#2020/02/29#, 1)      -> 2020-03-30
'
'   DateAddMonths(#9999/12/31#, 0.5)    -> Error
'
' 2017-09-27. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateAddMonths( _
    ByVal Date1 As Date, _
    ByVal Months As Double) _
    As Date
    
    Dim Date2       As Date
    Dim FullMonths  As Integer
    Dim PartMonths  As Double
    Dim Days        As Integer
    
    FullMonths = Int(Months)
    PartMonths = Months - FullMonths - 1
    Date2 = DateAdd("m", 1 + FullMonths, Date1)
    Days = PartMonths * DaysInMonth(Date2)
    Date2 = DateAdd("d", Days, Date2)
    
    DateAddMonths = Date2
    
End Function

' Calculates Easter Sunday for year 1583 to 4099.
' Returns the date of Easter Sunday for the passed year.
' Easter Sunday is the Sunday following the Paschal Full Moon
' (PFM) date for the year.
'
' Argument Year must be a year between 1583 and 4099.
' Values outside this range will either return non-verified
' results or raise an error.
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
Public Function DateEasterSunday( _
    ByVal Year As Integer) _
    As Date
    
    Const MetonicCycle  As Integer = 19
    Const March         As Integer = 3
    
    ' Main values.
    Dim Century         As Integer
    Dim Decade          As Integer
    Dim Remainder       As Integer
    Dim Days            As Integer
    Dim Day             As Integer
    ' Correction values A to E.
    Dim ValueA          As Integer
    Dim ValueB          As Integer
    Dim ValueC          As Integer
    Dim ValueD          As Integer
    Dim ValueE          As Integer
    ' Result date.
    Dim EasterSunday    As Date
    
    ' Calculate the century and the decade and
    ' the remainder of the Metonic Cycle.
    Century = Year \ 100
    Decade = Year Mod 100
    Remainder = Year Mod MetonicCycle
    
    ' Calculate the Paschal Full Moon date.
    Days = (Century - 15) \ 2 + 202 - 11 * Remainder
    
    ' Adjust for selected centuries.
    Select Case Century
        Case 21, 24, 25, 27 To 32, 34, 35, 38
            Days = Days - 1
        Case 33, 36, 37, 39, 40
            Days = Days - 2
    End Select
    Days = Days Mod 30
    
    ' Calculate the correction values.
    ValueA = Days + 21
    ' Correct for the lunar month being slightly shorter than 30 days.
    If Days = 29 Or (Days = 28 And Remainder > 10) Then
        ValueA = ValueA - 1
    End If
    
    ' Find the next Sunday.
    ValueB = (ValueA - MetonicCycle) Mod 7
    
    ValueC = (40 - Century) Mod 4
    If ValueC = 3 Then
        ValueC = ValueC + 1
    End If
    If ValueC > 1 Then
        ValueC = ValueC + 1
    End If
    
    ValueD = (Decade + Decade \ 4) Mod 7
    
    ValueE = ((20 - ValueB - ValueC - ValueD) Mod 7) + 1
    
    ' Calculate the day.
    Day = ValueA + ValueE
    
    ' Build the date of Easter Sunday.
    EasterSunday = DateSerial(Year, March, Day)
    
    ' Return the date.
    DateEasterSunday = EasterSunday

End Function

' Returns the maximum date/time value of elements in a parameter array.
' If no elements of array Dates() are dates, the minimum value of Date is returned.
'
' Example:
'   DateMax(Null, "k", 0, 5, Date) -> Current date.
'
' 2016-02-14. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateMax( _
    ParamArray Dates() As Variant) _
    As Date

    Dim Element     As Variant
    Dim MaxFound    As Date
      
    MaxFound = MinDateValue
    
    For Each Element In Dates()
        If IsDateExt(Element) Then
            If VarType(Element) <> vbDate Then
                Element = CDate(Element)
            End If
            If Element > MaxFound Then
                MaxFound = Element
            End If
        End If
    Next
    
    DateMax = MaxFound
  
End Function

' Returns the minimum date/time value of elements in a parameter array.
' If no elements of array Dates() are dates, the maximum value of Date is returned.
'
' Example:
'   DateMax(Null, "k", 0, -5, Date) -> 1899-12-25.
'
' 2016-02-14. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateMin( _
    ParamArray Dates() As Variant) _
    As Date

    Dim Element     As Variant
    Dim MinFound    As Date
      
    MinFound = MaxDateValue
    
    For Each Element In Dates()
        If IsDateExt(Element) Then
            If VarType(Element) <> vbDate Then
                Element = CDate(Element)
            End If
            If Element < MinFound Then
                MinFound = Element
            End If
        End If
    Next
    
    DateMin = MinFound
  
End Function

' Returns the difference in full date parts between Date1 and Date2.
' The type of period is set by parameter Interval.
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
' 2019-11-06. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateParts( _
    ByVal Interval As String, _
    ByVal Date1 As Date, _
    ByVal Date2 As Date, _
    Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbSunday, _
    Optional ByVal FirstWeekOfYear As VbFirstWeekOfYear = vbFirstJan1, _
    Optional ByVal LinearSequence As Boolean) _
    As Double

    Dim Parts           As Double
    Dim Diff            As Double
    Dim ZeroInterval    As String
    
    If Not IsIntervalSetting(Interval, True) Then
        Err.Raise DtError.dtInvalidProcedureCallOrArgument
        Exit Function
    End If
    
    If IsIntervalDate(IntervalValue(Interval, True), True) Then
        ' Default and extended date parts for day and above.
        ZeroInterval = IntervalSetting(dtDay)
    ElseIf IsIntervalTime(IntervalValue(Interval, True), False) Then
        ' Default date parts under one day (time).
        ZeroInterval = IntervalSetting(dtSecond)
    Else
        ' Extended date parts under one day (time).
        ZeroInterval = IntervalSetting(dtMillisecond, True)
    End If
    
    Diff = DateDiffExt(ZeroInterval, Date1, Date2)
    If Diff = 0 Then
        ' The date values are equal in this context.
        ' Any date part of this type will be zero.
    Else
        ' Find difference in calendar date parts.
        Parts = DateDiffExt(Interval, Date1, Date2, FirstDayOfWeek, FirstWeekOfYear)
        ' For positive resp. negative intervals, check if the second date
        ' falls before, on, or after the crossing date for one period
        ' while at the same time correcting for February 29. of leap years.
        If Diff > 0 Then
            If DateDiffExt(ZeroInterval, DateAddExt(Interval, Parts, Date1), Date2, FirstDayOfWeek, FirstWeekOfYear) < 0 Then
                Parts = Parts - 1
            End If
        Else
            If DateDiffExt(ZeroInterval, DateAddExt(Interval, -Parts, Date2), Date1, FirstDayOfWeek, FirstWeekOfYear) < 0 Then
                Parts = Parts + 1
            End If
            ' Offset negative count of date parts to continuous sequence if requested.
            If LinearSequence = True Then
                Parts = Parts - 1
            End If
        End If
    End If
    
    ' Return count of date parts as count of full date parts.
    DateParts = Parts
  
End Function

' Rounds Date1 to hours and/or minutes and/or seconds as
' specified in parameters Hours, Minutes, and Seconds.
'
' Will accept any value within the range of data type Date:
'
'   From 100-01-01 00:00:00 to 9999-12-31 23:59:59
'
' In case the range is exceeded due to rounding, the native
' minimum or maximum value of Date will be returned.
'
' Examples:
'   DateRound(#9999-12-31 23:57:50#, 0, 5, 0)
'     returns: 9999-12-31 23:59:59
'   DateRound(#9999-12-31 23:57:10#, 0, 5, 0)
'     returns: 9999-12-31 23:55:00
'   DateRound(#9999-12-30 22:57:50#, 0, 5, 0)
'     returns: 9999-12-30 23:00:00
'   DateRound(#2015-02-28 12:37:50#, 0, 15, 0)
'     returns: 2015-02-28 12:45:00
'   DateRound(#2015-05-05 11:27:52#, 3, 0, 0)
'     returns: 2015-05-05 12:00:00
'   DateRound(#2015-05-25 11:11:13#, 0, 0, 2)
'     returns: 2015-05-25 11:11:14
'
'   Round to the tenth of a day:
'   DateRound(#2012-11-15 15:00:00#, 2, 24, 0)
'     returns: 2012-11-15 14:24:00
'
' 2017-01-21. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateRoundMid( _
    ByVal Date1 As Date, _
    Optional ByVal Hours As Integer, _
    Optional ByVal Minutes As Integer, _
    Optional ByVal Seconds As Integer) _
    As Date
    
    Dim Factor      As Date
    Dim Timespan    As Date
    Dim DateValue   As Date
    Dim TimeValue   As Date
    Dim Result      As Date
    
    On Error GoTo Err_DateRoundMid
    
    Factor = TimeSerial(Hours, Minutes, Seconds)
    If Factor <= 0 Then
        ' Round to the second.
        Factor = TimeSerial(0, 0, 1)
    End If
    
    ' Convert to timespan to round numerical negative dates (before 1899-12-30) correctly.
    Timespan = DateToTimespan(Date1)
    
    ' Get the time part only, to obtain precise rounding.
    ' TimeValue() cannot be used as it is buggy for the date of 9999-12-31.
    DateValue = Fix(Timespan)
    TimeValue = Timespan - DateValue
    
    ' Round the timepart.
    ' Apply CDec to prevent rounding errors from Doubles and allow large values.
    TimeValue = CDate(Int(CDec(TimeValue) / CDec(Factor) + 0.5) * Factor)
    
    ' Convert the date part and the rounded time part from timespan to date and time.
    Result = DateFromTimespan(DateValue + TimeValue)
    
Exit_DateRoundMid:
    DateRoundMid = Result
    Exit Function
    
Err_DateRoundMid:
    If Date1 < 0 Then
        Result = MinDateValue
    Else
        Result = MaxDateValue
    End If
    Resume Exit_DateRoundMid
    
End Function

' Returns for any date value a positive numeric value for Date1 including
' milliseconds that can be sorted on correctly even for negative date values.
'
' Return values span from:
'       0
'   for
'        100-01-01 00:00:00.000
'   to:
'       312,413,759,999,999
'   for:
'       9999-12-31 23:59:59.999
'
' 2016-09-17. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateSort( _
    ByVal Date1 As Date) _
    As Double

    Dim Result  As Double
  
    Result = MsecDiff(MinDateValue, Date1)
  
    DateSort = Result
    
End Function

' Returns the sum of a series of date values (timespans).
' Accepts also numeric values.
' Invalid entries are ignored.
'
' Examples:
'   SumValue = DateSum(#07:45#, #6:00#, Null, #1:00:44#, #4:10#)
'   SumValue -> 18:55:44
'
'   SumValue = DateSum(#08:30#, Null, 0.5, 1/86400)
'   SumValue -> 20:30:01
'
' 2019-11-04. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateSum( _
    ParamArray Expressions() As Variant) _
    As Date

    Dim Values()    As Date
    Dim Index       As Integer
    Dim Result      As Date
    
    ' Create array for JoinDate.
    ReDim Values(LBound(Expressions) To UBound(Expressions))
    
    ' Validate values for array for JoinDate.
    For Index = LBound(Expressions) To UBound(Expressions)
        ' Validate with IsDateExt that allow pure numerics.
        If IsDateExt(Expressions(Index)) Then
            Values(Index) = CDate(Expressions(Index))
        End If
    Next
    
    Result = JoinDate(Values())
    
    DateSum = Result

End Function

' Returns the count of days of the month of Date1.
'
' 2016-02-14. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DaysInMonth( _
    ByVal Date1 As Date) _
    As Integer
  
    Dim Days    As Integer
  
    If DateDiff(IntervalSetting(DtInterval.dtMonth), Date1, MaxDateValue) = 0 Then
        Days = MaxDayValue
    Else
        Days = Day(DateSerial(Year(Date1), Month(Date1) + 1, 0))
    End If
  
    DaysInMonth = Days
  
End Function

' Returns the count of days of the year of Date1.
'
' 2016-02-14. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DaysInYear( _
    ByVal Date1 As Date) _
    As Integer
  
    Dim Days    As Integer
  
    Days = DatePart(IntervalSetting(DtInterval.dtDayOfYear), DateSerial(Year(Date1), MaxMonthValue, MaxDayValue))
  
    DaysInYear = Days
  
End Function

' Returns the difference in full decades between Date1 and Date2.
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
Public Function Decades( _
    ByVal Date1 As Date, _
    ByVal Date2 As Date, _
    Optional ByVal LinearSequence As Boolean) _
    As Long

    Dim DecadeCount As Long
    
    DecadeCount = DateParts(IntervalSetting(DtInterval.dtDecade, True), Date1, Date2, LinearSequence)
    
    ' Return count of decades as count of full decade date parts.
    Decades = DecadeCount
  
End Function

' Returns the count of fortnights based on the ISO 8601 week count of years of two dates.
'
' 2016-02-14. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function Fortnights( _
    ByVal Date1 As Date, _
    ByVal Date2 As Date, _
    Optional ByVal LinearSequence As Boolean) _
    As Long

    Dim FortnightCount  As Long
    
    FortnightCount = DateParts(IntervalSetting(DtInterval.dtFortnight, True), Date1, Date2, LinearSequence)
    
    ' Return count of fortnights as count of full fortnight date parts.
    Fortnights = FortnightCount
    
End Function

' Returns the count of fortnights based on the ISO 8601 week count of the ISO year of a date.
'
' 2016-02-14. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function FortnightsOfYearOfDate( _
    ByVal Date1 As Date) _
    As Integer

    Dim IsoYear     As Integer
    Dim Result      As Integer

    Fortnight Date1, IsoYear
    Result = FortnightsOfYear(IsoYear)
    
    FortnightsOfYearOfDate = Result

End Function

' Returns the count of fortnights based on the ISO 8601 week count of a year span.
'
' 2016-02-14. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function FortnightsOfYears( _
    ByVal Year1 As Integer, _
    ByVal Year2 As Integer) _
    As Long

    Dim Year    As Integer
    Dim UpDown  As Integer
    Dim Result  As Long

    If Not (IsYear(Year1) And IsYear(Year2)) Then
        ' Don't accept years outside the range of Date.
        Err.Raise DtError.dtOverflow
        Exit Function
    End If
    
    If Year1 <= Year2 Then
        UpDown = 1
    Else
        UpDown = -1
    End If
    For Year = Year1 To Year2 Step UpDown
        Result = Result + FortnightsOfYear(Year) * UpDown
    Next
    
    FortnightsOfYears = Result

End Function

' Returns True if Date1 is of a leap year.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsDateLeapYear( _
    ByVal Date1 As Date) _
    As Boolean

    Dim LeapYear    As Boolean

    LeapYear = IsLeapYear(Year(Date1))
    
    IsDateLeapYear = LeapYear
    
End Function

' Returns True if Date1 is the first day of the month.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsDatePrimoMonth( _
    ByVal Date1 As Date) _
    As Boolean

    Dim Primo   As Boolean

    Primo = (Day(Date1) = MinDayValue)
    
    IsDatePrimoMonth = Primo
    
End Function

' Returns True if Date1 is the first day of the quarter.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsDatePrimoQuarter( _
    ByVal Date1 As Date) _
    As Boolean

    Dim Primo   As Boolean

    If Month(Date1) Mod DtIntervalMonths.dtQuarter = 0 Then
        Primo = IsDatePrimoMonth(Date1)
    End If
    
    IsDatePrimoQuarter = Primo
    
End Function

' Returns True if Date1 is the first day of the week.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsDatePrimoWeek( _
    ByVal Date1 As Date, _
    Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbSunday) _
    As Boolean

    Dim Primo   As Boolean

    If Weekday(Date1, FirstDayOfWeek) = FirstWeekday Then
        Primo = True
    End If
    
    IsDatePrimoWeek = Primo
    
End Function

' Returns True if Date1 is the first day of the year.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsDatePrimoYear( _
    ByVal Date1 As Date) _
    As Boolean

    Dim Primo   As Boolean

    If Month(Date1) = MinMonthValue And Day(Date1) = MinDayValue Then
        Primo = True
    End If
    
    IsDatePrimoYear = Primo
    
End Function

' Returns True if Date1 is the last day of the month.
' If Include2830 is True, also February 28th of leap years
' and the 30th of any month will be regarded as ultimo.
'
' 2019-10-04. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsDateUltimoMonth( _
    ByVal Date1 As Date, _
    Optional ByVal Include2830 As Boolean) _
    As Boolean

    Dim Ultimo  As Boolean
    Dim Day     As Integer

    Day = VBA.Day(Date1)
    
    If Day = MaxDayValue Then
        ' Will always be true, also for MaxDateValue.
        Ultimo = True
    ElseIf Date1 >= MaxDateValue Then
        ' Special case for 9999-12-31 to avoid an error in the next check.
        Ultimo = True
    ElseIf (VBA.Day(DateAdd("d", 1, Date1)) = 1) Then
        ' The next day belongs to the next month, thus Date1 is ultimo.
        Ultimo = True
    ElseIf Include2830 = True Then
        If Day = MaxDayValue - 1 Then
            ' Accept the 30th of any month as ultimo.
            Ultimo = True
        ElseIf Day = 28 Then
            ' Accept February 28th of leap years as ultimo.
            If Month(Date1) = 2 Then
                Ultimo = True
            End If
        End If
    End If
    
    IsDateUltimoMonth = Ultimo
    
End Function

' Returns True if Date1 is the last day of the quarter.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsDateUltimoQuarter( _
    ByVal Date1 As Date) _
    As Boolean

    Dim Ultimo  As Boolean

    If Month(Date1) Mod DtIntervalMonths.dtQuarter = 0 Then
        Ultimo = IsDateUltimoMonth(Date1)
    End If
    
    IsDateUltimoQuarter = Ultimo
    
End Function

' Returns True if Date1 is the last day of the week.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsDateUltimoWeek( _
    ByVal Date1 As Date, _
    Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbSunday) _
    As Boolean

    Dim Ultimo  As Boolean

    If Weekday(Date1, FirstDayOfWeek) = LastWeekday Then
        Ultimo = True
    End If
    
    IsDateUltimoWeek = Ultimo
    
End Function

' Returns True if Date1 is the last day of the year.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsDateUltimoYear( _
    ByVal Date1 As Date) _
    As Boolean

    Dim Ultimo  As Boolean

    If Month(Date1) = MaxMonthValue And Day(Date1) = MaxDayValue Then
        Ultimo = True
    End If
    
    IsDateUltimoYear = Ultimo
    
End Function

' Returns True if the passed date is a weekend day ("off day") as
' specified by parameter WeekendType.
'
' Default check is for the days of a long (Western) weekend, Saturday and Sunday.
'
' 2016-09-27. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsDateWeekend( _
    ByVal Date1 As Date, _
    Optional ByVal WeekendType As DtWeekendType = DtWeekendType.dtLongWeekend) _
    As Boolean
    
    Dim DaysOfWeek()    As VbDayOfWeek
    
    Dim Item            As Integer
    Dim Result          As Boolean

    ' Ignore error caused by passing an invalid value for WeekendType,
    ' thus WeekendDays returns an empty array.
    On Error GoTo IsDateWeekend_Error
    
    DaysOfWeek() = WeekendDays(WeekendType)
    For Item = LBound(DaysOfWeek()) To UBound(DaysOfWeek())
        If Weekday(Date1) = DaysOfWeek(Item) Then
            Result = True
            Exit For
        End If
    Next
    
    IsDateWeekend = Result

IsDateWeekend_Exit:
    Exit Function

IsDateWeekend_Error:
    ' Return False.
    Resume IsDateWeekend_Exit

End Function

' Returns the assembled (summed) parts of a series of date values.
'
' Support function for DateSum to sum an array of date values.
'
' Typical usage:
'   Assemble an array of date values created by SplitDate.
'
' Example:
'   SumValue = JoinDate(<array from DateSum: #07:45#, #6:00#, #1:00:44#, #4:10#>)
'   SumValue -> 18:55:44
'
' 2019-11-04. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function JoinDate( _
    ByRef Values() As Date) _
    As Date

    Dim Index       As Integer
    Dim Result      As Date
    
    ' Convert values to timespans and sum these.
    For Index = LBound(Values) To UBound(Values)
        Result = Result + CDec(DateToTimespan(Values(Index)))
    Next
    ' Convert the summed values to a date value.
    ConvTimespanToDate Result
    
    JoinDate = Result

End Function

' Returns the difference in full milleniums between Date1 and Date2.
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
Public Function Milleniums( _
    ByVal Date1 As Date, _
    ByVal Date2 As Date, _
    Optional ByVal LinearSequence As Boolean) _
    As Long

    Dim MilleniumCount  As Long
    
    MilleniumCount = DateParts(IntervalSetting(DtInterval.dtMillenium, True), Date1, Date2, LinearSequence)
    
    ' Return count of milleniums as count of full millenium date parts.
    Milleniums = MilleniumCount
  
End Function

' Returns the difference in full months between Date1 and Date2.
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
Public Function Months( _
    ByVal Date1 As Date, _
    ByVal Date2 As Date, _
    Optional ByVal LinearSequence As Boolean) _
    As Long

    Dim MonthCount  As Long
    
    MonthCount = DateParts(IntervalSetting(DtInterval.dtMonth), Date1, Date2, LinearSequence)
    
    ' Return count of months as count of full month date parts.
    Months = MonthCount
  
End Function

' Returns the difference in full quarters between Date1 and Date2.
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
Public Function Quarters( _
    ByVal Date1 As Date, _
    ByVal Date2 As Date, _
    Optional ByVal LinearSequence As Boolean) _
    As Long
    
    Dim QuarterCount    As Long
    
    QuarterCount = DateParts(IntervalSetting(DtInterval.dtQuarter), Date1, Date2, LinearSequence)
    
    ' Return count of quarters as count of full quarter date parts.
    Quarters = QuarterCount
    
End Function

' Returns the difference in full semimonths between Date1 and Date2.
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
' 2016-02-18. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function Semimonths( _
    ByVal Date1 As Date, _
    ByVal Date2 As Date, _
    Optional ByVal LinearSequence As Boolean) _
    As Long

    Dim SemimonthCount  As Long
    
    SemimonthCount = DateParts(IntervalSetting(DtInterval.dtSemimonth, True), Date1, Date2, LinearSequence)
    
    ' Return count of semimonths as count of full semimonth date parts.
    Semimonths = SemimonthCount
    
End Function

' Returns the difference in full semiyears between Date1 and Date2.
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
' when the resulting semiyear has 30 or less days in the last month.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function Semiyears( _
    ByVal Date1 As Date, _
    ByVal Date2 As Date, _
    Optional ByVal LinearSequence As Boolean) _
    As Long
    
    Dim SemiyearCount   As Long
    
    SemiyearCount = DateParts(IntervalSetting(DtInterval.dtSemiyear, True), Date1, Date2, LinearSequence)
    
    ' Return count of semiyears as count of full semiyear date parts.
    Semiyears = SemiyearCount
    
End Function

' Returns the difference in full sextayears between Date1 and Date2.
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
' when the resulting sextayear has 30 or less days in the last month.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function Sextayears( _
    ByVal Date1 As Date, _
    ByVal Date2 As Date, _
    Optional ByVal LinearSequence As Boolean) _
    As Long
    
    Dim SextayearCount   As Long
    
    SextayearCount = DateParts(IntervalSetting(DtInterval.dtSextayear, True), Date1, Date2, LinearSequence)
    
    ' Return count of semiyears as count of full semiyear date parts.
    Sextayears = SextayearCount
    
End Function

' Splits a date value in its date, time, and millesecond parts,
' and returns these as an array of date values.
'
' The returned array can be handled as needed, and then
' assembled with JoinDate.
'
' 2019-11-04. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function SplitDate( _
    ByVal Date1 As Date) _
    As Date()
    
    Const DateIndex     As Integer = 0
    Const TimeIndex     As Integer = 1
    Const MsecIndex     As Integer = 2
    
    Dim Parts(DateIndex To MsecIndex)   As Date
    
    ' Get millisecond part.
    Parts(MsecIndex) = MsecSerial(Millisecond(Date1))
    
    ' Strip milliseconds from Date1.
    RoundOffMilliseconds Date1
    
    ' Get date and time parts.
    Parts(TimeIndex) = TimeValue(Date1)
    Parts(DateIndex) = DateValue(Date1)
    
    SplitDate = Parts()
    
End Function

' Add/subtract a series of timespans.
'
' Typical usage:
'   Value = SumTimespans(Value1, Value2, ..., ValueN)
'
' Example:
'   Value = SumTimespans(#10:11#, #2:11#, -#3:22#, 3)
'   Value -> 1900-01-02 09:00:00.000
'
' 2017-04-28. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function SumTimespans( _
    ParamArray Timespans() As Variant) _
    As Date

    Dim Item    As Integer
    Dim Value   As Variant
    Dim Result  As Date
    
    Value = 0
    If Not IsEmpty(Timespans) Then
        For Item = LBound(Timespans) To UBound(Timespans)
            Value = Value + CDec(Timespans(Item))
        Next
    End If
    Result = DateFromTimespan(Value)
    
    SumTimespans = Result
  
End Function

' Returns the difference in full tertiayears between Date1 and Date2.
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
' DateAdd() is, when adding a count of tertiayear to dates of 31th (29th),
' used for check for tertiayear end as it correctly returns the 30th (28th)
' when the resulting tertiayear has 30 or less days in the last month.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function Tertiayears( _
    ByVal Date1 As Date, _
    ByVal Date2 As Date, _
    Optional ByVal LinearSequence As Boolean) _
    As Long
    
    Dim TertiayearCount As Long
    
    TertiayearCount = DateParts(IntervalSetting(DtInterval.dtTertiayear, True), Date1, Date2, LinearSequence)
    
    ' Return count of tertiayears as count of full tertiayear date parts.
    Tertiayears = TertiayearCount
    
End Function

' Converts a time value - or a time part of a date value - to 12-hour time,
' effectively the time value as to the AM/PM format without the AM/PM label.
'
' Examples:
'   TimeToAm(#21:56:07#) -> #09:56:07#
'   TimeToAm(#02:34:51#) -> #02:34:51#
'
' Alternative, though much slower:
'   Value12 = CDate(Format(Value24, "h:nn:ss ampm"))
'
' 2017-11-27. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function TimeToAm( _
    ByVal Time1 As Date) _
    As Date
    
    Dim Time2   As Date
    Dim Result  As Date
    
    Time2 = Time1 - Fix(Time1)
    Result = TimeSerial(Hour(Time2) Mod (HoursPerDay / 2), Minute(Time2), Second(Time2))
    
    TimeToAm = Result
    
End Function

' Returns the decimal count of months between Date1 and Date2.
'
' Rounds by default to two decimals, as more decimals has no meaning
' due to the varying count of days of a month.
' Optionally, don't round, by setting Round2 to False.
'
' 2017-01-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function TotalMonths( _
    ByVal Date1 As Date, _
    ByVal Date2 As Date, _
    Optional Round2 As Boolean = True) _
    As Double
    
    Dim Months      As Double
    Dim Part1       As Double
    Dim Part2       As Double
    Dim Fraction    As Double
    Dim Result      As Double
    
    Months = DateDiff("m", Date1, Date2)
    Part1 = (Day(Date1) - 1) / DaysInMonth(Date1)
    Part2 = (Day(Date2) - 1) / DaysInMonth(Date2)
    
    If Round2 = True Then
        ' Round to two decimals.
        Fraction = (-Part1 + Part2) * 100
        Result = Months + Int(Fraction + 0.5) / 100
    Else
        Result = Months - Part1 + Part2
    End If
    
    TotalMonths = Result
    
End Function

' Returns the decimal count of seconds between Date1 and Date2.
'
' Date1 and Date2 can be any valid date value between
'   100-1-1 00:00:00.000 and 9999-12-31 23:59:59.999
'
' 2019-11-06. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function TotalSeconds( _
    ByVal Date1 As Date, _
    ByVal Date2 As Date) _
    As Double
    
    Dim Result      As Double
    
    Result = DateDiffMsec("l", Date1, Date2)
    
    TotalSeconds = Result
    
End Function

' Returns the decimal count of years between Date1 and Date2.
'
' Rounds by default to three decimals, as more decimals has no meaning
' because of the leap years.
' Optionally, don't round, by setting Round3 to False.
'
' 2017-01-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function TotalYears( _
    ByVal Date1 As Date, _
    ByVal Date2 As Date, _
    Optional Round3 As Boolean = True) _
    As Double
    
    Dim Years       As Double
    Dim Part1       As Double
    Dim Part2       As Double
    Dim Fraction    As Double
    Dim Result      As Double
    
    Years = DateDiff(IntervalSetting(DtInterval.dtYear), Date1, Date2)
    Part1 = (DatePart(IntervalSetting(DtInterval.dtDayOfYear), Date1) - 1) / DaysInYear(Date1)
    Part2 = (DatePart(IntervalSetting(DtInterval.dtDayOfYear), Date2) - 1) / DaysInYear(Date2)
    
    If Round3 = True Then
        ' Round to three decimals.
        Fraction = (-Part1 + Part2) * 1000
        Result = Years + Int(Fraction + 0.5) / 1000
    Else
        Result = Years - Part1 + Part2
    End If
    
    TotalYears = Result
    
End Function

' Returns the count of occurrences of the weekday of Date1 from Date1 to Date2 not including Date1.
'
' 2016-02-14. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function WeekdayCount( _
    ByVal Date1 As Date, _
    ByVal Date2 As Date) _
    As Long
    
    Dim Result  As Long
    
    Result = DateDiff(IntervalSetting(DtInterval.dtWeekday), Date1, Date2)
    
    WeekdayCount = Result

End Function

' Returns the count of occurrences of a weekday in a month.
' If DayOfWeek is not passed, the weekday of the passed date is used.
'
' Results:
'   If the weekday exists between the 29th and ultimo of the month,
'   the count is five, else four.
'
' 2017-01-21. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function WeekdayCountOfMonth( _
    ByVal DateOfMonth As Date, _
    Optional ByVal DayOfWeek As VbDayOfWeek = VbDayOfWeek.vbUseSystemDayOfWeek) _
    As Integer
    
    Dim Primo           As Date
    Dim Ultimo          As Date
    Dim WeekdayCount    As Integer
    
    If Not IsWeekday(DayOfWeek) Then
        ' Don't accept invalid values for DayOfWeek.
        Err.Raise DtError.dtInvalidProcedureCallOrArgument
        Exit Function
    End If
    
    If DayOfWeek = VbDayOfWeek.vbUseSystemDayOfWeek Then
        ' Use weekday of DateOfMonth.
        DayOfWeek = Weekday(DateOfMonth)
    End If

    Primo = DateThisMonthPrimo(DateOfMonth)
    Ultimo = DateThisMonthUltimo(DateOfMonth)
    ' Cannot just count from ultimo of previous month as
    ' that would fail for January of year 100.
    WeekdayCount = DateDiff("ww", Primo, Ultimo, DayOfWeek)
    ' Include primo which DateDiff excludes.
    If Weekday(Primo) = DayOfWeek Then
        WeekdayCount = WeekdayCount + 1
    End If

    WeekdayCountOfMonth = WeekdayCount

End Function

' Returns the count of occurrences of a weekday in a year.
' If Year is not passed, the current year is used.
' If DayOfWeek is not passed, the weekday of the current date is used.
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
Public Function WeekdayCountOfYear( _
    Optional ByVal Year As Integer, _
    Optional ByVal DayOfWeek As VbDayOfWeek = VbDayOfWeek.vbUseSystemDayOfWeek) _
    As Integer
    
    Dim Primo           As Date
    Dim Secondo         As Date
    Dim WeekdayCount    As Integer
    
    If Year = 0 Then
        ' Use year of current date.
        Year = VBA.Year(Date)
    ElseIf Not IsYear(Year) Then
        ' Don't accept years outside the range of Date.
        Err.Raise DtError.dtOverflow
        Exit Function
    End If
    
    If Not IsWeekday(DayOfWeek) Then
        ' Don't accept invalid values for DayOfWeek.
        Err.Raise DtError.dtInvalidProcedureCallOrArgument
        Exit Function
    End If
    
    If DayOfWeek = VbDayOfWeek.vbUseSystemDayOfWeek Then
        ' Use weekday of today.
        DayOfWeek = Weekday(Date)
    End If
    
    Primo = DateYearPrimo(Year)
    WeekdayCount = MaxWeekValue - 1
    If Weekday(Primo) = DayOfWeek Then
        WeekdayCount = MaxWeekValue
    ElseIf IsLeapYear(Year) Then
        Secondo = DateAdd(IntervalSetting(DtInterval.dtDay), 1, Primo)
        If Weekday(Secondo) = DayOfWeek Then
            WeekdayCount = MaxWeekValue
        End If
    End If
    
    WeekdayCountOfYear = WeekdayCount

End Function

' Returns the signed count of a weekday between Date1 and Date2 not including Date1.
'
' 2016-02-14. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function WeekdayDiff( _
    ByVal Date1 As Date, _
    ByVal Date2 As Date, _
    Optional ByVal DayOfWeek As VbDayOfWeek = vbUseSystemDayOfWeek) _
    As Long
    
    Dim Result  As Long
    
    If Not IsWeekday(DayOfWeek) Then
        ' Don't accept invalid values for DayOfWeek.
        Err.Raise DtError.dtInvalidProcedureCallOrArgument
        Exit Function
    End If
    
    If DayOfWeek = VbDayOfWeek.vbUseSystemDayOfWeek Then
        ' Use weekday of today.
        DayOfWeek = Weekday(Date)
    End If
    
    Result = DateDiff(IntervalSetting(DtInterval.dtWeek), Date1, Date2, DayOfWeek)
    
    WeekdayDiff = Result
      
End Function

' Calculates the occurrence of the weekday of Date1 of the month of Date1.
' Returns this as an integer between 1 and 5.
'
' 2015-09-12, Cactus Data ApS, CPH.
'
Public Function WeekdayOccurrenceOfMonth( _
    ByVal Date1 As Date) _
    As Integer
  
    Dim WeekdayCount    As Integer
  
    WeekdayCount = -Int(-Day(Date1) / DaysPerWeek)
  
    WeekdayOccurrenceOfMonth = WeekdayCount
  
End Function

' Returns an array of weekend days ("off days") as
' specified by parameter WeekendType.
' Default is a long (Western) weekend of Saturday and Sunday.
'
' 2016-09-27. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function WeekendDays( _
    Optional ByVal WeekendType As DtWeekendType = DtWeekendType.dtLongWeekend) _
    As VbDayOfWeek()
    
    Dim DaysOfWeek()    As VbDayOfWeek
    
    Select Case WeekendType
        Case dtSunday To dtSaturday
            ReDim DaysOfWeek(0) As VbDayOfWeek
            DaysOfWeek(0) = WeekendType
        Case dtLongWeekend
            ReDim DaysOfWeek(1) As VbDayOfWeek
            DaysOfWeek(0) = vbSaturday
            DaysOfWeek(1) = vbSunday
        Case dtShortWeekend
            ReDim DaysOfWeek(0) As VbDayOfWeek
            DaysOfWeek(0) = vbSunday
        Case dtSabbath
            ReDim DaysOfWeek(0) As VbDayOfWeek
            DaysOfWeek(0) = vbSaturday
    End Select
    
    WeekendDays = DaysOfWeek()

End Function

' Returns an array of weekend days ("off days") as
' specified by parameter WeekendNumber.
' WeekendNumber should match a Weekend Number of Excel.
'
' Default is a long (Western) weekend of Saturday and Sunday.
' Invalid values will return a long (Western) weekend.
'
' 2020-04-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function WeekendDaysExcel( _
    Optional ByVal WeekendNumber As Integer) _
    As VbDayOfWeek()
    
    ' Excel Weekend Number ranges.
    Const MonoDualStep  As Integer = 10
    Const MaxMonoValue  As Integer = 7
    Const MinMonoValue  As Integer = 1
    Const MaxDualValue  As Integer = MaxMonoValue + MonoDualStep
    Const MinDualValue  As Integer = MinMonoValue + MonoDualStep
    
    Dim DaysOfWeek()    As VbDayOfWeek
    
    If WeekendNumber >= MinMonoValue And WeekendNumber <= MaxMonoValue Then
        ' Weekend of one day only.
        ReDim DaysOfWeek(1) As VbDayOfWeek
        DaysOfWeek(0) = ((WeekendNumber - 1) + DaysPerWeek - 1) Mod DaysPerWeek + 1
        DaysOfWeek(1) = WeekendNumber
    ElseIf WeekendNumber >= MinDualValue And WeekendNumber <= MaxDualValue Then
        ' Weekend of two days.
        ReDim DaysOfWeek(0) As VbDayOfWeek
        DaysOfWeek(0) = WeekendNumber - MonoDualStep
    Else
        ' Default Western weekend.
        ReDim DaysOfWeek(1) As VbDayOfWeek
        DaysOfWeek(0) = vbSaturday
        DaysOfWeek(1) = vbSunday
    End If
        
    WeekendDaysExcel = DaysOfWeek()

End Function

' Calculates the "weeknumber of the month" for a date.
' The value will be between 1 and 5.
'
' Numbering is similar to the ISO 8601 numbering having Monday
' as the first day of the week and the first week beginning
' with Thursday or later as week number 1.
' Thus, the first day of a month may belong to the last week
' of the previous month, having a week number of 4 or 5.
'
' 2020-09-23. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function WeekOfMonth( _
    ByVal Date1 As Date) _
    As Integer
    
    Dim ThursdayInWeek  As Date
    Dim FirstThursday   As Date
    Dim WeekNumber      As Integer
    
    ThursdayInWeek = DateWeekdayInWeek(Date1, vbThursday, vbMonday)
    FirstThursday = DateWeekdayInMonth(ThursdayInWeek, 1, vbThursday)
    WeekNumber = 1 + DateDiff("ww", FirstThursday, Date1, vbMonday)
    
    WeekOfMonth = WeekNumber
    
End Function

' Returns the ISO 8601 week count of years of two dates.
'
' 2016-02-14. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function Weeks( _
    ByVal Date1 As Date, _
    ByVal Date2 As Date, _
    Optional ByVal LinearSequence As Boolean) _
    As Long

    Dim WeekCount   As Long
    
    WeekCount = DateParts(IntervalSetting(DtInterval.dtWeek), Date1, Date2, LinearSequence)
    
    ' Return count of weeks as count of full week date parts.
    Weeks = WeekCount
    
End Function

' Returns the ISO 8601 week count of the ISO year of a date.
'
' 2015-12-04. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function WeeksOfYearOfDate( _
    ByVal Date1 As Date) _
    As Integer

    Dim IsoYear As Integer
    Dim Weeks   As Integer

    Week Date1, IsoYear
    Weeks = WeeksOfYear(IsoYear)
    
    WeeksOfYearOfDate = Weeks

End Function

' Returns the ISO 8601 week count of a year span.
'
' 2016-02-14. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function WeeksOfYears( _
    ByVal Year1 As Integer, _
    ByVal Year2 As Integer) _
    As Long

    Dim Year    As Integer
    Dim UpDown  As Integer
    Dim Result  As Long

    If Not (IsYear(Year1) And IsYear(Year2)) Then
        ' Don't accept years outside the range of Date.
        Err.Raise DtError.dtOverflow
        Exit Function
    End If
    
    If Year1 <= Year2 Then
        UpDown = 1
    Else
        UpDown = -1
    End If
    For Year = Year1 To Year2 Step UpDown
        Result = Result + WeeksOfYear(Year) * UpDown
    Next
    
    WeeksOfYears = Result

End Function

' Returns the relative value of the day count value of the year of the passed date.
' Optionally, return the reverse value - counting backwards from Dec. 31.
'
' Rounds to three decimals, as more decimals has no meaning because of the leap years.
'
' Examples, Reverse = False:
'   Date        DayOfYear   YearFraction    YearFraction * 365  Round(YearFraction * 365)
'   2015-01-01    1         0.003             1.095               1
'   2015-01-02    2         0.005             1.825               2
'   2015-07-01  182         0.499           182.135             182
'   2015-07-02  183         0.501           182.865             183
'   2015-12-31  365         1.000           365.000             365
'
'   Date        DayOfYear   YearFraction    YearFraction * 366  Round(YearFraction * 366)
'   2020-01-01    1         0.003             1.098               1
'   2020-01-02    2         0.005             1.830               2
'   2020-07-01  183         0.500           183.000             183
'   2020-12-30  365         0.997           364.902             365
'   2020-12-31  366         1.000           366.000             366
'
' Examples, Reverse = True:
'   Date        DayOfYear   YearFraction    YearFraction * 365  Round(YearFraction * 365)
'   2015-01-01    1         1.000           365.000             365
'   2015-12-31  365         0.003             1.095               1
'
'   Date        DayOfYear   YearFraction    YearFraction * 366  Round(YearFraction * 366)
'   2020-01-01    1         1.000           366.000             366
'   2020-12-31  366         0.003             1.098               1
'
' 2018-04-17. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function YearFraction( _
    ByVal Date1 As Date, _
    Optional ByVal Reverse As Boolean) _
    As Double
    
    Const Factor    As Integer = 1000
    
    Dim YearDays    As Integer
    Dim YearDay     As Integer
    Dim Fraction    As Double
    Dim Result      As Double
    
    YearDays = DaysInYear(Date1)
    YearDay = DatePart(IntervalSetting(DtInterval.dtDayOfYear), Date1)
    
    If Reverse = False Then
        Fraction = YearDay / YearDays
    Else
        Fraction = (1 + YearDays - YearDay) / YearDays
    End If
    
    ' Round to three decimals.
    Result = Int((Fraction * Factor) + 0.5) / Factor
    
    YearFraction = Result
    
End Function

' Returns the difference in full years between Date1 and Date2.
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
Public Function Years( _
    ByVal Date1 As Date, _
    ByVal Date2 As Date, _
    Optional ByVal LinearSequence As Boolean) _
    As Long

    Dim YearCount   As Long
    
    YearCount = DateParts(IntervalSetting(DtInterval.dtYear), Date1, Date2, LinearSequence)
    
    ' Return count of years as count of full year date parts.
    Years = YearCount
  
End Function


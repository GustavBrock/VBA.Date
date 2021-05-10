Attribute VB_Name = "DateCore"
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
'

' Supporting function for DateAddExt.
'
' Adds a positive or negative number of semimonths to the passed date.
' The handling of dates around ultimo/primo of semimonths is identical
' to that of VBA.DateAdd for a month.
'
' 2019-10-30. Gustav Brock, Cactus Data ApS, CPH.
'
Private Function AddSemimonths( _
    ByVal Number As Double, _
    ByVal Date1 As Date) _
    As Date
    
    Dim Parts           As Integer
    Dim MonthPart       As Integer
    Dim FullMonths      As Integer
    Dim Months          As Integer
    Dim ResultDate      As Date
    Dim ResultDay       As Integer
    
    ' Add/subtract full months and partial months.
    FullMonths = Number \ SemimonthsPerMonth
    Parts = Number Mod SemimonthsPerMonth
    
    Select Case Parts
        Case Is > 0
            ' Add semimonths.
            MonthPart = (Semimonth(Date1) + Number) Mod SemimonthsPerMonth
            Months = Month(Date1) + FullMonths + 1
            Select Case MonthPart
                Case 1
                    ' Result date shall belong to the first semimonth of the month.
                    ResultDate = DateSerial(Year(Date1), Months, Day(Date1) - (SemimonthsPerMonth - Parts) * DaysPerSemimonth)
                    If Semimonth(ResultDate) > (Semimonth(Date1) + Number) Mod SemimonthsPerYear Then
                        ' Adjust for ultimo.
                        ResultDate = DateSerial(Year(Date1), Months, Parts * DaysPerSemimonth)
                    End If
                Case Else
                    ' Result date shall belong to the last semimonth of the month.
                    ResultDate = DateSerial(Year(Date1), Months - 1, Day(Date1) + Parts * DaysPerSemimonth)
                    If Semimonth(ResultDate) Mod SemimonthsPerYear > (Semimonth(Date1) + Number) Mod SemimonthsPerYear Then
                        ' Adjust for ultimo February.
                        ResultDate = DateSerial(Year(Date1), Months, 0)
                    End If
            End Select
        Case Is < 0
            ' Subtract semimonths.
            MonthPart = ((Semimonth(Date1) + Number) Mod SemimonthsPerMonth + SemimonthsPerMonth) Mod SemimonthsPerMonth
            Months = Month(Date1) + FullMonths
            Select Case MonthPart
                Case 1
                    ' Result date shall belong to the first semimonth of a month.
                    If Day(Date1) <= SemimonthsPerMonth * DaysPerSemimonth Then
                        ResultDay = Day(Date1) + (Number Mod SemimonthsPerMonth) * DaysPerSemimonth
                    Else
                        ' Adjust for ultimo.
                        ResultDay = DaysPerSemimonth
                    End If
                    If IsYearMonth(Year(Date1), Months - 1) Then
                        ResultDate = DateSerial(Year(Date1), Months, ResultDay)
                    Else
                        ' No year below 100 as such years will be offset to 1900-2000 by DateSerial.
                        Err.Raise DtError.dtInvalidProcedureCallOrArgument
                    End If
                Case Else
                    ' Result date shall belong to the last semimonth of a month.
                    If IsYearMonth(Year(Date1), Months - 1) Then
                        ResultDate = DateSerial(Year(Date1), Months - 1, Day(Date1) + (SemimonthsPerMonth + Parts) * DaysPerSemimonth)
                        ' Correct for ultimo February.
                        If Semimonth(ResultDate) Mod SemimonthsPerMonth <> MonthPart Then
                            ' Move ResultDate to ultimo of the month.
                            ResultDate = DateSerial(Year(Date1), Months, 0)
                        End If
                    Else
                        ' No year below 100 as such years will be offset to 1900-2000 by DateSerial.
                        Err.Raise DtError.dtInvalidProcedureCallOrArgument
                    End If
            End Select
        Case Else
            ' Integer count of full months.
            ResultDate = DateAdd("m", FullMonths, Date1)
    End Select
    
    AddSemimonths = ResultDate
    
End Function

' Supporting function for DateAddExt.
'
' Adds a positive or negative number of tertiamonths to the passed date.
' The handling of dates around ultimo/primo of a tertiamonth is identical
' to that of VBA.DateAdd for a month.
'
' 2019-10-30. Gustav Brock, Cactus Data ApS, CPH.
'
Private Function AddTertiamonths( _
    ByVal Number As Double, _
    ByVal Date1 As Date) _
    As Date
    
    Dim Parts           As Integer
    Dim MonthPart       As Integer
    Dim FullMonths      As Integer
    Dim Months          As Integer
    Dim ResultDate      As Date
    Dim ResultDay       As Integer
    
    ' Add/subtract full months and partial months.
    FullMonths = Number \ TertiamonthsPerMonth
    Parts = Number Mod TertiamonthsPerMonth
    
    Select Case Parts
        Case Is > 0
            ' Add tertiamonths.
            MonthPart = (Tertiamonth(Date1) + Number) Mod TertiamonthsPerMonth
            Months = Month(Date1) + FullMonths + 1
            Select Case MonthPart
                Case 1
                    ' Result date shall belong to the first tertiamonth of the month.
                    ResultDate = DateSerial(Year(Date1), Months, Day(Date1) - (TertiamonthsPerMonth - Parts) * DaysPerTertiamonth)
                    If Tertiamonth(ResultDate) > (Tertiamonth(Date1) + Number) Mod TertiamonthsPerYear Then
                        ' Adjust for ultimo.
                        ResultDate = DateSerial(Year(Date1), Months, Parts * DaysPerTertiamonth)
                    End If
                Case 2
                    ' Result date shall belong to the middle tertiamonth of the month.
                    If Day(Date1) <= MonthPart * DaysPerTertiamonth Then
                        ' No ultimo month days.
                        ResultDate = DateSerial(Year(Date1), Months - 1, Day(Date1) + Parts * DaysPerTertiamonth)
                    Else
                        ResultDate = DateSerial(Year(Date1), Months, Day(Date1) - (TertiamonthsPerMonth - Parts) * DaysPerTertiamonth)
                        If Tertiamonth(ResultDate) > (Tertiamonth(Date1) + Number) Mod TertiamonthsPerYear Then
                            ' Adjust for ultimo.
                            ResultDate = DateSerial(Year(Date1), Months, Parts * DaysPerTertiamonth)
                        End If
                    End If
                Case Else
                    ' Result date shall belong to the last tertiamonth of the month.
                    ResultDate = DateSerial(Year(Date1), Months - 1, Day(Date1) + Parts * DaysPerTertiamonth)
                    If Tertiamonth(ResultDate) Mod TertiamonthsPerYear > (Tertiamonth(Date1) + Number) Mod TertiamonthsPerYear Then
                        ' Adjust for ultimo February.
                        ResultDate = DateSerial(Year(Date1), Months, 0)
                    End If
            End Select
        Case Is < 0
            ' Subtract tertiamonths.
            MonthPart = ((Tertiamonth(Date1) + Number) Mod TertiamonthsPerMonth + TertiamonthsPerMonth) Mod TertiamonthsPerMonth
            Months = Month(Date1) + FullMonths
            Select Case MonthPart
                Case 1
                    ' Result date shall belong to the first tertiamonth of a month.
                    If Day(Date1) <= TertiamonthsPerMonth * DaysPerTertiamonth Then
                        ResultDay = Day(Date1) + (Number Mod TertiamonthsPerMonth) * DaysPerTertiamonth
                    Else
                        ' Adjust for ultimo.
                        ResultDay = DaysPerTertiamonth
                    End If
                    If IsYearMonth(Year(Date1), Months - 1) Then
                        ResultDate = DateSerial(Year(Date1), Months, ResultDay)
                    Else
                        ' No year below 100 as such years will be offset to 1900-2000 by DateSerial.
                        Err.Raise DtError.dtInvalidProcedureCallOrArgument
                    End If
                Case 2
                    ' Result date shall belong to the middle tertiamonth of the month.
                    If Day(Date1) <= DaysPerTertiamonth Then
                        ' No ultimo month days.
                        ResultDate = DateSerial(Year(Date1), Months - 1, Day(Date1) + DaysPerTertiamonth)
                    Else
                        If IsYearMonth(Year(Date1), Months - 1) Then
                            ResultDate = DateSerial(Year(Date1), Months, Day(Date1) - DaysPerTertiamonth)
                            ' Correct for ultimo.
                            If Tertiamonth(ResultDate) Mod TertiamonthsPerMonth <> MonthPart Then
                                ' Move ResultDate to ultimo of the second tertiamonth of the month.
                                ResultDate = DateSerial(Year(Date1), Months, MonthPart * DaysPerTertiamonth)
                            End If
                        Else
                            ' No year below 100 as such years will be offset to 1900-2000 by DateSerial.
                            Err.Raise DtError.dtInvalidProcedureCallOrArgument
                        End If
                    End If
                Case Else
                    ' Result date shall belong to the last tertiamonth of a month.
                    If IsYearMonth(Year(Date1), Months - 1) Then
                        ResultDate = DateSerial(Year(Date1), Months - 1, Day(Date1) + (TertiamonthsPerMonth + Parts) * DaysPerTertiamonth)
                        ' Correct for ultimo February.
                        If Tertiamonth(ResultDate) Mod TertiamonthsPerMonth <> MonthPart Then
                            ' Move ResultDate to ultimo of the month.
                            ResultDate = DateSerial(Year(Date1), Months, 0)
                        End If
                    Else
                        ' No year below 100 as such years will be offset to 1900-2000 by DateSerial.
                        Err.Raise DtError.dtInvalidProcedureCallOrArgument
                    End If
            End Select
        Case Else
            ' Integer count of full months.
            ResultDate = DateAdd("m", FullMonths, Date1)
    End Select
    
    AddTertiamonths = ResultDate
    
End Function

' Returns the century of a date as the first year of the century.
' Value is 100 for the first century and 9900 for the last.
'
' 2016-04-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function Century( _
    ByVal Date1 As Date) _
    As Integer
    
    Dim Result  As Integer
    
    ' Find the century.
    Result = (Year(Date1) \ YearsPerCentury) * YearsPerCentury
    
    Century = Result
    
End Function

  
' Converts a date value by reference to a linear timespan value.
' Example:
'
'   Date     Time  Timespan      Date
'   19000101 0000  2             2
'
'   18991231 1800  1,75          1,75
'   18991231 1200  1,5           1,5
'   18991231 0600  1,25          1,25
'   18991231 0000  1             1
'
'   18991230 1800  0,75          0,75
'   18991230 1200  0,5           0,5
'   18991230 0600  0,25          0,25
'   18991230 0000  0             0
'
'   18991229 1800 -0,25         -1,75
'   18991229 1200 -0,5          -1,5
'   18991229 0600 -0,75         -1,25
'   18991229 0000 -1            -1
'
'   18991228 1800 -1,25         -2,75
'   18991228 1200 -1,5          -2,5
'   18991228 0600 -1,75         -2,25
'   18991228 0000 -2            -2
'
' 2015-12-15. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub ConvDateToTimespan( _
    ByRef Value As Date)
  
    Dim DatePart    As Double
    Dim TimePart    As Double
    
    If Value < 0 Then
        ' Get date (integer) part of Value shifted one day
        ' if a time part is present as -Int() rounds up.
        DatePart = -Int(-Value)
        ' Retrieve and reverse time (decimal) part.
        TimePart = DatePart - Value
        ' Assemble date and time part to return a timespan value.
        Value = CDate(DatePart + TimePart)
    Else
        ' Positive date values are identical to timespan values by design.
    End If

End Sub

' Converts a linear timespan value by reference to a date value.
' Example:
'
'   Date     Time  Timespan      Date
'   19000101 0000  2             2
'
'   18991231 1800  1,75          1,75
'   18991231 1200  1,5           1,5
'   18991231 0600  1,25          1,25
'   18991231 0000  1             1
'
'   18991230 1800  0,75          0,75
'   18991230 1200  0,5           0,5
'   18991230 0600  0,25          0,25
'   18991230 0000  0             0
'
'   18991229 1800 -0,25         -1,75
'   18991229 1200 -0,5          -1,5
'   18991229 0600 -0,75         -1,25
'   18991229 0000 -1            -1
'
'   18991228 1800 -1,25         -2,75
'   18991228 1200 -1,5          -2,5
'   18991228 0600 -1,75         -2,25
'   18991228 0000 -2            -2
'
' 2015-12-15. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub ConvTimespanToDate( _
    ByRef Value As Date)
   
    Dim DatePart    As Double
    Dim TimePart    As Double
  
    If Value < 0 Then
        ' Get date (integer) part of TimeSpan shifted one day
        ' if a time part is present as Int() rounds down.
        DatePart = Int(CDbl(Value))
        ' Retrieve and reverse time (decimal) part.
        TimePart = DatePart - Value
        ' Assemble the date and time parts to return a date value.
        Value = CDate(DatePart + TimePart)
    Else
        ' Positive timespan values are identical to date values by design.
    End If
  
End Sub

' An extended direct replacement for DateAdd, that can handle any
' value pair for Interval and Number that DateDiff and DateDiffExt
' can return.
'
' For native values of Interval, the maximum value for Number is given by
' seconds of the full range of data type Date:
'
'   SecondsMax = DateDiff("s", #1/1/100#, #12/31/9999 11:59:59 PM#) =>
'   SecondsMax = 312413759999
'
' However, the maximum value for Number, that DateAdd can accept, is
' only 2 ^ 31 - 1 or:
'
'   MaxNumber = 2147483647
'
' Thus, for larger values of Number, intervals are added in a loop,
' each adding the maximum number. Maximum loop count is:
'
'   Maximum loops = 312413759999 / 2147483647 = 145
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
' For extended values of Interval, the maximum value for Number is given by
' milliseconds of the full range of data type Date:
'
'   MillisecondsMax = DateDiffExt("f", #1/1/100#, MsecSerial(999, #12/31/9999 11:59:59 PM#)) =>
'   MillisecondsMax = 312413759999999
'
'   2019-10-19. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateAddExt( _
    ByVal Interval As String, _
    ByVal Number As Double, _
    ByVal Date1 As Date) _
    As Date

    Dim ResultDate  As Date
    Dim ExtValue    As Double
    
    If IsIntervalSetting(Interval, True) Then
        If Number = 0 Then
            ResultDate = Date1
        Else
            ' Only intervals of minutes, seconds, decimal seconds, or milliseconds
            ' can hold values larger than MaxNumber within the range of Date.
            ' If so, calculate as milliseconds.
            '
            ' Accept additions of semiyears, semimonths, and tertiamonths by custom calculation.
            ' Handle all other intervals directly by DateAdd.
            Select Case IntervalValue(Interval, True)
                Case DtInterval.dtSecond, DtInterval.dtMinute
                    If Number < MaxAddNumber Then
                        ' Default DateAdd operation.
                        ResultDate = DateAdd(Interval, Number, Date1)
                    Else
                        ' Calculate as milliseconds.
                        ResultDate = DateAddMsec(Interval, Number, Date1)
                    End If
                Case DtInterval.dtMillisecond, DtInterval.dtDecimalSecond
                    ResultDate = DateAddMsec(Interval, Number, Date1)
                
                Case DtInterval.dtDecade
                    ' Add decade.
                    ExtValue = DtIntervalMonths.dtDecade
                Case DtInterval.dtCentury
                    ' Add century.
                    ExtValue = DtIntervalMonths.dtCentury
                Case DtInterval.dtMillenium
                    ' Add millenium.
                    ExtValue = DtIntervalMonths.dtMillenium
                Case DtInterval.dtSemiyear
                    ' Add semiyear.
                    ExtValue = DtIntervalMonths.dtSemiyear
                Case DtInterval.dtTertiayear
                    ' Add thirdyear.
                    ExtValue = DtIntervalMonths.dtTertiayear
                Case DtInterval.dtSextayear
                    ' Add sixthyear.
                    ExtValue = DtIntervalMonths.dtSextayear
                
                Case DtInterval.dtSemimonth
                    ' Add semimonth.
                    ResultDate = AddSemimonths(Number, Date1)
                Case DtInterval.dtTertiamonth
                    ' Add tertiamonth.
                    ResultDate = AddTertiamonths(Number, Date1)
                Case DtInterval.dtFortnight
                    ' Add half of the fortnights by adding weeks.
                    ResultDate = DateAdd(IntervalSetting(DtInterval.dtWeek), Number, Date1)
                    ' One or more short fortnights may be included. Correct for these.
                    If Number > 0 Then
                        While DiffFortnights(Date1, ResultDate) < Number
                            ResultDate = DateAdd(IntervalSetting(DtInterval.dtWeek), 1, ResultDate)
                        Wend
                    Else
                        While DiffFortnights(Date1, ResultDate) > Number
                            ResultDate = DateAdd(IntervalSetting(DtInterval.dtWeek), -1, ResultDate)
                        Wend
                    End If
                Case DtInterval.dtFortnightday
                    ' Handle as Weekday.
                    ResultDate = DateAdd(IntervalSetting(DtInterval.dtWeekday), Number, Date1)
                Case Else
                    ' Default DateAdd operation.
                    ResultDate = DateAdd(Interval, Number, Date1)
            End Select
            If ExtValue > 0 Then
                ResultDate = DateAdd(IntervalSetting(DtInterval.dtMonth), Number * ExtValue, Date1)
            End If
        End If
    Else
        Err.Raise DtError.dtInvalidProcedureCallOrArgument
        Exit Function
    End If
    
    DateAddExt = ResultDate
    
End Function

' Returns the difference between two dates.
' Will also return the difference in extended intervals of
' half, third, and sixth years, and fortnights, and
' half and third months - as well as milliseconds and
' decimal seconds.
'
' Note, that optional parameters for week settings
' are ignored, as weeks always is handled according
' to ISO 8601.
'
' 2019-10-28. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateDiffExt( _
    ByVal Interval As String, _
    ByVal Date1 As Date, _
    ByVal Date2 As Date, _
    Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbMonday, _
    Optional ByVal FirstWeekOfYear As VbFirstWeekOfYear = vbFirstFourDays) _
    As Double
    
    Dim ExtValue    As Long
    Dim Result      As Double
    
    If IsIntervalSetting(Interval, False) Then
        Result = DateDiff(Interval, Date1, Date2, FirstDayOfWeek, FirstWeekOfYear)
    ElseIf IsIntervalSetting(Interval, True) Then
        Select Case IntervalValue(Interval, True)
            Case DtInterval.dtMillenium
                Result = (Millenium(Date2) - Millenium(Date1)) \ YearsPerMillenium
            Case DtInterval.dtCentury
                Result = (Century(Date2) - Century(Date1)) \ YearsPerCentury
            Case DtInterval.dtDecade
                Result = (Decade(Date2) - Decade(Date1)) \ YearsPerDecade
            Case DtInterval.dtDimidiae
                Result = Semiyear(Date2) - Semiyear(Date1)
                ExtValue = DtIntervalMonths.dtYear / DtIntervalMonths.dtSemiyear
            Case DtInterval.dtTertiayear
                Result = Tertiayear(Date2) - Tertiayear(Date1)
                ExtValue = DtIntervalMonths.dtYear / DtIntervalMonths.dtTertiayear
            Case DtInterval.dtSextayear
                Result = Sextayear(Date2) - Sextayear(Date1)
                ExtValue = DtIntervalMonths.dtYear / DtIntervalMonths.dtSextayear
            Case DtInterval.dtSemimonth
                Result = Semimonth(Date2) - Semimonth(Date1)
                ExtValue = DtIntervalMonths.dtYear * SemimonthsPerMonth
            Case DtInterval.dtTertiamonth
                Result = Tertiamonth(Date2) - Tertiamonth(Date1)
                ExtValue = DtIntervalMonths.dtYear * TertiamonthsPerMonth
                
            Case DtInterval.dtFortnight
                ' Implemented by external function.
                Result = DiffFortnights(Date1, Date2)
            Case DtInterval.dtFortnightday
                ' Implemented by external function.
                Result = FortnightdayCount(Date1, Date2)
                
            Case DtInterval.dtMillisecond
                ' Implemented by external function.
                Result = DateDiffMsec(IntervalSetting(DtInterval.dtMillisecond, True), Date1, Date2)
            Case DtInterval.dtDecimalSecond
                ' Implemented by external function.
                Result = DateDiffMsec(IntervalSetting(DtInterval.dtDecimalSecond, True), Date1, Date2)
        End Select
        If ExtValue > 0 Then
            Result = Result + DateDiff(IntervalSetting(DtInterval.dtYear), Date1, Date2) * ExtValue
        End If
    Else
        Err.Raise DtError.dtInvalidProcedureCallOrArgument
        Exit Function
    End If
        
    DateDiffExt = Result
    
End Function

' Returns the date of the earliest day of January that
' always will fall in the first ISO 8601 week of a year.
' If Year is not passed, the current year is used.
'
' 2018-07-05. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateFirstWeekYear( _
    Optional ByVal Year As Integer) _
    As Date

    ' The earliest day that always falls in the first week of the year.
    Const MonthOfFirstWeek  As Integer = 1
    Const DayOfFirstWeek    As Integer = 4
    
    Dim WeekDate            As Date

    If Year = 0 Then
        ' Use year of current date.
        Year = VBA.Year(Date)
    End If
    
    WeekDate = DateSerial(Year, MonthOfFirstWeek, DayOfFirstWeek)
    
    DateFirstWeekYear = WeekDate

End Function

' Converts a timespan value to a date value.
' Useful only for result date values prior to 1899-12-30 as
' these have a negative numeric value.
'
' 2015-12-15. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateFromTimespan( _
    ByVal Value As Date) _
    As Date
  
    ConvTimespanToDate Value
  
    DateFromTimespan = Value
  
End Function

' Returns the first or earliest date and/or time of an interval of the date
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
'   2017-11-27. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateIntervalPrimo( _
    ByVal Interval As String, _
    ByVal Number As Double, _
    ByVal Date1 As Date, _
    Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbSunday, _
    Optional AcceptMilliseconds As Boolean) _
    As Date
    
    Dim Value   As DtInterval
    Dim Time1   As Date
    Dim Month   As Integer
    Dim Months  As Integer
    Dim Day     As Integer
    Dim Minute  As Integer
    Dim Second  As Integer
    Dim Result  As Date
    
    If IsIntervalSetting(Interval, True) Then
        Value = IntervalValue(Interval, True)
    Else
        Err.Raise DtError.dtInvalidProcedureCallOrArgument
        Exit Function
    End If
    
    If Number <> 0 Then
        ' Offset date.
        Date1 = DateAddExt(Interval, Number, Date1)
    End If
    
    If Value = DtInterval.dtDay _
            Or Value = DtInterval.dtDayOfYear _
            Or Value = DtInterval.dtWeekday _
            Or Value = DtInterval.dtFortnightday Then
        ' Remove a time part.
        Result = Fix(Date1)
    ElseIf IsIntervalTime(Value, True) Then
        ' Clean a time part.
        Time1 = Date1 - Fix(Date1)
        Select Case Value
            Case DtInterval.dtHour
                ' Set hours to Hour:00:00.000
            Case DtInterval.dtMinute
                ' Set hours to Hour:Minute:00.000
                Minute = VBA.Minute(Time1)
            Case DtInterval.dtSecond
                ' Set hours to Hour:Minute:Seconds.000
                Minute = VBA.Minute(Time1)
                Second = VBA.Second(Time1)
            Case DtInterval.dtMillisecond
                If AcceptMilliseconds = False Then
                    ' Return seconds with no milliseconds.
                Else
                    ' Return seconds with rounded down milliseconds.
                    ' Not yet implemented.
                End If
                Minute = VBA.Minute(Time1)
                Second = VBA.Second(Time1)
        End Select
        Result = DateTimeSerial(VBA.Year(Date1), VBA.Month(Date1), VBA.Day(Date1), VBA.Hour(Date1), Minute, Second)
    ElseIf Value = DtInterval.dtWeek Then
        ' Return first weekday with no time part.
        Result = DateAdd(IntervalSetting(DtInterval.dtDay), FirstWeekday - Weekday(Date1, FirstDayOfWeek), Fix(Date1))
    ElseIf Value = DtInterval.dtFortnight Then
        ' Return first fortnight day with no time part. Will always be a Monday.
        Result = DateAdd(IntervalSetting(DtInterval.dtDay), FirstFortnightday - Fortnightday(Date1, vbMonday), Fix(Date1))
    ElseIf Value = DtInterval.dtSemimonth Then
        If VBA.Day(Date1) <= DaysPerSemimonth Then
            Day = FirstSemimonthday
        Else
            Day = SecondSemimonthday
        End If
        Result = DateSerial(Year(Date1), VBA.Month(Date1), Day)
    ElseIf IsIntervalDate(Value, True) Then
        Select Case Value
            Case DtInterval.dtDay, DtInterval.dtDayOfYear, DtInterval.dtWeekday, DtInterval.dtFortnightday
                ' Handled as time above.
            Case DtInterval.dtWeek
                ' Handled as week above.
            Case DtInterval.dtFortnight
                ' Handled as fortnight above.
            Case DtInterval.dtSemimonth
                ' Handled as semimonth above.
            Case DtInterval.dtMonth
                Months = DtIntervalMonths.dtMonth
            Case DtInterval.dtSextayear
                Months = DtIntervalMonths.dtSextayear
            Case DtInterval.dtQuarter
                Months = DtIntervalMonths.dtQuarter
            Case DtInterval.dtTertiayear
                Months = DtIntervalMonths.dtTertiayear
            Case DtInterval.dtDimidiae
                Months = DtIntervalMonths.dtDimidiae
            Case DtInterval.dtSemiyear
                Months = DtIntervalMonths.dtSemiyear
            Case DtInterval.dtYear
                ' Offset has been set as years.
                Months = 0
        End Select
        Month = MinMonthValue + (DatePartExt(IntervalSetting(Value, True), Date1) - 1) * Months
        Result = DateSerial(Year(Date1), Month, MinDayValue)
    End If
    
    DateIntervalPrimo = Result

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
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateIntervalUltimo( _
    ByVal Interval As String, _
    ByVal Number As Double, _
    ByVal Date1 As Date, _
    Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbSunday, _
    Optional AcceptMilliseconds As Boolean) _
    As Date
    
    Dim Value   As DtInterval
    Dim Time1   As Date
    Dim Month   As Integer
    Dim Months  As Integer
    Dim Day     As Integer
    Dim Minute  As Integer
    Dim Second  As Integer
    Dim Result  As Date
    
    If IsIntervalSetting(Interval, True) Then
        Value = IntervalValue(Interval, True)
    Else
        Err.Raise DtError.dtInvalidProcedureCallOrArgument
        Exit Function
    End If
    
    If Number <> 0 Then
        ' Offset date.
        Date1 = DateAddExt(Interval, Number, Date1)
    End If
    
    If Value = DtInterval.dtDay _
            Or Value = DtInterval.dtDayOfYear _
            Or Value = DtInterval.dtWeekday _
            Or Value = DtInterval.dtFortnightday Then
        ' Replace a time part with last second of the day.
        Result = DateAdd(IntervalSetting(DtInterval.dtSecond), -1, DateAdd(IntervalSetting(DtInterval.dtDay), 1, Fix(Date1)))
    ElseIf IsIntervalTime(Value, True) Then
        ' Clean a time part.
        Time1 = Date1 - Fix(Date1)
        Select Case Value
            Case DtInterval.dtHour
                ' Set hours to Hour:59:59.000 or .999
                Minute = VBA.Minute(MaxTimeValue)
                Second = VBA.Second(MaxTimeValue)
            Case DtInterval.dtMinute
                ' Set hours to Hour:Minute:59.000 or .999
                Minute = VBA.Minute(Time1)
                Second = VBA.Second(MaxTimeValue)
            Case DtInterval.dtSecond
                ' Set hours to Hour:Minute:Seconds.000 or .999
                Minute = VBA.Minute(Time1)
                Second = VBA.Second(Time1)
            Case DtInterval.dtMillisecond
                ' Set hours to Hour:Minute:Seconds and milliseconds.
                ' Not yet implemented.
                Minute = VBA.Minute(Time1)
                Second = VBA.Second(Time1)
        End Select
        If AcceptMilliseconds = False Then
            ' Return seconds.
        Else
            ' Return seconds and rounded up milliseconds.
            ' Not yet implemented.
        End If
        Result = DateTimeSerial(VBA.Year(Date1), VBA.Month(Date1), VBA.Day(Date1), VBA.Hour(Date1), Minute, Second)
    ElseIf Value = DtInterval.dtWeek Then
        ' Return last weekday with no time part.
        Result = DateAdd(IntervalSetting(DtInterval.dtDay), LastWeekday - Weekday(Date1, FirstDayOfWeek), Fix(Date1))
    ElseIf Value = DtInterval.dtFortnight Then
        ' Return last fortnight day with no time part. Will always be a Sunday.
        Result = DateAdd(IntervalSetting(DtInterval.dtDay), LastFortnightday - Fortnightday(Date1, vbMonday), Fix(Date1))
        If Fortnight(Date1) = MaxFortnightValue Then
            ' Shorted fortnight. Reduce by one week.
            Result = DateAdd(IntervalSetting(DtInterval.dtWeek), -1, Result)
        End If
    ElseIf Value = DtInterval.dtSemimonth Then
        Month = VBA.Month(Date1)
        If VBA.Day(Date1) <= DaysPerSemimonth Then
            Day = SecondSemimonthday
            Result = DateSerial(VBA.Year(Date1), Month, Day - 1)
        Else
            If DateDiff(IntervalSetting(DtInterval.dtMonth), Date1, MaxDateValue) = 0 Then
                ' Last month of Date range.
                Result = MaxDateValue
            Else
                Day = FirstSemimonthday
                Result = DateSerial(VBA.Year(Date1), Month + 1, Day - 1)
            End If
        End If
    ElseIf IsIntervalDate(Value, True) Then
        Select Case Value
            Case DtInterval.dtDay, DtInterval.dtDayOfYear, DtInterval.dtWeekday, DtInterval.dtFortnightday
                ' Handled as time above.
            Case DtInterval.dtWeek
                ' Handled as week above.
            Case DtInterval.dtFortnight
                ' Handled as fortnight above.
            Case DtInterval.dtSemimonth
                ' Handled as semimonth above.
            Case DtInterval.dtMonth
                Months = DtIntervalMonths.dtMonth
            Case DtInterval.dtSextayear
                Months = DtIntervalMonths.dtSextayear
            Case DtInterval.dtQuarter
                Months = DtIntervalMonths.dtQuarter
            Case DtInterval.dtTertiayear
                Months = DtIntervalMonths.dtTertiayear
            Case DtInterval.dtDimidiae
                Months = DtIntervalMonths.dtDimidiae
            Case DtInterval.dtSemiyear
                Months = DtIntervalMonths.dtSemiyear
            Case DtInterval.dtYear
                ' Offset has been set as years.
                Months = 0
        End Select
        Select Case DateDiff(IntervalSetting(DtInterval.dtMonth), Date1, MaxDateValue)
            Case Is < Months
                Result = MaxDateValue
            Case Else
                Month = DatePartExt(IntervalSetting(Value, True), Date1) * Months
                Result = DateSerial(Year(Date1), Month + 1, MinDayValue - 1)
        End Select
    End If
    
    DateIntervalUltimo = Result

End Function

' Returns the date of the latest day of December that
' always will fall in the last ISO 8601 week of a year.
' If Year is not passed, the current year is used.
'
' 2018-07-05. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateLastWeekYear( _
    Optional ByVal Year As Integer) _
    As Date

    ' The latest day that always falls in the last week of the year.
    Const MonthOfLastWeek   As Integer = 12
    Const DayOfLastWeek     As Integer = 28
    
    Dim WeekDate            As Date

    If Year = 0 Then
        ' Use year of current date.
        Year = VBA.Year(Date)
    End If
    
    WeekDate = DateSerial(Year, MonthOfLastWeek, DayOfLastWeek)
    
    DateLastWeekYear = WeekDate

End Function

' Returns the date part of a date.
' Also returns the extended period types of a date, including decimal seconds.
' Return value is Double to allow decimal seconds to be returned.
'
' Will return correct reading of seconds of the last day of Date (9999-12-31).
' See function DateTest.TheLastSeconds for full information.
'
' Will return the correct ISO 8601 week number with parameters:
'
'    FirstDayOfWeek = vbMonday
'    FirstWeekOfYear = vbFirstFourDays
'
' 2019-10-28. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DatePartExt( _
    ByVal Interval As String, _
    ByVal Date1 As Date, _
    Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbSunday, _
    Optional ByVal FirstWeekOfYear As VbFirstWeekOfYear = vbFirstJan1) _
    As Double
    
    Dim Result  As Double
    
    If IsIntervalSetting(Interval, False) Then
        ' Default date part intervals.
        Select Case IntervalValue(Interval, False)
            Case DtInterval.dtWeek
                If FirstDayOfWeek = vbMonday And FirstWeekOfYear = vbFirstFourDays Then
                    ' Return ISO 8601 week number.
                    Result = Week(Date1)
                Else
                    ' Return non-standard week number.
                    Result = DatePart(Interval, Date1, FirstDayOfWeek, FirstWeekOfYear)
                End If
            Case DtInterval.dtSecond
                If DateDiff(IntervalSetting(DtInterval.dtDay), Date1, MaxDateValue) = 0 Then
                    ' Native reading of seconds is buggy for the last date.
                    ' Obtain correct reading by removing the date part, thus
                    ' reading the seconds from the time part only.
                    Result = Second(Date1 - Fix(Date1))
                Else
                    Result = DatePart(Interval, Date1)
                End If
            Case Else
                Result = DatePart(Interval, Date1)
        End Select
    ElseIf IsIntervalSetting(Interval, True) Then
        ' Extended date part intervals.
        Select Case IntervalValue(Interval, True)
            Case DtInterval.dtMillenium
                Result = Millenium(Date1)
            Case DtInterval.dtCentury
                Result = Century(Date1)
            Case DtInterval.dtDecade
                Result = Decade(Date1)
            Case DtInterval.dtDimidiae, DtInterval.dtSemiyear
                Result = Semiyear(Date1)
            Case DtInterval.dtTertiayear
                Result = Tertiayear(Date1)
            Case DtInterval.dtSextayear
                Result = Sextayear(Date1)
            Case DtInterval.dtSemimonth
                Result = Semimonth(Date1)
            Case DtInterval.dtTertiamonth
                Result = Tertiamonth(Date1)
            
            Case DtInterval.dtFortnight
                If FirstDayOfWeek = vbMonday And FirstWeekOfYear = vbFirstFourDays Then
                    ' Return fortnight based on ISO 8601 week number.
                    Result = Fortnight(Date1)
                Else
                    ' Return fortnight based on non-standard number.
                    Result = -Int(-DatePart(IntervalSetting(DtInterval.dtWeek), Date1, FirstDayOfWeek, FirstWeekOfYear) / 2)
                End If
            Case DtInterval.dtFortnightday
                Result = Fortnightday(Date1, FirstDayOfWeek)
            
            Case DtInterval.dtDecimalSecond
                Result = DecimalSecond(Date1)
            Case DtInterval.dtMillisecond
                Result = Millisecond(Date1)
        End Select
    Else
        Err.Raise DtError.dtInvalidProcedureCallOrArgument
        Exit Function
    End If
        
    DatePartExt = Result
    
End Function

' Generates a random date/time - optionally within the range of LowerDate and/or UpperDate.
' Optionally, return value can be set to include date and/or time and/or milliseconds.
'
' 2017-03-12. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateRandom( _
    Optional ByVal UpperDate As Date = MaxDateValue, _
    Optional ByVal LowerDate As Date = MinDateValue, _
    Optional ByVal DatePart As Boolean = True, _
    Optional ByVal TimePart As Boolean = True, _
    Optional ByVal MillisecondPart As Boolean = False) _
    As Date
    
    Dim DateValue   As Date
    Dim TimeValue   As Date
    Dim MsecValue   As Date
    
    ' Shuffle the start position of the sequence of Rnd.
    Randomize
    
    ' If no part is selected, select date and time.
    If (DatePart Or TimePart Or MillisecondPart) = False Then
        DatePart = True
        TimePart = True
    End If
    If DatePart = True Then
        ' Remove time parts from UpperDate and LowerDate as well from the result value.
        ' Add 1 to include LowerDate as a possible return value.
        DateValue = CDate(Int((Int(UpperDate) - Int(LowerDate) + 1) * Rnd) + Int(LowerDate))
    End If
    If TimePart = True Then
        ' Calculate a time value rounded to the second.
        TimeValue = CDate(Int(SecondsPerDay * Rnd) / SecondsPerDay)
    End If
    If MillisecondPart = True Then
        ' Calculate a millisecond value rounded to the millisecond.
        MsecValue = CDate(Int(MillisecondsPerSecond * Rnd) / MillisecondsPerSecond / SecondsPerDay)
    End If
    
    ' Assemble date and time parts.
    If DateDiff(IntervalSetting(DtInterval.dtDay), ZeroDateValue, DateValue) >= 0 Then
        DateRandom = DateValue + TimeValue + MsecValue
    Else
        DateRandom = DateValue - TimeValue - MsecValue
    End If

End Function

' Returns a date value from its year, month, day,
' hour, minute, and second part.
' Default values are used for parameters omitted.
'
' Will accept any combination of parameters that can build a
' date/time within the range of Date, including dates before
' 1899-12-30.
'
' For Year values between -1900 and 29, year will be
' relative to year 2000.
' For Year values between 30 and 99, year will be
' relative to year 1900.
'
' A negative timepart will reduce the result relative to the
' datepart, just as if DateAdd had been applied to the date part.
'
' Examples:
'   DateTimeSerial(1899, 12, 29, -1, -1, -1) -> 1899-12-28 22:58:59
'   DateTimeSerial(1899, 12, 29,  1,  1,  1) -> 1899-12-29 01:01:01
'   DateTimeSerial(1899, 12, 30, -1, -1, -1) -> 1899-12-29 22:58:59
'   DateTimeSerial(1899, 12, 30,  1,  1,  1) -> 1899-12-30 01:01:01
'   DateTimeSerial(1899, 12, 31, -1, -1, -1) -> 1899-12-30 22:58:59
'   DateTimeSerial(1899, 12, 31,  1,  1,  1) -> 1899-12-31 01:01:01
'   DateTimeSerial(, , , , 1,  90)           -> 1899-12-30 00:02:30
'   DateTimeSerial(, , , , 1, -90)           -> 1899-12-29 23:59:30
'
' 2016-02-12. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateTimeSerial( _
    Optional ByVal Year As Integer = 1899, _
    Optional ByVal Month As Integer = 12, _
    Optional ByVal Day As Integer = 30, _
    Optional ByVal Hour As Integer, _
    Optional ByVal Minute As Integer, _
    Optional ByVal Second As Integer) _
    As Date
    
    Dim TimePart    As Date
    Dim DatePart    As Date
    Dim ResultDate  As Date
    
    ' DateSerial always returns a date without a time part, thus
    ' it can be used as a timespan as is.
    DatePart = DateSerial(Year, Month, Day)
    ' TimeSerial always returns a timespan, thus can be used as is.
    TimePart = TimeSerial(Hour, Minute, Second)
    
    ' Assemble date and time parts.
    ResultDate = DateFromTimespan(DatePart + TimePart)
        
    DateTimeSerial = ResultDate
    
End Function

' Converts a date value to a timespan value.
' Useful only for date values prior to 1899-12-30 as
' these have a negative numeric value.
'
' 2015-12-15. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateToTimespan( _
    ByVal Value As Date) _
    As Date

    ConvDateToTimespan Value
  
    DateToTimespan = Value

End Function

' Returns the day of the month like the native VBA.Day.
' However, the ultimo date(s) will always be returned as day 30.
'
' Examples:
'   Day30(#2000-01-29#) -> 29
'   Day30(#2000-01-30#) -> 30
'   Day30(#2000-01-31#) -> 30
'   Day30(#2000-02-27#) -> 27
'   Day30(#2000-02-28#) -> 30
'   Day30(#2000-02-29#) -> 30
'   Day30(#2000-03-29#) -> 29
'   Day30(#2000-03-31#) -> 30
'
' 2019-01-26. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function Day30( _
    ByVal Date1 As Date) _
    As Integer
    
    ' Day value for the ultimo date(s) of a banking month.
    Const UltimoDay30   As Integer = 30

    Dim Day     As Integer
    
    If IsDateUltimoMonth(Date1, True) Then
        Day = UltimoDay30
    Else
        Day = VBA.Day(Date1)
    End If
    
    Day30 = Day
    
End Function

' Returns the decade of a date as the first year of the decade.
' Value is 100 for the first decade and 9990 for the last.
'
' 2016-04-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function Decade( _
    ByVal Date1 As Date) _
    As Integer
    
    Dim Result  As Integer
    
    ' Find the decade.
    Result = (Year(Date1) \ YearsPerDecade) * YearsPerDecade
    
    Decade = Result
    
End Function

' Returns the difference in fortnights of two dates based on the ISO 8601 week numbers.
' Result is similar to that of DateDiff.
' For use with DateAddExt and DateDiffExt.
'
' 2016-02-14. Gustav Brock, Cactus Data ApS, CPH.
'
Private Function DiffFortnights( _
    ByVal Date1 As Date, _
    ByVal Date2 As Date) _
    As Long

    Dim Fortnight1  As Integer
    Dim Fortnight2  As Integer
    
    Dim Year1       As Integer
    Dim Year2       As Integer
    Dim Result      As Long

    If DateDiff(IntervalSetting(DtInterval.dtDay), Date1, Date2) <> 0 Then
        ' Retrieve the ISO years of the dates.
        Fortnight1 = Fortnight(Date1, Year1)
        Fortnight2 = Fortnight(Date2, Year2)
        
        ' Count fortnights of Year1 and Year2 and years between these.
        If Year1 = Year2 Then
            Result = Fortnight2 - Fortnight1
        ElseIf Year1 < Year2 Then
            Result = FortnightsOfYear(Year1) - Fortnight1
            While Year1 < Year2
                Year1 = Year1 + 1
                If Year1 = Year2 Then
                    Result = Result + Fortnight2
                Else
                    Result = Result + FortnightsOfYear(Year1)
                End If
            Wend
        ElseIf Year2 < Year1 Then
            Result = Fortnight2 - FortnightsOfYear(Year2)
            While Year2 < Year1
                Year2 = Year2 + 1
                If Year1 = Year2 Then
                    Result = Result - Fortnight1
                Else
                    Result = Result - FortnightsOfYear(Year2)
                End If
            Wend
        End If
    End If
    
    DiffFortnights = Result

End Function

' Converts a numeric negative date value less than one day
' to its numeric positive equivalent.
'
' 2015-12-15. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub EmendTime( _
    ByRef Value As Date)
  
    If Value < 0 Then
        If Value > -1 Then
            ' Convert this invalid date value to a valid date value.
            Value = -Value
        End If
    End If

End Sub

' Returns the fortnight of a date based on the ISO 8601 week number.
' The related ISO year is returned by ref.
'
' Examples:
'   Week  1 -> Fortnight  1
'   Week  2 -> Fortnight  1
'   Week  3 -> Fortnight  2
'   Week  4 -> Fortnight  2
'   Week  5 -> Fortnight  3
'   Week 51 -> Fortnight 26
'   Week 52 -> Fortnight 26
'   Week 53 -> Fortnight 27
'
' 2016-02-06. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function Fortnight( _
    ByVal Date1 As Date, _
    Optional ByRef IsoYear As Integer) _
    As Integer
    
    Dim Month       As Integer
    Dim Result  As Integer
    
    Month = VBA.Month(Date1)
    ' Initially, set the ISO year to the calendar year.
    IsoYear = VBA.Year(Date1)
    
    Result = -Int(-Week(Date1) / 2)
    
    ' Adjust year where week number belongs to next or previous year.
    If Month = MinMonthValue Then
        If Result >= MaxFortnightValue - 1 Then
            ' This is an early date of January belonging to the last fortnight of the previous ISO year.
            IsoYear = IsoYear - 1
        End If
    ElseIf Month = MaxMonthValue Then
        If Result = MinFortnightValue Then
            ' This is a late date of December belonging to the first fortnight of the next ISO year.
            IsoYear = IsoYear + 1
        End If
    End If
    
    ' IsoYear is returned by reference.
    Fortnight = Result
    
End Function

' Returns the weekday within a fortnight based on an ISO 8601 week.
' Parameter FirstDayOfWeek is ignored as the first day must be Monday.
'
' Return values for the first seven days, Monday-Sunday, are 1 to 7.
' Return values for the next seven days, Monday-Sunday, are 8 to 14.
' Note that a fortnight of 27 will only have days of the first week.
'
' 2016-02-06. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function Fortnightday( _
    ByVal Date1 As Date, _
    Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbMonday) _
    As Integer
    
    Dim Result  As Integer
    
    FirstDayOfWeek = vbMonday
    Result = Weekday(Date1, FirstDayOfWeek)
    If Week(Date1) Mod 2 = 0 Then
        Result = Result + DaysPerWeek
    End If
    
    Fortnightday = Result
    
End Function

' Returns the count of a fortnightday between two dates.
'
' 2016-02-14. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function FortnightdayCount( _
    ByVal Date1 As Date, _
    ByVal Date2 As Date) _
    As Long
    
    Dim NextWeek            As Long
    Dim Weeks               As Long
    Dim ThisDate            As Date
    Dim FindFortnightday    As Integer
    Dim UpDown              As Integer
    Dim Result              As Long
    
    FindFortnightday = Fortnightday(Date1)
    
    Weeks = DateDiff(IntervalSetting(DtInterval.dtWeek), Date1, Date2, vbMonday)
    UpDown = Sgn(Weeks)
    If UpDown <> 0 Then
        For NextWeek = UpDown To Weeks Step UpDown
            ThisDate = DateAdd(IntervalSetting(DtInterval.dtWeek), NextWeek, Date1)
            If Fortnightday(ThisDate) = FindFortnightday Then
                ' This week is of the half of a fortnight that does contain FindFortnightDay.
                Result = Result + UpDown
            ElseIf Week(ThisDate) = MaxWeekValue Then
                ' Short fortnight.
                ' This week is always counted as a fortnight.
                Result = Result + UpDown
            Else
                ' This week is of the half of a fortnight that doesn't contain FindFortnightDay.
            End If
        Next
    End If
    
    FortnightdayCount = Result

End Function

' Returns the count of fortnights based on the ISO 8601 week count of a year.
' If Year is not passed, the current year is used.
'
' 2016-02-14. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function FortnightsOfYear( _
    Optional ByVal Year As Integer) _
    As Integer

    Dim Result  As Integer

    If Year = 0 Then
        ' Use year of current date.
        Year = VBA.Year(Date)
    ElseIf Not IsYear(Year) Then
        ' Don't accept years outside the range of Date.
        Err.Raise DtError.dtOverflow
        Exit Function
    End If
    
    ' Fortnight of the last fortnight is the fortnight count of the year.
    Result = Fortnight(DateLastWeekYear(Year))
    
    FortnightsOfYear = Result

End Function

' Returns True if Value can be a century.
'
' 2016-04-11. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsCentury( _
    ByVal Value As Double) _
    As Boolean

    Dim Result  As Boolean

    If IsYear(Value) Then
        If Int(Value / YearsPerCentury) * YearsPerCentury = Value Then
            Result = True
        End If
    End If
   
    IsCentury = Result
    
End Function

' Returns True if Expression can be a date or time.
' Valid numeric values are accepted as CDate will convert these.
' Returns False if Expression is Null or an invalid value.
'
' 2018-04-11. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsDateExt( _
    ByVal Expression As Variant) _
    As Boolean

    Dim Result  As Boolean
    
    If IsNumeric(Expression) Then
        ' Check if the resulting numeric date value is withing limits.
        If CDec(Expression) >= MinNumericDateValue And CDec(Expression) <= MaxNumericDateValue Then
            Expression = CDate(Expression)
        End If
    End If
    Result = IsDate(Expression)
    
    IsDateExt = Result
    
End Function

' Returns True if Value can be a day of a month, optionally,
' a day of a specific month as specified by a date of this month.
'
' 2020-02-12. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsDay( _
    ByVal Value As Double, _
    Optional ByVal DateOfMonth As Date) _
    As Boolean

    Dim Year    As Integer
    Dim Month   As Integer
    Dim Day     As Integer
    Dim Result  As Boolean

    If Value >= MinDayValue And Value <= MaxDayValue Then
        Year = VBA.Year(DateOfMonth)
        Month = VBA.Month(DateOfMonth) + 1
        Day = MinDayValue - 1
        Result = (Value <= VBA.Day(DateSerial(Year, Month, Day)))
    End If
    
    IsDay = Result
    
End Function

' Returns True if Value can be a day of any month.
'
' 2020-02-12. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsDayAllMonths( _
    ByVal Value As Double) _
    As Boolean

    Dim Result  As Boolean

    If Value >= MinDayValue And Value <= MaxDayAllMonthsValue Then
        Result = True
    End If
   
    IsDayAllMonths = Result
    
End Function

' Returns True if Value can be a decade.
'
' 2016-04-11. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsDecade( _
    ByVal Value As Double) _
    As Boolean

    Dim Result  As Boolean

    If IsYear(Value) Then
        If Int(Value / YearsPerDecade) * YearsPerDecade = Value Then
            Result = True
        End If
    End If
   
    IsDecade = Result
    
End Function

' Returns True if Value can be a fortnight of Year.
' If Year is not specified, the current year is used.
'
' 2016-02-14. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsFortnight( _
    ByVal Value As Double, _
    Optional ByVal Year As Integer) _
    As Boolean

    Dim Result  As Boolean
    
    If Year = 0 Then
        Year = VBA.Year(Date)
    ElseIf Not IsYear(Year) Then
        ' Don't accept years outside the range of Date.
        Err.Raise DtError.dtOverflow
        Exit Function
    End If
    
    If Value >= MinFortnightValue And Value <= MaxFortnightValue Then
        If Value <= FortnightsOfYear(Year) Then
            Result = True
        End If
    End If
   
    IsFortnight = Result
    
End Function

' Returns True if Year is a leap year.
' If Year is not passed, the current year is used.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsLeapYear( _
    Optional ByVal Year As Integer) _
    As Boolean

    Const February  As Integer = 2
    Const LastDay   As Integer = 29
    
    Dim Result      As Boolean

    If Year = 0 Then
        ' Use year of current date.
        Year = VBA.Year(Date)
    ElseIf Not IsYear(Year) Then
        ' Don't accept years outside the range of Date.
        Err.Raise DtError.dtOverflow
        Exit Function
    End If
    
    Result = (Day(DateSerial(Year, February, LastDay)) = LastDay)
    
    IsLeapYear = Result
    
End Function

' Returns True if Value can be a millenium.
'
' 2016-04-11. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsMillenium( _
    ByVal Value As Double) _
    As Boolean

    Dim Result  As Boolean

    If IsYear(Value) Then
        If Int(Value / YearsPerMillenium) * YearsPerMillenium = Value Then
            Result = True
        End If
    End If
   
    IsMillenium = Result
    
End Function

' Returns True if Value can be a month.
'
' 2020-02-12. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsMonth( _
    ByVal Value As Double) _
    As Boolean

    Dim Result  As Boolean

    If Value >= MinMonthValue And Value <= MaxMonthValue Then
        Result = True
    End If
   
    IsMonth = Result
    
End Function

' Returns True if Quarter can be a quarter.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsQuarter( _
    ByVal Value As Double) _
    As Boolean

    Dim Result  As Boolean

    If Value >= Quarter(MinDateValue) And Value <= Quarter(MaxDateValue) Then
        Result = True
    End If
   
    IsQuarter = Result
    
End Function

' Returns True if Value can be a semimonth.
'
' 2016-02-15. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsSemimonth( _
    ByVal Value As Double) _
    As Boolean

    Dim Result  As Boolean

    If Value >= MinSemimonthValue And Value <= MaxSemimonthValue Then
        Result = True
    End If
   
    IsSemimonth = Result
    
End Function

' Returns True if Value can be a semiyear.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsSemiyear( _
    ByVal Value As Double) _
    As Boolean

    Dim Result  As Boolean

    If Value >= Semiyear(MinDateValue) And Value <= Semiyear(MaxDateValue) Then
        Result = True
    End If
   
    IsSemiyear = Result
    
End Function

' Returns True if Value can be a sextayear.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsSextayear( _
    ByVal Value As Double) _
    As Boolean

    Dim Result  As Boolean

    If Value >= Sextayear(MinDateValue) And Value <= Sextayear(MaxDateValue) Then
        Result = True
    End If
   
    IsSextayear = Result
    
End Function

' Returns True if Value can be a tertiamonth.
'
' 2019-10-11. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsTertiamonth( _
    ByVal Value As Double) _
    As Boolean

    Dim Result  As Boolean

    If Value >= MinTertiamonthValue And Value <= MaxTertiamonthValue Then
        Result = True
    End If
   
    IsTertiamonth = Result
    
End Function

' Returns True if Value can be a tertiayear.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsTertiayear( _
    ByVal Value As Double) _
    As Boolean

    Dim Result  As Boolean

    If Value >= Tertiayear(MinDateValue) And Value <= Tertiayear(MaxDateValue) Then
        Result = True
    End If
   
    IsTertiayear = Result
    
End Function

' Returns True if IsoWeek can be an ISO 8601 week of IsoYear and
' IsoYear is valid.
' If IsoYear is not specified, the current year is used.
'
' 2017-05-04. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsWeek( _
    ByVal IsoWeek As Integer, _
    Optional ByVal IsoYear As Integer) _
    As Boolean

    Dim Result  As Boolean

    If IsoYear = 0 Then
        IsoYear = VBA.Year(Date)
    End If
        
    If IsYear(IsoYear) Then
        If IsoWeek >= MinWeekValue And IsoWeek <= MaxWeekValue Then
            If IsoWeek <= WeeksOfYear(IsoYear) Then
                Result = True
            End If
        End If
    End If
    
    IsWeek = Result
    
End Function

' Returns True if Expression can be a value of VbDayOfWeek.
' Returns False if Expression is Null or an invalid value.
'
' 2017-01-09. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsWeekday( _
    ByVal Expression As Variant) _
    As Boolean
    
    Dim Result  As Boolean

    Select Case Expression
        Case _
            VbDayOfWeek.vbMonday, _
            VbDayOfWeek.vbTuesday, _
            VbDayOfWeek.vbWednesday, _
            VbDayOfWeek.vbThursday, _
            VbDayOfWeek.vbFriday, _
            VbDayOfWeek.vbSaturday, _
            VbDayOfWeek.vbSunday, _
            VbDayOfWeek.vbUseSystemDayOfWeek
            Result = True
    End Select

    IsWeekday = Result

End Function

' Returns True if Value can be a year.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsYear( _
    ByVal Value As Double) _
    As Boolean

    Dim Result  As Boolean

    If Value >= Year(MinDateValue) And Value <= Year(MaxDateValue) Then
        Result = True
    End If
   
    IsYear = Result
    
End Function

' Returns True if the combination of YearValue and MonthValue
' not will result in a two digit year that DateSerial would
' offset by 2000 years.
'
' Examples:
'   Year:  100  Month:      1   Result: True    DateSerial(100, 1, 1)       ->  100-01-01
'   Year:  100  Month: -23999   Result: False   DateSerial(100, -23999, 1)  ->  100-01-01
'   Year:  100  Month: -24000   Result: False   DateSerial(100, -24000, 1)  -> Error 5, Invalid procedure call.
'   Year:   98  Month:     24   Result: False   DateSerial(98, 24, 1)       -> 1999-12-01
'   Year:   98  Month:     26   Result: True    DateSerial(98, 26, 1)       ->  100-02-01
'   Year:10000  Month:      0   Result: True    DateSerial(10000, 0, 1)     -> 9999-12-01
'   Year: 9998  Month:     25   Result: False   DateSerial(9998, 25, 1)     -> Error 5, Invalid procedure call.
'   Year: 8000  Month:  24000   Result: True    DateSerial(8000, 24000, 1)  -> 9999-12-01
'   Year: 8000  Month:  24001   Result: False   DateSerial(8000, 24001, 1)  -> Error 5, Invalid procedure call.
'
' 2019-10-30. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsYearMonth( _
    ByVal YearValue As Integer, _
    ByVal MonthValue As Integer) _
    As Boolean
    
    Dim Months  As Long
    Dim Result  As Boolean
    
    ' Convert to Long to prevent overflow of Integer for extreme values of year.
    Months = YearValue * CLng(MonthsPerYear) + MonthValue
    
    ' Validate minimum value.
    If Months >= 1 + Year(MinDateValue) * MonthsPerYear Then
        ' Validate maximum value.
        If Months <= (1 + Year(MaxDateValue)) * MonthsPerYear Then
            Result = True
        End If
    End If
    
    IsYearMonth = Result
    
End Function

' Joins the date and time parts of a DateTime value.
' Returns the result as a single Date value.
'
' The numeric time part must be >= 0 and < 1 or an
' error is raised.
'
' 2016-07-05. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function JoinDateTime( _
    ByRef DateTime1 As DateTime) _
    As Date
    
    Dim Value   As Double
    Dim Result  As Date
    
    ' Raise error if Time value is outside the allowed range.
    Value = CDbl(DateTime1.Time)
    If Value < 0 Or Value >= 1 Then
        Err.Raise DtError.dtInvalidProcedureCallOrArgument
    End If
    
    If DateTime1.Date >= 0 Then
        Result = DateTime1.Date + DateTime1.Time
    Else
        Result = DateTime1.Date - DateTime1.Time
    End If
    
    JoinDateTime = Result
    
End Function

' Returns the millenium of a date as the first year of the millenium.
' Value is 0 for the first millenium and 9000 for the last.
'
' 2016-04-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function Millenium( _
    ByVal Date1 As Date) _
    As Integer
    
    Dim Result  As Integer
    
    ' Find the millenium.
    Result = (Year(Date1) \ YearsPerMillenium) * YearsPerMillenium
    
    Millenium = Result
    
End Function

' Returns the quarter of a date.
' Value is 1 for the first quarter of the year,
' 4 for the last.
'
' 2015-12-21. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function Quarter( _
    ByVal Date1 As Date) _
    As Integer
    
    Dim Result  As Integer
    
    ' Find the quarter.
    Result = DatePart(IntervalSetting(DtInterval.dtQuarter), Date1)
    
    Quarter = Result
    
End Function

' Returns the second of a date.
'
' Will return correct reading of seconds of the last day of Date (9999-12-31).
' See function DateTest.TheLastSeconds for full information.
'
' 2015-12-21. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function SecondExt( _
    ByVal Time1 As Date) _
    As Integer
    
    Dim Result  As Integer
    
    ' Native reading of seconds is buggy for the last date.
    ' Obtain correct reading by removing the date part, thus
    ' reading the seconds from the time part only.
    Result = Second(Time1 - Fix(Time1))
    
    SecondExt = Result

End Function

' Returns the semimonth of a date.
' Value is 1 for the first half of a month of the year,
' 24 for the last.
'
' 2019-10-11. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function Semimonth( _
    ByVal Date1 As Date) _
    As Integer
    
    Dim Result  As Integer
    
    ' Find the semimonth.
    Result = Month(Date1) * SemimonthsPerMonth
    If Day(Date1) <= DaysPerSemimonth Then
        Result = Result - 1
    End If
    
    Semimonth = Result
    
End Function

' Returns for the semimonth of the date passed, this
' semimonth's part of the month of that date.
'
' Examples:
'   SemimonthPart(#2000-02-01#)     -> 1
'   SemimonthPart(#2000-02-15#)     -> 1
'   SemimonthPart(#2000-04-16#)     -> 2
'   SemimonthPart(#2000-05-31#)     -> 2
'
' 2019-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function SemimonthPart( _
    ByVal Date1 As Date) _
    As Integer
    
    Const UltimoLastPart    As Integer = SemimonthsPerMonth * DaysPerSemimonth
    
    Dim Part    As Integer
    Dim Day     As Integer
    
    Day = VBA.Day(Date1)
    If Day > UltimoLastPart Then
        Day = UltimoLastPart
    End If
    
    Part = 1 + (Day - 1) \ DaysPerSemimonth
    
    SemimonthPart = Part
    
End Function

' Returns the semiyear of a date.
' Value is 1 for the first semiyear of a year,
' 2 for the second.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function Semiyear( _
    ByVal Date1 As Date) _
    As Integer
    
    Dim Result  As Integer
    
    ' Calculate the semiyear from months.
    Result = -Int(-Month(Date1) / DtIntervalMonths.dtSemiyear)
    
    Semiyear = Result
    
End Function

' Returns the sextayear of a date.
' Value is 1 for the first sextayear of a year,
' 2 for the second, 6 for the last.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function Sextayear( _
    ByVal Date1 As Date) _
    As Integer
    
    Dim Result  As Integer
    
    ' Calculate the sextayear from months.
    Result = -Int(-Month(Date1) / DtIntervalMonths.dtSextayear)
    
    Sextayear = Result
    
End Function

' Splits a Date value into its date and time parts.
' Returns the result as a DateTime value.
'
' 2016-01-21. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function SplitDateTime( _
    ByVal Date1 As Date) _
    As DateTime
    
    Dim Result  As DateTime
    
    Result.Date = Fix(Date1)
    Result.Time = Abs(Date1 - Result.Date)
    
    SplitDateTime = Result
    
End Function

' Returns the tertiamonth of a date.
' Value is 1 for the first third of a month of the year,
' 36 for the last.
'
' 2019-10-11. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function Tertiamonth( _
    ByVal Date1 As Date) _
    As Integer
    
    Dim Result  As Integer
    
    ' Find the tertiamonth.
    Result = Month(Date1) * TertiamonthsPerMonth
    Select Case Day(Date1)
        Case Is <= DaysPerTertiamonth
            Result = Result - 2
        Case Is <= DaysPerTertiamonth * 2
            Result = Result - 1
    End Select
    
    Tertiamonth = Result
    
End Function

' Returns for the tertiamonth of the date passed, this
' tertiamonth's part of the month of that date.
'
' Examples:
'   TertiamonthPart(#2000-02-01#)     -> 1
'   TertiamonthPart(#2000-02-10#)     -> 1
'   TertiamonthPart(#2000-02-11#)     -> 2
'   TertiamonthPart(#2000-04-20#)     -> 2
'   TertiamonthPart(#2000-04-21#)     -> 3
'   TertiamonthPart(#2000-05-30#)     -> 3
'   TertiamonthPart(#2000-05-31#)     -> 3
'
' 2019-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function TertiamonthPart( _
    ByVal Date1 As Date) _
    As Integer
    
    Const UltimoLastPart    As Integer = TertiamonthsPerMonth * DaysPerTertiamonth
    
    Dim Part    As Integer
    Dim Day     As Integer
    
    Day = VBA.Day(Date1)
    If Day > UltimoLastPart Then
        Day = UltimoLastPart
    End If
    
    Part = 1 + (Day - 1) \ DaysPerTertiamonth
    
    TertiamonthPart = Part
    
End Function

' Returns the tertiayear of a date.
' Value is 1 for the first tertiayear of a year,
' 2 for the second, 3 for the last.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function Tertiayear( _
    ByVal Date1 As Date) _
    As Integer
    
    Dim Result  As Integer
    
    ' Calculate the tertiayear from months.
    Result = -Int(-Month(Date1) / DtIntervalMonths.dtTertiayear)
    
    Tertiayear = Result
    
End Function

' Returns an invalid numeric negative time value
' as its positive equivalent.
'
' 2015-12-15. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function TimeEmend( _
    ByVal Value As Date) _
    As Date
  
    ' Emend the time value.
    EmendTime Value
  
    TimeEmend = Value
  
End Function

' Returns numeric negative true date values, which TimeSerial() does not.
'
' Example sequence for 1899-12-30 +/-48 hours:
'
' TimeSerialDate                                 TimeSerial
' Hours Date value               Numeric value   Date value               Numeric value
'   48  1900-01-01 00:00:00.000  2               1900-01-01 00:00:00.000  2
'   42  1899-12-31 18:00:00.000  1.75            1899-12-31 18:00:00.000  1.75
'   36  1899-12-31 12:00:00.000  1.5             1899-12-31 12:00:00.000  1.5
'   30  1899-12-31 06:00:00.000  1.25            1899-12-31 06:00:00.000  1.25
'   24  1899-12-31 00:00:00.000  1               1899-12-31 00:00:00.000  1
'   18  1899-12-30 18:00:00.000  0.75            1899-12-30 18:00:00.000  0.75
'   12  1899-12-30 12:00:00.000  0.5             1899-12-30 12:00:00.000  0.5
'    6  1899-12-30 06:00:00.000  0.25            1899-12-30 06:00:00.000  0.25
'    0  1899-12-30 00:00:00.000  0               1899-12-30 00:00:00.000  0
'  - 6  1899-12-29 18:00:00.000 -1.75          * 1899-12-30 06:00:00.000 -0.25
'  -12  1899-12-29 12:00:00.000 -1.5           * 1899-12-30 12:00:00.000 -0.5
'  -18  1899-12-29 06:00:00.000 -1.25          * 1899-12-30 18:00:00.000 -0.75
'  -24  1899-12-29 00:00:00.000 -1               1899-12-29 00:00:00.000 -1
'  -30  1899-12-28 18:00:00.000 -2.75          * 1899-12-29 06:00:00.000 -1.25
'  -36  1899-12-28 12:00:00.000 -2.5           * 1899-12-29 12:00:00.000 -1.5
'  -42  1899-12-28 06:00:00.000 -2.25          * 1899-12-29 18:00:00.000 -1.75
'  -48  1899-12-28 00:00:00.000 -2               1899-12-28 00:00:00.000 -2
'
' 2016-09-18. Cactus Data ApS, CPH.
'
Public Function TimeSerialDate( _
    ByVal Hour As Integer, _
    ByVal Minute As Integer, _
    ByVal Second As Integer) _
    As Date
    
    Dim DateValue   As Date
    Dim DatePart    As Double
    Dim TimePart    As Double
  
    DateValue = TimeSerial(Hour, Minute, Second)
    If DateValue < 0 Then
        ' Get the date (integer) part of DateValue shifted by one day
        ' if a time part is present as Int() rounds down.
        DatePart = Int(DateValue)
        ' Retrieve and reverse the time (decimal) part.
        TimePart = DatePart - DateValue
        ' Assemble and convert date and time parts.
        DateValue = CDate(DatePart + TimePart)
    End If
      
    TimeSerialDate = DateValue

End Function

' Returns the ISO 8601 week of a date.
' The related ISO year is returned by ref.
'
' 2016-01-06. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function Week( _
    ByVal Date1 As Date, _
    Optional ByRef IsoYear As Integer) _
    As Integer

    Dim Month       As Integer
    Dim Interval    As String
    Dim Result      As Integer
    
    Interval = IntervalSetting(dtWeek)
    
    Month = VBA.Month(Date1)
    ' Initially, set the ISO year to the calendar year.
    IsoYear = VBA.Year(Date1)
    
    Result = DatePart(Interval, Date1, vbMonday, vbFirstFourDays)
    If Result = MaxWeekValue Then
        If DatePart(Interval, DateAdd(Interval, 1, Date1), vbMonday, vbFirstFourDays) = MinWeekValue Then
            ' OK. The next week is the first week of the following year.
        Else
            ' This is really the first week of the next ISO year.
            ' Correct for DatePart bug.
            Result = MinWeekValue
        End If
    End If
        
    ' Adjust year where week number belongs to next or previous year.
    If Month = MinMonthValue Then
        If Result >= MaxWeekValue - 1 Then
            ' This is an early date of January belonging to the last week of the previous ISO year.
            IsoYear = IsoYear - 1
        End If
    ElseIf Month = MaxMonthValue Then
        If Result = MinWeekValue Then
            ' This is a late date of December belonging to the first week of the next ISO year.
            IsoYear = IsoYear + 1
        End If
    End If
    
    ' IsoYear is returned by reference.
    Week = Result
        
End Function

' Returns the ISO 8601 week count of a year.
' If Year is not passed, the current year is used.
'
' 2015-12-04. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function WeeksOfYear( _
    Optional ByVal IsoYear As Integer) _
    As Integer

    Dim Result  As Integer

    If IsoYear = 0 Then
        ' Use year of current date.
        IsoYear = VBA.Year(Date)
    ElseIf Not IsYear(IsoYear) Then
        ' Don't accept years outside the range of Date.
        Err.Raise DtError.dtOverflow
        Exit Function
    End If
    
    ' Weeknumber of the last week is the week count of the year.
    Result = Week(DateLastWeekYear(IsoYear))
    
    WeeksOfYear = Result

End Function

' Returns the ISO 8601 year of a date.
'
' 2016-01-06. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function YearOfWeek( _
    ByVal Date1 As Date) _
    As Integer
    
    Dim IsoYear As Integer

    ' Get the ISO 8601 year of Date1.
    Week Date1, IsoYear
    
    YearOfWeek = IsoYear

End Function


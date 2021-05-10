Attribute VB_Name = "DateFind"
Option Explicit
'
' DateFind
' Version 1.2.1
'
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.Date
'
' Functions for finding various dates and values from one or more given date and time values.
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
'   DateMsec
'

' Used by:
'   SystemFirstWeekOfYear
'
Private Declare PtrSafe Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" ( _
    ByVal Locale As Long, _
    ByVal LCType As Long, _
    ByVal lpLCData As String, _
    ByVal cchData As Long) _
    As Long

' Returns a date added a number of months.
' Result date is identical to the result of DateAdd("m", Number, Date1)
' except that if an ultimo date of a month is passed, an ultimo date will
' always be returned as well.
' Optionally, days of the 30th (or 28th of February of leap years) will
' also be regarded as ultimo.
'
' Examples:
'   2020-02-28, 1, False -> 2020-03-28
'   2020-02-28, 1, True  -> 2020-03-31
'   2020-02-28,-2, False -> 2019-12-28
'   2020-02-28,-2, True  -> 2019-12-31
'   2020-02-28, 4, False -> 2020-06-28
'   2020-02-28, 4, True  -> 2020-06-30
'   2020-02-29, 4, False -> 2020-06-30
'   2020-06-30, 2, False -> 2020-08-31
'   2020-06-30, 2, True  -> 2020-08-31
'   2020-07-30, 2, False -> 2020-08-30
'   2020-07-30, 1, True  -> 2020-08-31
'
' 2015-11-30. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateAddMonth( _
    ByVal Date1 As Date, _
    Number As Double, _
    Optional Include2830 As Boolean) _
    As Date
    
    Dim DateNext    As Date
    
    ' Add number of months the normal way.
    DateNext = DateAdd("m", Number, Date1)
    If Day(Date1) = MaxDayValue Then
        ' Resulting day will be ultimo of the month.
    ElseIf IsDateUltimoMonth(Date1, Include2830) = True Then
        ' Months are added to month ultimo.
        ' Adjust resulting day to be ultimo if not.
        If IsDateUltimoMonth(DateNext, False) = False Then
            ' Resulting day is not ultimo of the month.
            ' Set resulting day to ultimo of the month.
            DateNext = DateThisMonthUltimo(DateNext)
        End If
    End If
    
    DateAddMonth = DateNext
    
End Function

' Calculates next annual day following SomeDate based on AnnualDay.
' Calculates correctly for leap years if AnnualDay is Feb. 29th.
' If SomeDate is earlier than AnnualDay, AnnualDay is returned.
' If next annual day should be later than 9999-12-31, the annual day
' of year 9999 is returned.
'
' 2015-11-21. Gustav Brock. Cactus Data ApS, CPH.
'                                                   wrapper til DateNextPeriodicDay?
Public Function DateNextAnnualDay( _
    ByVal AnnualDay As Date, _
    Optional ByVal SomeDate As Variant) _
    As Date
    
    Dim NextAnnualDay   As Date
    Dim Years           As Integer
    
    If Not IsDateExt(SomeDate) Then
        ' Empty or invalid parameter SomeDate.
        ' Use today as SomeDate.
        SomeDate = Date
    End If
    
    NextAnnualDay = AnnualDay
    If DateDiff(IntervalSetting(DtInterval.dtYear), AnnualDay, MaxDateValue) = 0 Then
        ' No later annual day can be calculated.
    Else
        Years = DateDiff(IntervalSetting(DtInterval.dtYear), AnnualDay, SomeDate)
        If Years < 0 Then
            ' Don't calculate hypothetical annual days.
        Else
            NextAnnualDay = DateAdd(IntervalSetting(DtInterval.dtYear), Years, AnnualDay)
            If DateDiff("d", SomeDate, NextAnnualDay) <= 0 Then
                ' Next annual day falls earlier in the year than SomeDate.
                If DateDiff(IntervalSetting(DtInterval.dtYear), NextAnnualDay, MaxDateValue) = 0 Then
                    ' No later annual day can be calculated.
                Else
                    NextAnnualDay = DateAdd(IntervalSetting(DtInterval.dtYear), Years + 1, AnnualDay)
                End If
            End If
        End If
    End If
    
    DateNextAnnualDay = NextAnnualDay
  
End Function

' Returns the earliest date and time of the date following the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateNextDayPrimo( _
    ByVal DateThisDay As Date) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = 1
    Interval = IntervalSetting(DtInterval.dtDay)
    
    ResultDate = DateIntervalPrimo(Interval, Number, DateThisDay)
    
    DateNextDayPrimo = ResultDate
    
End Function

' Returns the latest date and time of the date following the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateNextDayUltimo( _
    ByVal DateThisDay As Date) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = 1
    Interval = IntervalSetting(DtInterval.dtDay)
    
    ResultDate = DateIntervalUltimo(Interval, Number, DateThisDay)
    
    DateNextDayUltimo = ResultDate
    
End Function

' Returns the primo date of the month following the month of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateNextMonthPrimo( _
    ByVal DateThisMonth As Date) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = 1
    Interval = IntervalSetting(DtInterval.dtMonth)
    
    ResultDate = DateIntervalPrimo(Interval, Number, DateThisMonth)
    
    DateNextMonthPrimo = ResultDate
    
End Function

' Returns the ultimo date of the month following the month of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateNextMonthUltimo( _
    ByVal DateThisMonth As Date) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = 1
    Interval = IntervalSetting(DtInterval.dtMonth)
    
    ResultDate = DateIntervalUltimo(Interval, Number, DateThisMonth)
    
    DateNextMonthUltimo = ResultDate
    
End Function

   
' Calculates next periodic date based on Date1 using the specified
' interval and, optionally, an interval count larger than 1.
' Optionally, another date than today can be specified and used
' as reference to find the next periodic date.
'
' Will accept any value for Date1 including ultimo month dates.
' If the resulting date would be outside the range of data type Date,
' an error is raised.
'
' 2020-07-16. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function DateNextPeriodicDay( _
    ByVal Date1 As Date, _
    ByVal Interval As DtInterval, _
    Optional ByVal IntervalCount As Long = 1, _
    Optional ByVal ReferenceDate As Variant) _
    As Date
    
    Dim Intervals       As Long
    Dim NextPeriodicDay As Date
    
    If Not IsDateExt(ReferenceDate) Then
        ' Empty or invalid parameter ReferenceDate.
        ' Use today as ReferenceDate.
        ReferenceDate = Date
    End If
    
    If IntervalCount > 0 Then
        If DateDiff(IntervalSetting(Interval), Date1, MaxDateValue) = 0 Then
            ' No later periodic day can be calculated.
            Err.Raise DtError.dtInvalidProcedureCallOrArgument
        Else
            Intervals = IntervalCount * Int(DateDiff(IntervalSetting(Interval), Date1, ReferenceDate) / IntervalCount)
            NextPeriodicDay = DateAdd(IntervalSetting(Interval), Intervals, Date1)
            If DateDiff(IntervalSetting(dtDay), ReferenceDate, NextPeriodicDay) <= 0 Then
                ' The next periodic day would fall earlier in the period than ReferenceDate.
                ' Find the next periodic day.
                If DateDiff(IntervalSetting(Interval), NextPeriodicDay, MaxDateValue) = 0 Then
                    ' No later periodic day can be calculated.
                    Err.Raise DtError.dtInvalidProcedureCallOrArgument
                Else
                    NextPeriodicDay = DateAdd(IntervalSetting(Interval), Intervals + IntervalCount, Date1)
                End If
            End If
        End If
    Else
        NextPeriodicDay = Date1
    End If
      
    DateNextPeriodicDay = NextPeriodicDay
  
End Function

' Returns the primo date of the quarter following the quarter of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateNextQuarterPrimo( _
    ByVal DateThisQuarter As Date) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = 1
    Interval = IntervalSetting(DtInterval.dtQuarter)
    
    ResultDate = DateIntervalPrimo(Interval, Number, DateThisQuarter)
    
    DateNextQuarterPrimo = ResultDate
    
End Function

' Returns the ultimo date of the quarter following the quarter of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateNextQuarterUltimo( _
    ByVal DateThisQuarter As Date) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = 1
    Interval = IntervalSetting(DtInterval.dtQuarter)
    
    ResultDate = DateIntervalUltimo(Interval, Number, DateThisQuarter)
    
    DateNextQuarterUltimo = ResultDate
    
End Function

' Returns the primo date of the semiyear following the semiyear of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateNextSemiyearPrimo( _
    ByVal DateThisSemiyear As Date) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = 1
    Interval = IntervalSetting(DtInterval.dtSemiyear, True)
    
    ResultDate = DateIntervalPrimo(Interval, Number, DateThisSemiyear)
    
    DateNextSemiyearPrimo = ResultDate
    
End Function

' Returns the ultimo date of the semiyear following the semiyear of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateNextSemiyearUltimo( _
    ByVal DateThisSemiyear As Date) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = 1
    Interval = IntervalSetting(DtInterval.dtSemiyear, True)
    
    ResultDate = DateIntervalUltimo(Interval, Number, DateThisSemiyear)
    
    DateNextSemiyearUltimo = ResultDate
    
End Function

' Returns the primo date of the sextayear following the sextayear of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateNextSextayearPrimo( _
    ByVal DateThisSextayear As Date) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = 1
    Interval = IntervalSetting(DtInterval.dtSextayear, True)
    
    ResultDate = DateIntervalPrimo(Interval, Number, DateThisSextayear)
    
    DateNextSextayearPrimo = ResultDate
    
End Function

' Returns the ultimo date of the sextayear following the sextayear of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateNextSextayearUltimo( _
    ByVal DateThisSextayear As Date) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = 1
    Interval = IntervalSetting(DtInterval.dtSextayear, True)
    
    ResultDate = DateIntervalUltimo(Interval, Number, DateThisSextayear)
    
    DateNextSextayearUltimo = ResultDate
    
End Function

' Returns the primo date of the tertiayear following the tertiayear of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateNextTertiayearPrimo( _
    ByVal DateThisTertiayear As Date) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = 1
    Interval = IntervalSetting(DtInterval.dtTertiayear, True)
    
    ResultDate = DateIntervalPrimo(Interval, Number, DateThisTertiayear)
    
    DateNextTertiayearPrimo = ResultDate
    
End Function

' Returns the ultimo date of the tertiayear following the tertiayear of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateNextTertiayearUltimo( _
    ByVal DateThisTertiayear As Date) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = 1
    Interval = IntervalSetting(DtInterval.dtTertiayear, True)
    
    ResultDate = DateIntervalUltimo(Interval, Number, DateThisTertiayear)
    
    DateNextTertiayearUltimo = ResultDate
    
End Function

' Returns the date of the weekday as specified by DayOfWeek
' following Date1.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateNextWeekday( _
    ByVal Date1 As Date, _
    Optional ByVal DayOfWeek As VbDayOfWeek = vbUseSystemDayOfWeek) _
    As Date

    Dim Interval    As String
    Dim ResultDate  As Date
    
    Interval = IntervalSetting(DtInterval.dtDay)
    
    If DayOfWeek = vbUseSystemDayOfWeek Then
        DayOfWeek = Weekday(Date1)
    End If
    
    ResultDate = DateAdd(Interval, DaysPerWeek - (Weekday(Date1, DayOfWeek) - 1), Date1)
    
    DateNextWeekday = ResultDate
    
End Function

' Returns the primo date of the week following the week of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateNextWeekPrimo( _
    ByVal DateThisWeek As Date, _
    Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbSunday) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = 1
    Interval = IntervalSetting(DtInterval.dtWeek)
    
    ResultDate = DateIntervalPrimo(Interval, Number, DateThisWeek, FirstDayOfWeek)
    
    DateNextWeekPrimo = ResultDate
    
End Function

' Returns the ultimo date of the week following the week of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateNextWeekUltimo( _
    ByVal DateThisWeek As Date, _
    Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbSunday) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = 1
    Interval = IntervalSetting(DtInterval.dtWeek)
    
    ResultDate = DateIntervalUltimo(Interval, Number, DateThisWeek, FirstDayOfWeek)
    
    DateNextWeekUltimo = ResultDate
    
End Function

' Returns the primo date of the year following the year of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateNextYearPrimo( _
    ByVal DateThisYear As Date) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = 1
    Interval = IntervalSetting(DtInterval.dtYear)
    
    ResultDate = DateIntervalPrimo(Interval, Number, DateThisYear)
    
    DateNextYearPrimo = ResultDate
    
End Function

' Returns the ultimo date of the year following the year of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateNextYearUltimo( _
    ByVal DateThisYear As Date) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = 1
    Interval = IntervalSetting(DtInterval.dtYear)
    
    ResultDate = DateIntervalUltimo(Interval, Number, DateThisYear)
    
    DateNextYearUltimo = ResultDate
    
End Function

' Returns the earliest date and time of the date preceding the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DatePreviousDayPrimo( _
    ByVal DateThisDay As Date) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = -1
    Interval = IntervalSetting(DtInterval.dtDay)
    
    ResultDate = DateIntervalPrimo(Interval, Number, DateThisDay)
    
    DatePreviousDayPrimo = ResultDate
    
End Function

' Returns the latest date and time of the date preceding the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DatePreviousDayUltimo( _
    ByVal DateThisDay As Date) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = -1
    Interval = IntervalSetting(DtInterval.dtDay)
    
    ResultDate = DateIntervalUltimo(Interval, Number, DateThisDay)
    
    DatePreviousDayUltimo = ResultDate
    
End Function

' Returns the primo date of the month preceding the month of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DatePreviousMonthPrimo( _
    ByVal DateThisMonth As Date) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = -1
    Interval = IntervalSetting(DtInterval.dtMonth)
    
    ResultDate = DateIntervalPrimo(Interval, Number, DateThisMonth)
    
    DatePreviousMonthPrimo = ResultDate
    
End Function

' Returns the ultimo date of the month preceding the month of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DatePreviousMonthUltimo( _
    ByVal DateThisMonth As Date) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = -1
    Interval = IntervalSetting(DtInterval.dtMonth)
    
    ResultDate = DateIntervalUltimo(Interval, Number, DateThisMonth)
    
    DatePreviousMonthUltimo = ResultDate
    
End Function

' Returns the primo date of the quarter preceding the quarter of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DatePreviousQuarterPrimo( _
    ByVal DateThisQuarter As Date) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = -1
    Interval = IntervalSetting(DtInterval.dtQuarter)
    
    ResultDate = DateIntervalPrimo(Interval, Number, DateThisQuarter)
    
    DatePreviousQuarterPrimo = ResultDate
    
End Function

' Returns the ultimo date of the quarter preceding the quarter of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DatePreviousQuarterUltimo( _
    ByVal DateThisQuarter As Date) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = -1
    Interval = IntervalSetting(DtInterval.dtQuarter)
    
    ResultDate = DateIntervalUltimo(Interval, Number, DateThisQuarter)
    
    DatePreviousQuarterUltimo = ResultDate
    
End Function

' Returns the primo date of the semiyear preceding the semiyear of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DatePreviousSemiyearPrimo( _
    ByVal DateThisSemiyear As Date) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = -1
    Interval = IntervalSetting(DtInterval.dtSemiyear, True)
    
    ResultDate = DateIntervalPrimo(Interval, Number, DateThisSemiyear)
    
    DatePreviousSemiyearPrimo = ResultDate
    
End Function

' Returns the ultimo date of the semiyear preceding the semiyear of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DatePreviousSemiyearUltimo( _
    ByVal DateThisSemiyear As Date) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = -1
    Interval = IntervalSetting(DtInterval.dtSemiyear, True)
    
    ResultDate = DateIntervalUltimo(Interval, Number, DateThisSemiyear)
    
    DatePreviousSemiyearUltimo = ResultDate
    
End Function

' Returns the primo date of the sextayear preceding the sextayear of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DatePreviousSextayearPrimo( _
    ByVal DateThisSextayear As Date) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = -1
    Interval = IntervalSetting(DtInterval.dtSextayear, True)
    
    ResultDate = DateIntervalPrimo(Interval, Number, DateThisSextayear)
    
    DatePreviousSextayearPrimo = ResultDate
    
End Function

' Returns the ultimo date of the sextayear preceding the sextayear of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DatePreviousSextayearUltimo( _
    ByVal DateThisSextayear As Date) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = -1
    Interval = IntervalSetting(DtInterval.dtSextayear, True)
    
    ResultDate = DateIntervalUltimo(Interval, Number, DateThisSextayear)
    
    DatePreviousSextayearUltimo = ResultDate
    
End Function

' Returns the primo date of the tertiayear preceding the tertiayear of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DatePreviousTertiayearPrimo( _
    ByVal DateThisTertiayear As Date) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = -1
    Interval = IntervalSetting(DtInterval.dtTertiayear, True)
    
    ResultDate = DateIntervalPrimo(Interval, Number, DateThisTertiayear)
    
    DatePreviousTertiayearPrimo = ResultDate
    
End Function

' Returns the ultimo date of the tertiayear preceding the tertiayear of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DatePreviousTertiayearUltimo( _
    ByVal DateThisTertiayear As Date) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = -1
    Interval = IntervalSetting(DtInterval.dtTertiayear, True)
    
    ResultDate = DateIntervalUltimo(Interval, Number, DateThisTertiayear)
    
    DatePreviousTertiayearUltimo = ResultDate
    
End Function

' Returns the date of the weekday as specified by DayOfWeek
' preceding Date1.
'
' Note: If DayOfWeek is omitted, the weekday of Date1 is used.
' If so, the date returned will always be Date1 - 7.
'
' 2019-06-29. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DatePreviousWeekday( _
    ByVal Date1 As Date, _
    Optional ByVal DayOfWeek As VbDayOfWeek = vbUseSystemDayOfWeek) _
    As Date

    Dim Interval    As String
    Dim ResultDate  As Date
    
    Interval = IntervalSetting(DtInterval.dtDay)
    
    If DayOfWeek = vbUseSystemDayOfWeek Then
        DayOfWeek = Weekday(Date1)
    End If
    
    ResultDate = DateAdd(Interval, -DaysPerWeek - ((Weekday(Date1, DayOfWeek) - DaysPerWeek - 1) Mod DaysPerWeek), Date1)
    
    DatePreviousWeekday = ResultDate
    
End Function

' Returns the primo date of the week preceding the week of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DatePreviousWeekPrimo( _
    ByVal DateThisWeek As Date, _
    Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbSunday) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = -1
    Interval = IntervalSetting(DtInterval.dtWeek)
    
    ResultDate = DateIntervalPrimo(Interval, Number, DateThisWeek, FirstDayOfWeek)
    
    DatePreviousWeekPrimo = ResultDate
    
End Function

' Returns the ultimo date of the week preceding the week of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DatePreviousWeekUltimo( _
    ByVal DateThisWeek As Date, _
    Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbSunday) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = -1
    Interval = IntervalSetting(DtInterval.dtWeek)
    
    ResultDate = DateIntervalUltimo(Interval, Number, DateThisWeek, FirstDayOfWeek)
    
    DatePreviousWeekUltimo = ResultDate
    
End Function

' Returns the primo date of the year preceding the year of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DatePreviousYearPrimo( _
    ByVal DateThisYear As Date) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = -1
    Interval = IntervalSetting(DtInterval.dtYear)
    
    ResultDate = DateIntervalPrimo(Interval, Number, DateThisYear)
    
    DatePreviousYearPrimo = ResultDate
    
End Function

' Returns the ultimo date of the year preceding the year of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DatePreviousYearUltimo( _
    ByVal DateThisYear As Date) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = -1
    Interval = IntervalSetting(DtInterval.dtYear)
    
    ResultDate = DateIntervalUltimo(Interval, Number, DateThisYear)
    
    DatePreviousYearUltimo = ResultDate
    
End Function

' Returns the earliest date and time of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateThisDayPrimo( _
    ByVal DateThisDay As Date) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = 0
    Interval = IntervalSetting(DtInterval.dtDay)
    
    ResultDate = DateIntervalPrimo(Interval, Number, DateThisDay)
    
    DateThisDayPrimo = ResultDate
    
End Function

' Returns the latest date and time of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateThisDayUltimo( _
    ByVal DateThisDay As Date) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = 0
    Interval = IntervalSetting(DtInterval.dtDay)
    
    ResultDate = DateIntervalUltimo(Interval, Number, DateThisDay)
    
    DateThisDayUltimo = ResultDate
    
End Function

' Returns the primo date of the month of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateThisMonthPrimo( _
    ByVal DateThisMonth As Date) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = 0
    Interval = IntervalSetting(DtInterval.dtMonth)
    
    ResultDate = DateIntervalPrimo(Interval, Number, DateThisMonth)
    
    DateThisMonthPrimo = ResultDate
    
End Function

' Returns the ultimo date of the month of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateThisMonthUltimo( _
    ByVal DateThisMonth As Date) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = 0
    Interval = IntervalSetting(DtInterval.dtMonth)
    
    ResultDate = DateIntervalUltimo(Interval, Number, DateThisMonth)
    
    DateThisMonthUltimo = ResultDate
    
End Function

' Returns the primo date of the quarter of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateThisQuarterPrimo( _
    ByVal DateThisQuarter As Date) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = 0
    Interval = IntervalSetting(DtInterval.dtQuarter)
    
    ResultDate = DateIntervalPrimo(Interval, Number, DateThisQuarter)
    
    DateThisQuarterPrimo = ResultDate
    
End Function

' Returns the ultimo date of the quarter of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateThisQuarterUltimo( _
    ByVal DateThisQuarter As Date) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = 0
    Interval = IntervalSetting(DtInterval.dtQuarter)
    
    ResultDate = DateIntervalUltimo(Interval, Number, DateThisQuarter)
    
    DateThisQuarterUltimo = ResultDate
    
End Function

' Returns the primo date of the semiyear of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateThisSemiyearPrimo( _
    ByVal DateThisSemiyear As Date) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = 0
    Interval = IntervalSetting(DtInterval.dtSemiyear, True)
    
    ResultDate = DateIntervalPrimo(Interval, Number, DateThisSemiyear)
    
    DateThisSemiyearPrimo = ResultDate
    
End Function

' Returns the ultimo date of the semiyear of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateThisSemiyearUltimo( _
    ByVal DateThisSemiyear As Date) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = 0
    Interval = IntervalSetting(DtInterval.dtSemiyear, True)
    
    ResultDate = DateIntervalUltimo(Interval, Number, DateThisSemiyear)
    
    DateThisSemiyearUltimo = ResultDate
    
End Function

' Returns the primo date of the sextayear of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateThisSextayearPrimo( _
    ByVal DateThisSextayear As Date) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = 0
    Interval = IntervalSetting(DtInterval.dtSextayear, True)
    
    ResultDate = DateIntervalPrimo(Interval, Number, DateThisSextayear)
    
    DateThisSextayearPrimo = ResultDate
    
End Function

' Returns the ultimo date of the sextayear of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateThisSextayearUltimo( _
    ByVal DateThisSextayear As Date) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = 0
    Interval = IntervalSetting(DtInterval.dtSextayear, True)
    
    ResultDate = DateIntervalUltimo(Interval, Number, DateThisSextayear)
    
    DateThisSextayearUltimo = ResultDate
    
End Function

' Returns the primo date of the tertiayear of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateThisTertiayearPrimo( _
    ByVal DateThisTertiayear As Date) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = 0
    Interval = IntervalSetting(DtInterval.dtTertiayear, True)
    
    ResultDate = DateIntervalPrimo(Interval, Number, DateThisTertiayear)
    
    DateThisTertiayearPrimo = ResultDate
    
End Function

' Returns the ultimo date of the tertiayear of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateThisTertiayearUltimo( _
    ByVal DateThisTertiayear As Date) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = 0
    Interval = IntervalSetting(DtInterval.dtTertiayear, True)
    
    ResultDate = DateIntervalUltimo(Interval, Number, DateThisTertiayear)
    
    DateThisTertiayearUltimo = ResultDate
    
End Function

' Returns the primo date of the week of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateThisWeekPrimo( _
    ByVal DateThisWeek As Date, _
    Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbSunday) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = 0
    Interval = IntervalSetting(DtInterval.dtWeek)
    
    ResultDate = DateIntervalPrimo(Interval, Number, DateThisWeek, FirstDayOfWeek)
    
    DateThisWeekPrimo = ResultDate
    
End Function

' Returns the ultimo date of the week of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateThisWeekUltimo( _
    ByVal DateThisWeek As Date, _
    Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbSunday) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = 0
    Interval = IntervalSetting(DtInterval.dtWeek)
    
    ResultDate = DateIntervalUltimo(Interval, Number, DateThisWeek, FirstDayOfWeek)
    
    DateThisWeekUltimo = ResultDate
    
End Function

' Returns the primo date of the year of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateThisYearPrimo( _
    ByVal DateThisYear As Date) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = 0
    Interval = IntervalSetting(DtInterval.dtYear)
    
    ResultDate = DateIntervalPrimo(Interval, Number, DateThisYear)
    
    DateThisYearPrimo = ResultDate
    
End Function

' Returns the ultimo date of the year of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateThisYearUltimo( _
    ByVal DateThisYear As Date) _
    As Date

    Dim Interval    As String
    Dim Number      As Double
    Dim ResultDate  As Date
    
    Number = 0
    Interval = IntervalSetting(DtInterval.dtYear)
    
    ResultDate = DateIntervalUltimo(Interval, Number, DateThisYear)
    
    DateThisYearUltimo = ResultDate
    
End Function

' Calculates the date of the occurrence of Weekday in the month of DateInMonth.
'
' If Occurrence is 0 or negative, the first occurrence of Weekday in the month is assumed.
' If Occurrence is 5 or larger, the last occurrence of Weekday in the month is assumed.
'
' If Weekday is invalid or not specified, the weekday of DateInMonth is used.
'
' 2019-12-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateWeekdayInMonth( _
    ByVal DateInMonth As Date, _
    Optional ByVal Occurrence As Integer, _
    Optional ByVal Weekday As VbDayOfWeek = vbUseSystemDayOfWeek) _
    As Date
    
    Dim Offset          As Integer
    Dim Month           As Integer
    Dim Year            As Integer
    Dim ResultDate      As Date
    
    ' Validate Weekday.
    Select Case Weekday
        Case _
            vbMonday, _
            vbTuesday, _
            vbWednesday, _
            vbThursday, _
            vbFriday, _
            vbSaturday, _
            vbSunday
        Case Else
            ' vbUseSystemDayOfWeek, zero, none or invalid value for VbDayOfWeek.
            Weekday = VBA.Weekday(DateInMonth)
    End Select
    
    ' Validate Occurence.
    If Occurrence < 1 Then
        ' Find first occurrence.
        Occurrence = 1
    ElseIf Occurrence > MaxWeekdayCountInMonth Then
        ' Find last occurrence.
        Occurrence = MaxWeekdayCountInMonth
    End If
    
    ' Start date.
    Month = VBA.Month(DateInMonth)
    Year = VBA.Year(DateInMonth)
    ResultDate = DateSerial(Year, Month, 1)
    
    ' Find offset of Weekday from first day of month.
    Offset = DaysPerWeek * (Occurrence - 1) + (Weekday - VBA.Weekday(ResultDate) + DaysPerWeek) Mod DaysPerWeek
    ' Calculate result date.
    ResultDate = DateAdd("d", Offset, ResultDate)
    
    If Occurrence = MaxWeekdayCountInMonth Then
        ' The latest occurrency of Weekday is requested.
        ' Check if there really is a fifth occurrence of Weekday in this month.
        If VBA.Month(ResultDate) <> Month Then
            ' There are only four occurrencies of Weekday in this month.
            ' Return the fourth as the latest.
            ResultDate = DateAdd("d", -DaysPerWeek, ResultDate)
        End If
    End If
    
    DateWeekdayInMonth = ResultDate
  
End Function

' Calculates the date of the first occurrence of DayOfWeek in the month of DateInMonth.
'
' 2017-01-21. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateWeekdayInMonthFirst( _
    ByVal DateInMonth As Date, _
    Optional ByVal DayOfWeek As VbDayOfWeek = VbDayOfWeek.vbUseSystemDayOfWeek) _
    As Date
    
    Const FirstOccurrence   As Integer = 0
    
    Dim ResultDate  As Date
    
    ResultDate = DateWeekdayInMonth(DateInMonth, FirstOccurrence, DayOfWeek)
    
    DateWeekdayInMonthFirst = ResultDate
  
End Function

' Calculates the date of the last occurrence of DayOfWeek in the month of DateInMonth.
'
' 2017-01-21. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateWeekdayInMonthLast( _
    ByVal DateInMonth As Date, _
    Optional ByVal DayOfWeek As VbDayOfWeek = VbDayOfWeek.vbUseSystemDayOfWeek) _
    As Date
    
    Const LastOccurrence    As Integer = 5
    
    Dim ResultDate  As Date
    
    ResultDate = DateWeekdayInMonth(DateInMonth, LastOccurrence, DayOfWeek)
    
    DateWeekdayInMonthLast = ResultDate
  
End Function

' Calculates the date of DayOfWeek in the week of DateInWeek.
' By default, the returned date is the first day in the week
' as defined by the current Windows settings.
'
' Optionally, parameter DayOfWeek can be specified to return
' any other weekday of the week.
' Further, parameter FirstDayOfWeek can be specified to select
' any other weekday as the first weekday of a week.
'
' Limitation:
' For the first and the last week of the range of Date, some
' combinations of DayOfWeek and FirstDayOfWeek that would result
' in dates outside the range of Date, will raise an overflow error.
'
' 2017-05-03. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateWeekdayInWeek( _
    ByVal DateInWeek As Date, _
    Optional ByVal DayOfWeek As VbDayOfWeek = VbDayOfWeek.vbUseSystemDayOfWeek, _
    Optional ByVal FirstDayOfWeek As VbDayOfWeek = VbDayOfWeek.vbUseSystemDayOfWeek) _
    As Date
    
    Dim DayInWeek   As VbDayOfWeek
    Dim OffsetZero  As Integer
    Dim OffsetFind  As Integer
    Dim ResultDate  As Date
    
    ' Validate parameters.
    If Not IsWeekday(DayOfWeek) Then
        ' Don't accept invalid values for DayOfWeek.
        Err.Raise DtError.dtInvalidProcedureCallOrArgument
        Exit Function
    End If
    If Not IsWeekday(FirstDayOfWeek) Then
        ' Don't accept invalid values for FirstDayOfWeek.
        Err.Raise DtError.dtInvalidProcedureCallOrArgument
        Exit Function
    End If
    
    ' Apply system setting.
    If DayOfWeek = VbDayOfWeek.vbUseSystemDayOfWeek Then
        ' Find enum index of VbDayOfWeek for the requested weekday.
        'DayOfWeek = (vbSunday + Weekday(Date) - Weekday(Date, VbDayOfWeek.vbUseSystemDayOfWeek) + DaysPerWeek) Mod DaysPerWeek
        DayOfWeek = SystemDayOfWeek()
    End If
    If FirstDayOfWeek = VbDayOfWeek.vbUseSystemDayOfWeek Then
        ' Find enum index of VbDayOfWeek for the first weekday.
        'FirstDayOfWeek = (vbSunday + Weekday(Date) - Weekday(Date, VbDayOfWeek.vbUseSystemDayOfWeek) + DaysPerWeek) Mod DaysPerWeek
        FirstDayOfWeek = SystemDayOfWeek()
    End If
    
    ' Find the date of DayOfWeek.
    DayInWeek = Weekday(DateInWeek)
    ' Find the offset of the weekday of DateInWeek from the first day of the week.
    ' Will always be <= 0.
    OffsetZero = (FirstDayOfWeek - DayInWeek - DaysPerWeek) Mod DaysPerWeek
    ' Find the offset of DayOfWeek from the first day of the week.
    ' Will always be >= 0.
    OffsetFind = (DayOfWeek - FirstDayOfWeek + DaysPerWeek) Mod DaysPerWeek
    ' Calculate result date using the sum of the offset parts.
    ResultDate = DateAdd(IntervalSetting(dtDay), OffsetZero + OffsetFind, DateInWeek)
    
    DateWeekdayInWeek = ResultDate
  
End Function

' Calculates the occurrence of the weekday of DateInMonth in the month of DateInMonth and
' returns the date of the same occurrence of this weekday in the same month in the next year.
'
' If the found occurrence is five, the last occurrence of the weekday in the month of the
' next year is returned.
'
' 2017-01-21. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateWeekdayOfMonthNextYear( _
    ByVal DateInMonth As Date) _
    As Date

    Dim Occurrence  As Integer
    Dim DayOfWeek   As Integer
    Dim ResultDate  As Date
  
    ResultDate = DateAdd("yyyy", 1, DateInMonth)
    ' Occurrence of the weekday of DateInMonth in present month.
    Occurrence = WeekdayOccurrenceOfMonth(DateInMonth)
    ' Weekday of present date.
    DayOfWeek = Weekday(DateInMonth)
    
    ' Offset DateInMonth to match weekday and occurrence of weekday for the month of next year.
    ResultDate = DateWeekdayInMonth(ResultDate, Occurrence, DayOfWeek)
  
    DateWeekdayOfMonthNextYear = ResultDate
  
End Function

' Returns the first date (1. January) of the year passed.
' If Year is not passed, the current year is used.
'
' 2017-01-09. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateYearPrimo( _
    Optional ByVal Year As Integer) _
    As Date
    
    Dim ResultDate  As Date
    
    If Year = 0 Then
        ' Use year of current date.
        Year = VBA.Year(Date)
    ElseIf Not IsYear(Year) Then
        ' Don't accept years outside the range of Date.
        Err.Raise DtError.dtOverflow
        Exit Function
    End If
    
    ResultDate = DateSerial(Year, MinMonthValue, MinDayValue)
    
    DateYearPrimo = ResultDate
    
End Function

' Returns the last date (31. December) of the year passed.
' If Year is not passed, the current year is used.
'
' 2017-01-09. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateYearUltimo( _
    Optional ByVal Year As Integer) _
    As Date
    
    Dim ResultDate  As Date
    
    If Year = 0 Then
        ' Use year of current date.
        Year = VBA.Year(Date)
    ElseIf Not IsYear(Year) Then
        ' Don't accept years outside the range of Date.
        Err.Raise DtError.dtOverflow
        Exit Function
    End If
    
    ResultDate = DateSerial(Year, MaxMonthValue, MaxDayValue)
    
    DateYearUltimo = ResultDate
    
End Function

' Returns the date of Monday for the ISO 8601 week of IsoYear and Week.
' Optionally, returns the date of any other weekday of that week.
'
' 2017-05-03. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateYearWeek( _
    ByVal IsoWeek As Integer, _
    Optional ByVal IsoYear As Integer, _
    Optional ByVal DayOfWeek As VbDayOfWeek = VbDayOfWeek.vbMonday) _
    As Date
    
    Dim WeekDate    As Date
    Dim ResultDate  As Date
    
    If IsoYear = 0 Then
        IsoYear = Year(Date)
    End If
    
    ' Validate parameters.
    If Not IsWeekday(DayOfWeek) Then
        ' Don't accept invalid values for DayOfWeek.
        Err.Raise DtError.dtInvalidProcedureCallOrArgument
        Exit Function
    End If
    If Not IsWeek(IsoWeek, IsoYear) Then
        ' A valid week number must be passed.
        Err.Raise DtError.dtInvalidProcedureCallOrArgument
        Exit Function
    End If
    
    WeekDate = DateAdd(IntervalSetting(dtWeek), IsoWeek - 1, DateFirstWeekYear(IsoYear))
    ResultDate = DateThisWeekPrimo(WeekDate, DayOfWeek)
    
    DateYearWeek = ResultDate

End Function

' Returns the weekday of the first day of the week according to the current Windows settings.
'
' 2017-05-03. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function SystemDayOfWeek() As VbDayOfWeek

    Const DateOfSaturday    As Date = #12:00:00 AM#
    
    Dim DayOfWeek   As VbDayOfWeek
    
    DayOfWeek = vbSunday + vbSaturday - Weekday(DateOfSaturday, vbUseSystemDayOfWeek)
    
    SystemDayOfWeek = DayOfWeek
    
End Function

' Returns the system setting for VbFirstWeekOfYear.
'
' 2019-01-14. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function SystemFirstWeekOfYear() As VbFirstWeekOfYear

    Const LocaleUserDefault         As Long = &H400
    Const LocaleIFirstWeekOfYear    As Long = &H100D
    Const BufferLength              As Long = 256

    Const DefaultFirstWeekOfYear    As Long = VbFirstWeekOfYear.vbUseSystem

    Dim Locale          As Long
    Dim LocaleType      As Long
    Dim Buffer          As String
    Dim Result          As Long
    Dim FirstWeekOfYear As VbFirstWeekOfYear

    Locale = LocaleUserDefault
    LocaleType = LocaleIFirstWeekOfYear
    Buffer = String(BufferLength, 0)
    
    Result = GetLocaleInfo(Locale, LocaleType, Buffer, BufferLength)
    If Result > 0 Then
        ' Convert API return values to those of the VBA enumeration.
        Select Case Val(Left(Buffer, Result - 1))
            Case 0
                ' Week containing 1/1 is the first week of the year.
                FirstWeekOfYear = vbFirstJan1
            Case 1
                ' First full week following 1/1 is the first week of the year.
                FirstWeekOfYear = vbFirstFullWeek
            Case 2
                ' First week containing at least four days is the first week of the year.
                FirstWeekOfYear = vbFirstFourDays
        End Select
    Else
        FirstWeekOfYear = DefaultFirstWeekOfYear
    End If
    
    SystemFirstWeekOfYear = FirstWeekOfYear

End Function


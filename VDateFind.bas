Attribute VB_Name = "VDateFind"
Option Explicit
'
' VDateFind
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
'   DateFind
'   DateMsec
'   VDateCore
'

' Returns a date added a number of months.
' Result date is identical to the result of DateAdd("m", Number, Date1)
' except that if an ultimo date of a month is passed, an ultimo date will
' always be returned as well.
' Optionally, days of the 30th (or 28th of February of leap years) will
' also be regarded as ultimo.
' Returns Null if Date1 is Null.
'
' 2015-11-30. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateAddMonth( _
    ByVal Date1 As Variant, _
    Number As Variant, _
    Optional Include2830 As Variant = False) _
    As Variant
    
    Dim DateNext    As Variant
    
    If IsDateExt(Date1) Then
        DateNext = DateAddMonth(CDate(Date1), Int(Val(Nz(Number))), Val(Nz(Include2830)))
    Else
        DateNext = Null
    End If
    
    VDateAddMonth = DateNext

End Function

' Calculates next annual day following SomeDate based on AnnualDay.
' Calculates correctly for leap years if AnnualDay is Feb. 29th.
' If SomeDate is earlier than AnnualDay, AnnualDay is returned.
' If next annual day should be later than 9999-12-31, the annual day
' of year 9999 is returned.
' Returns Null if AnnualDay is Null.
'
' 2015-11-21. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function VDateNextAnnualDay( _
    ByVal AnnualDay As Variant, _
    Optional ByVal SomeDate As Variant) _
    As Variant
    
    Dim NextAnnualDay   As Variant
    
    If IsDateExt(AnnualDay) Then
        NextAnnualDay = DateNextAnnualDay(CDate(AnnualDay), SomeDate)
    Else
        NextAnnualDay = Null
    End If
    
    VDateNextAnnualDay = NextAnnualDay
    
End Function

' Returns the earliest date and time of the date following the date passed.
' Returns Null if DateThisDay is Null.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateNextDayPrimo( _
    ByVal DateThisDay As Variant) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = 1
    Interval = IntervalSetting(DtInterval.dtDay)
    
    Result = VDateIntervalPrimo(Interval, Number, DateThisDay)
    
    VDateNextDayPrimo = Result
    
End Function

' Returns the latest date and time of the date following the date passed.
' Returns Null if DateThisDay is Null.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateNextDayUltimo( _
    ByVal DateThisDay As Variant) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = 1
    Interval = IntervalSetting(DtInterval.dtDay)
    
    Result = VDateIntervalUltimo(Interval, Number, DateThisDay)
    
    VDateNextDayUltimo = Result
    
End Function

' Returns the primo date of the month following the month of the date passed.
' Returns Null if DateThisMonth is Null.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateNextMonthPrimo( _
    ByVal DateThisMonth As Variant) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = 1
    Interval = IntervalSetting(DtInterval.dtMonth)
    
    Result = VDateIntervalPrimo(Interval, Number, DateThisMonth)
    
    VDateNextMonthPrimo = Result
    
End Function

' Returns the ultimo date of the month following the month of the date passed.
' Returns Null if DateThisMonth is Null.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateNextMonthUltimo( _
    ByVal DateThisMonth As Variant) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = 1
    Interval = IntervalSetting(DtInterval.dtMonth)
    
    Result = VDateIntervalUltimo(Interval, Number, DateThisMonth)
    
    VDateNextMonthUltimo = Result
    
End Function

' Returns the primo date of the quarter following the quarter of the date passed.
' Returns Null if DateThisQuarter is Null.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateNextQuarterPrimo( _
    ByVal DateThisQuarter As Variant) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = 1
    Interval = IntervalSetting(DtInterval.dtQuarter)
    
    Result = VDateIntervalPrimo(Interval, Number, DateThisQuarter)
    
    VDateNextQuarterPrimo = Result
    
End Function

' Returns the ultimo date of the quarter following the quarter of the date passed.
' Returns Null if DateThisQuarter is Null.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateNextQuarterUltimo( _
    ByVal DateThisQuarter As Variant) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = 1
    Interval = IntervalSetting(DtInterval.dtQuarter)
    
    Result = VDateIntervalUltimo(Interval, Number, DateThisQuarter)
    
    VDateNextQuarterUltimo = Result
    
End Function

' Returns the primo date of the semiyear following the semiyear of the date passed.
' Returns Null if DateThisSemiyear is Null.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateNextSemiyearPrimo( _
    ByVal DateThisSemiyear As Variant) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = 1
    Interval = IntervalSetting(DtInterval.dtSemiyear, True)
    
    Result = VDateIntervalPrimo(Interval, Number, DateThisSemiyear)
    
    VDateNextSemiyearPrimo = Result
    
End Function

' Returns the ultimo date of the semiyear following the semiyear of the date passed.
' Returns Null if DateThisSemiyear is Null.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateNextSemiyearUltimo( _
    ByVal DateThisSemiyear As Variant) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = 1
    Interval = IntervalSetting(DtInterval.dtSemiyear, True)
    
    Result = VDateIntervalUltimo(Interval, Number, DateThisSemiyear)
    
    VDateNextSemiyearUltimo = Result
    
End Function

' Returns the primo date of the sextayear following the sextayear of the date passed.
' Returns Null if DateThisSextayear is Null.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateNextSextayearPrimo( _
    ByVal DateThisSextayear As Variant) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = 1
    Interval = IntervalSetting(DtInterval.dtSextayear, True)
    
    Result = VDateIntervalPrimo(Interval, Number, DateThisSextayear)
    
    VDateNextSextayearPrimo = Result
    
End Function

' Returns the ultimo date of the sextayear following the sextayear of the date passed.
' Returns Null if DateThisSextayear is Null.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateNextSextayearUltimo( _
    ByVal DateThisSextayear As Variant) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = 1
    Interval = IntervalSetting(DtInterval.dtSextayear, True)
    
    Result = VDateIntervalUltimo(Interval, Number, DateThisSextayear)
    
    VDateNextSextayearUltimo = Result
    
End Function

' Returns the primo date of the tertiayear following the tertiayear of the date passed.
' Returns Null if DateThisTertiayear is Null.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateNextTertiayearPrimo( _
    ByVal DateThisTertiayear As Variant) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = 1
    Interval = IntervalSetting(DtInterval.dtTertiayear, True)
    
    Result = VDateIntervalPrimo(Interval, Number, DateThisTertiayear)
    
    VDateNextTertiayearPrimo = Result
    
End Function

' Returns the ultimo date of the tertiayear following the tertiayear of the date passed.
' Returns Null if DateThisTertiayear is Null.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateNextTertiayearUltimo( _
    ByVal DateThisTertiayear As Variant) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = 1
    Interval = IntervalSetting(DtInterval.dtTertiayear, True)
    
    Result = VDateIntervalUltimo(Interval, Number, DateThisTertiayear)
    
    VDateNextTertiayearUltimo = Result
    
End Function

' Returns the primo date of the week following the week of the date passed.
' Returns Null if DateThisWeek is Null or parameters are invalid.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateNextWeekPrimo( _
    ByVal DateThisWeek As Variant, _
    Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbSunday) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = 1
    Interval = IntervalSetting(DtInterval.dtWeek)
    
    Result = VDateIntervalPrimo(Interval, Number, DateThisWeek, FirstDayOfWeek)
    
    VDateNextWeekPrimo = Result
    
End Function

' Returns the ultimo date of the week following the week of the date passed.
' Returns Null if DateThisWeek is Null or parameters are invalid.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateNextWeekUltimo( _
    ByVal DateThisWeek As Variant, _
    Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbSunday) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = 1
    Interval = IntervalSetting(DtInterval.dtWeek)
    
    Result = VDateIntervalUltimo(Interval, Number, DateThisWeek, FirstDayOfWeek)
    
    VDateNextWeekUltimo = Result
    
End Function

' Returns the primo date of the year following the year of the date passed.
' Returns Null if DateThisYear is Null.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateNextYearPrimo( _
    ByVal DateThisYear As Variant) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = 1
    Interval = IntervalSetting(DtInterval.dtYear)
    
    Result = VDateIntervalPrimo(Interval, Number, DateThisYear)
    
    VDateNextYearPrimo = Result
    
End Function

' Returns the ultimo date of the year following the year of the date passed.
' Returns Null if DateThisYear is Null.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateNextYearUltimo( _
    ByVal DateThisYear As Variant) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = 1
    Interval = IntervalSetting(DtInterval.dtYear)
    
    Result = VDateIntervalUltimo(Interval, Number, DateThisYear)
    
    VDateNextYearUltimo = Result
    
End Function

' Returns the earliest date and time of the date preceding the date passed.
' Returns Null if DateThisDay is Null.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDatePreviousDayPrimo( _
    ByVal DateThisDay As Variant) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = -1
    Interval = IntervalSetting(DtInterval.dtDay)
    
    Result = VDateIntervalPrimo(Interval, Number, DateThisDay)
    
    VDatePreviousDayPrimo = Result
    
End Function

' Returns the latest date and time of the date preceding the date passed.
' Returns Null if DateThisDay is Null.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDatePreviousDayUltimo( _
    ByVal DateThisDay As Variant) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = -1
    Interval = IntervalSetting(DtInterval.dtDay)
    
    Result = VDateIntervalUltimo(Interval, Number, DateThisDay)
    
    VDatePreviousDayUltimo = Result
    
End Function

' Returns the primo date of the month preceding the month of the date passed.
' Returns Null if DateThisMonth is Null.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDatePreviousMonthPrimo( _
    ByVal DateThisMonth As Variant) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = -1
    Interval = IntervalSetting(DtInterval.dtMonth)
    
    Result = VDateIntervalPrimo(Interval, Number, DateThisMonth)
    
    VDatePreviousMonthPrimo = Result
    
End Function

' Returns the ultimo date of the month preceding the month of the date passed.
' Returns Null if DateThisMonth is Null.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDatePreviousMonthUltimo( _
    ByVal DateThisMonth As Variant) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = -1
    Interval = IntervalSetting(DtInterval.dtMonth)
    
    Result = VDateIntervalUltimo(Interval, Number, DateThisMonth)
    
    VDatePreviousMonthUltimo = Result
    
End Function

' Returns the primo date of the quarter preceding the quarter of the date passed.
' Returns Null if DateThisQuarter is Null.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDatePreviousQuarterPrimo( _
    ByVal DateThisQuarter As Variant) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = -1
    Interval = IntervalSetting(DtInterval.dtQuarter)
    
    Result = VDateIntervalPrimo(Interval, Number, DateThisQuarter)
    
    VDatePreviousQuarterPrimo = Result
    
End Function

' Returns the ultimo date of the quarter preceding the quarter of the date passed.
' Returns Null if DateThisQuarter is Null.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDatePreviousQuarterUltimo( _
    ByVal DateThisQuarter As Variant) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = -1
    Interval = IntervalSetting(DtInterval.dtQuarter)
    
    Result = VDateIntervalUltimo(Interval, Number, DateThisQuarter)
    
    VDatePreviousQuarterUltimo = Result
    
End Function

' Returns the primo date of the semiyear preceding the semiyear of the date passed.
' Returns Null if DateThisSemiyear is Null.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDatePreviousSemiyearPrimo( _
    ByVal DateThisSemiyear As Variant) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = -1
    Interval = IntervalSetting(DtInterval.dtSemiyear, True)
    
    Result = VDateIntervalPrimo(Interval, Number, DateThisSemiyear)
    
    VDatePreviousSemiyearPrimo = Result
    
End Function

' Returns the ultimo date of the semiyear preceding the semiyear of the date passed.
' Returns Null if DateThisSemiyear is Null.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDatePreviousSemiyearUltimo( _
    ByVal DateThisSemiyear As Variant) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = -1
    Interval = IntervalSetting(DtInterval.dtSemiyear, True)
    
    Result = VDateIntervalUltimo(Interval, Number, DateThisSemiyear)
    
    VDatePreviousSemiyearUltimo = Result
    
End Function

' Returns the primo date of the sextayear preceding the sextayear of the date passed.
' Returns Null if DateThisSextayear is Null.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDatePreviousSextayearPrimo( _
    ByVal DateThisSextayear As Variant) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = -1
    Interval = IntervalSetting(DtInterval.dtSextayear, True)
    
    Result = VDateIntervalPrimo(Interval, Number, DateThisSextayear)
    
    VDatePreviousSextayearPrimo = Result
    
End Function

' Returns the ultimo date of the sextayear preceding the sextayear of the date passed.
' Returns Null if DateThisSextayear is Null.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDatePreviousSextayearUltimo( _
    ByVal DateThisSextayear As Variant) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = -1
    Interval = IntervalSetting(DtInterval.dtSextayear, True)
    
    Result = VDateIntervalUltimo(Interval, Number, DateThisSextayear)
    
    VDatePreviousSextayearUltimo = Result
    
End Function

' Returns the primo date of the tertiayear preceding the tertiayear of the date passed.
' Returns Null if DateThisTertiayear is Null.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDatePreviousTertiayearPrimo( _
    ByVal DateThisTertiayear As Variant) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = -1
    Interval = IntervalSetting(DtInterval.dtTertiayear, True)
    
    Result = VDateIntervalPrimo(Interval, Number, DateThisTertiayear)
    
    VDatePreviousTertiayearPrimo = Result
    
End Function

' Returns the ultimo date of the tertiayear preceding the tertiayear of the date passed.
' Returns Null if DateThisTertiayear is Null.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDatePreviousTertiayearUltimo( _
    ByVal DateThisTertiayear As Variant) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = -1
    Interval = IntervalSetting(DtInterval.dtTertiayear, True)
    
    Result = VDateIntervalUltimo(Interval, Number, DateThisTertiayear)
    
    VDatePreviousTertiayearUltimo = Result
    
End Function

' Returns the primo date of the week preceding the week of the date passed.
' Returns Null if DateThisWeek is Null or parameters are invalid.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDatePreviousWeekPrimo( _
    ByVal DateThisWeek As Variant, _
    Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbSunday) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = -1
    Interval = IntervalSetting(DtInterval.dtWeek)
    
    Result = VDateIntervalPrimo(Interval, Number, DateThisWeek, FirstDayOfWeek)
    
    VDatePreviousWeekPrimo = Result
    
End Function

' Returns the ultimo date of the week preceding the week of the date passed.
' Returns Null if DateThisWeek is Null or parameters are invalid.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDatePreviousWeekUltimo( _
    ByVal DateThisWeek As Variant, _
    Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbSunday) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = -1
    Interval = IntervalSetting(DtInterval.dtWeek)
    
    Result = VDateIntervalUltimo(Interval, Number, DateThisWeek, FirstDayOfWeek)
    
    VDatePreviousWeekUltimo = Result
    
End Function

' Returns the primo date of the year preceding the year of the date passed.
' Returns Null if DateThisYear is Null.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDatePreviousYearPrimo( _
    ByVal DateThisYear As Variant) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = -1
    Interval = IntervalSetting(DtInterval.dtYear)
    
    Result = VDateIntervalPrimo(Interval, Number, DateThisYear)
    
    VDatePreviousYearPrimo = Result
    
End Function

' Returns the ultimo date of the year preceding the year of the date passed.
' Returns Null if DateThisYear is Null.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDatePreviousYearUltimo( _
    ByVal DateThisYear As Variant) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = -1
    Interval = IntervalSetting(DtInterval.dtYear)
    
    Result = VDateIntervalUltimo(Interval, Number, DateThisYear)
    
    VDatePreviousYearUltimo = Result
    
End Function

' Returns the earliest date and time of the date passed.
' Returns Null if DateThisDay is Null.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateThisDayPrimo( _
    ByVal DateThisDay As Variant) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = 0
    Interval = IntervalSetting(DtInterval.dtDay)
    
    Result = VDateIntervalPrimo(Interval, Number, DateThisDay)
    
    VDateThisDayPrimo = Result
    
End Function

' Returns the latest date and time of the date passed.
' Returns Null if DateThisDay is Null.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateThisDayUltimo( _
    ByVal DateThisDay As Variant) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = 0
    Interval = IntervalSetting(DtInterval.dtDay)
    
    Result = VDateIntervalUltimo(Interval, Number, DateThisDay)
    
    VDateThisDayUltimo = Result
    
End Function

' Returns the primo date of the month of the date passed.
' Returns Null if DateThisMonth is Null.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateThisMonthPrimo( _
    ByVal DateThisMonth As Variant) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = 0
    Interval = IntervalSetting(DtInterval.dtMonth)
    
    Result = VDateIntervalPrimo(Interval, Number, DateThisMonth)
    
    VDateThisMonthPrimo = Result
    
End Function

' Returns the ultimo date of the month of the date passed.
' Returns Null if DateThisMonth is Null.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateThisMonthUltimo( _
    ByVal DateThisMonth As Variant) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = 0
    Interval = IntervalSetting(DtInterval.dtMonth)
    
    Result = VDateIntervalUltimo(Interval, Number, DateThisMonth)
    
    VDateThisMonthUltimo = Result
    
End Function

' Returns the primo date of the quarter of the date passed.
' Returns Null if DateThisQuarter is Null.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateThisQuarterPrimo( _
    ByVal DateThisQuarter As Variant) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = 0
    Interval = IntervalSetting(DtInterval.dtQuarter)
    
    Result = VDateIntervalPrimo(Interval, Number, DateThisQuarter)
    
    VDateThisQuarterPrimo = Result
    
End Function

' Returns the ultimo date of the quarter of the date passed.
' Returns Null if DateThisQuarter is Null.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateThisQuarterUltimo( _
    ByVal DateThisQuarter As Variant) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = 0
    Interval = IntervalSetting(DtInterval.dtQuarter)
    
    Result = VDateIntervalUltimo(Interval, Number, DateThisQuarter)
    
    VDateThisQuarterUltimo = Result
    
End Function

' Returns the primo date of the semiyear of the date passed.
' Returns Null if DateThisSemiyear is Null.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateThisSemiyearPrimo( _
    ByVal DateThisSemiyear As Variant) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = 0
    Interval = IntervalSetting(DtInterval.dtSemiyear, True)
    
    Result = VDateIntervalPrimo(Interval, Number, DateThisSemiyear)
    
    VDateThisSemiyearPrimo = Result
    
End Function

' Returns the ultimo date of the semiyear of the date passed.
' Returns Null if DateThisSemiyear is Null.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateThisSemiyearUltimo( _
    ByVal DateThisSemiyear As Variant) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = 0
    Interval = IntervalSetting(DtInterval.dtSemiyear, True)
    
    Result = VDateIntervalUltimo(Interval, Number, DateThisSemiyear)
    
    VDateThisSemiyearUltimo = Result
    
End Function

' Returns the primo date of the sextayear of the date passed.
' Returns Null if DateThisSextayear is Null.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateThisSextayearPrimo( _
    ByVal DateThisSextayear As Variant) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = 0
    Interval = IntervalSetting(DtInterval.dtSextayear, True)
    
    Result = VDateIntervalPrimo(Interval, Number, DateThisSextayear)
    
    VDateThisSextayearPrimo = Result
    
End Function

' Returns the ultimo date of the sextayear of the date passed.
' Returns Null if DateThisSextayear is Null.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateThisSextayearUltimo( _
    ByVal DateThisSextayear As Variant) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = 0
    Interval = IntervalSetting(DtInterval.dtSextayear, True)
    
    Result = VDateIntervalUltimo(Interval, Number, DateThisSextayear)
    
    VDateThisSextayearUltimo = Result
    
End Function

' Returns the primo date of the tertiayear of the date passed.
' Returns Null if DateThisTertiayear is Null.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateThisTertiayearPrimo( _
    ByVal DateThisTertiayear As Variant) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = 0
    Interval = IntervalSetting(DtInterval.dtTertiayear, True)
    
    Result = VDateIntervalPrimo(Interval, Number, DateThisTertiayear)
    
    VDateThisTertiayearPrimo = Result
    
End Function

' Returns the ultimo date of the tertiayear of the date passed.
' Returns Null if DateThisTertiayear is Null.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateThisTertiayearUltimo( _
    ByVal DateThisTertiayear As Variant) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = 0
    Interval = IntervalSetting(DtInterval.dtTertiayear, True)
    
    Result = VDateIntervalUltimo(Interval, Number, DateThisTertiayear)
    
    VDateThisTertiayearUltimo = Result
    
End Function

' Returns the primo date of the week of the date passed.
' Returns Null if DateThisWeek is Null or parameters are invalid.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateThisWeekPrimo( _
    ByVal DateThisWeek As Variant, _
    Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbSunday) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = 0
    Interval = IntervalSetting(DtInterval.dtWeek)
    
    Result = VDateIntervalPrimo(Interval, Number, DateThisWeek, FirstDayOfWeek)
    
    VDateThisWeekPrimo = Result
    
End Function

' Returns the ultimo date of the week of the date passed.
' Returns Null if DateThisWeek is Null or parameters are invalid.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateThisWeekUltimo( _
    ByVal DateThisWeek As Variant, _
    Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbSunday) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = 0
    Interval = IntervalSetting(DtInterval.dtWeek)
    
    Result = VDateIntervalUltimo(Interval, Number, DateThisWeek, FirstDayOfWeek)
    
    VDateThisWeekUltimo = Result
    
End Function

' Returns the primo date of the year of the date passed.
' Returns Null if DateThisYear is Null.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateThisYearPrimo( _
    ByVal DateThisYear As Variant) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = 0
    Interval = IntervalSetting(DtInterval.dtYear)
    
    Result = VDateIntervalPrimo(Interval, Number, DateThisYear)
    
    VDateThisYearPrimo = Result
    
End Function

' Returns the ultimo date of the year of the date passed.
' Returns Null if DateThisYear is Null.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateThisYearUltimo( _
    ByVal DateThisYear As Variant) _
    As Variant

    Dim Interval    As String
    Dim Number      As Double
    Dim Result      As Variant
    
    Number = 0
    Interval = IntervalSetting(DtInterval.dtYear)
    
    Result = VDateIntervalUltimo(Interval, Number, DateThisYear)
    
    VDateThisYearUltimo = Result
    
End Function

' Calculates the date of the occurrence of DayOfWeek in the month of DateInMonth.
' If DateInMonth is Null or an invalid value, Null is returned.
'
' If Occurrence is 0 or negative, the first occurrence of DayOfWeek in the month is assumed.
' If Occurrence is 5 or larger, the last occurrence of DayOfWeek in the month is assumed.
'
' If DayOfWeek is invalid or not specified, the weekday of DateInMonth is used.
'
' 2016-06-10. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateWeekdayInMonth( _
    ByVal DateInMonth As Variant, _
    Optional ByVal Occurrence As Integer, _
    Optional ByVal DayOfWeek As VbDayOfWeek = VbDayOfWeek.vbUseSystemDayOfWeek) _
    As Variant
    
    Dim ResultDate  As Variant
    
    ResultDate = Null
    
    If IsDateExt(DateInMonth) Then
        If DateDiff(IntervalSetting(DtInterval.dtYear), DateInMonth, MaxDateValue) > 0 Then
            ResultDate = DateWeekdayInMonth(CDate(DateInMonth), Occurrence, DayOfWeek)
        End If
    End If
    
    VDateWeekdayInMonth = ResultDate

End Function

' Calculates the date of the first occurrence of DayOfWeek in the month of DateInMonth.
' If DateInMonth is Null or an invalid value, Null is returned.
'
' If DayOfWeek is invalid or not specified, the weekday of DateInMonth is used.
'
' 2017-01-21. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateWeekdayInMonthFirst( _
    ByVal DateInMonth As Variant, _
    Optional ByVal DayOfWeek As VbDayOfWeek = VbDayOfWeek.vbUseSystemDayOfWeek) _
    As Variant
    
    Const FirstOccurrence   As Integer = 0
    
    Dim ResultDate  As Variant
    
    ResultDate = VDateWeekdayInMonth(DateInMonth, FirstOccurrence, DayOfWeek)
    
    VDateWeekdayInMonthFirst = ResultDate
  
End Function

' Calculates the date of the last occurrence of DayOfWeek in the month of DateInMonth.
' If DateInMonth is Null or an invalid value, Null is returned.
'
' If DayOfWeek is invalid or not specified, the weekday of DateInMonth is used.
'
' 2017-01-21. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateWeekdayInMonthLast( _
    ByVal DateInMonth As Variant, _
    Optional ByVal DayOfWeek As VbDayOfWeek = VbDayOfWeek.vbUseSystemDayOfWeek) _
    As Variant
    
    Const LastOccurrence    As Integer = 5
    
    Dim ResultDate  As Variant
    
    ResultDate = VDateWeekdayInMonth(DateInMonth, LastOccurrence, DayOfWeek)
    
    VDateWeekdayInMonthLast = ResultDate
  
End Function

' Calculates the date of DayOfWeek in the week of DateInWeek.
' By default, the returned date is the first day in the week
' as defined by the current Windows settings.
' If DateInWeek is Null or an invalid value, Null is returned.
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
Public Function VDateWeekdayInWeek( _
    ByVal DateInWeek As Variant, _
    Optional ByVal DayOfWeek As VbDayOfWeek = VbDayOfWeek.vbUseSystemDayOfWeek, _
    Optional ByVal FirstDayOfWeek As VbDayOfWeek = VbDayOfWeek.vbUseSystemDayOfWeek) _
    As Variant
    
    Dim Result  As Variant
    
    If IsDateExt(DateInWeek) Then
        Result = DateWeekdayInWeek(CDate(DateInWeek), DayOfWeek, FirstDayOfWeek)
    Else
        Result = Null
    End If
        
    VDateWeekdayInWeek = Result
  
End Function

' Calculates the occurrence of the weekday of DateInMonth in the month of DateInMonth and
' returns the date of the same occurrence of this weekday in the same month in the next year.
' If DateInMonth is Null or an invalid value, Null is returned.
'
' If the found occurrence is five, the last occurrence of the weekday in the month of the
' next year is returned.
'
' 2017-01-21. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateWeekdayOfMonthNextYear( _
    ByVal DateInMonth As Variant) _
    As Variant

    Dim ResultDate  As Variant
  
    ResultDate = Null
    
    If IsDateExt(DateInMonth) Then
        If DateDiff(IntervalSetting(DtInterval.dtYear), DateInMonth, MaxDateValue) > 0 Then
            ResultDate = DateWeekdayOfMonthNextYear(CDate(DateInMonth))
        End If
    End If
    
    VDateWeekdayOfMonthNextYear = ResultDate
  
End Function

' Returns the first date (1. January) of the year passed.
' If Year is not passed, the current year is used.
' If Year is Null or an invalid value, Null is returned.
'
' 2017-01-09. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateYearPrimo( _
    Optional ByVal Year As Variant = 0) _
    As Variant
    
    Dim Result  As Variant
    
    Result = Null
    
    On Error Resume Next
    Result = DateYearPrimo(Year)
    On Error GoTo 0
    
    VDateYearPrimo = Result
    
End Function

' Returns the last date (31. December) of the year passed.
' If Year is not passed, the current year is used.
' If Year is Null or an invalid value, Null is returned.
'
' 2017-01-09. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateYearUltimo( _
    Optional ByVal Year As Variant = 0) _
    As Variant
    
    Dim Result  As Variant
    
    Result = Null
    
    On Error Resume Next
    Result = DateYearUltimo(Year)
    On Error GoTo 0
    
    VDateYearUltimo = Result
    
End Function

' Returns the date of Monday for the ISO 8601 week of IsoYear and Week.
' Optionally, returns the date of any other weekday of that week.
' If Week or IsoYear is Null or an invalid value, Null is returned.
'
' 2017-05-03. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDateYearWeek( _
    ByVal IsoWeek As Variant, _
    Optional ByVal IsoYear As Variant = 0, _
    Optional ByVal DayOfWeek As VbDayOfWeek = VbDayOfWeek.vbMonday) _
    As Variant
    
    Dim Result  As Variant
    
    Result = Null
    
    If IsWeekday(DayOfWeek) Then
        If IsoYear = 0 Then
            IsoYear = Year(Date)
        End If
        If VIsWeek(IsoWeek, IsoYear) Then
            Result = DateYearWeek(IsoWeek, IsoYear, DayOfWeek)
        End If
    End If
    
    VDateYearWeek = Result

End Function


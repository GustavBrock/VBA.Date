Attribute VB_Name = "DateBase"
Option Explicit
'
' DateBase
' Version 1.4.3
'
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.Date
'
' Common constants, enums, user defined types, and basic functions
' for the entire project.
'
' License: MIT (http://opensource.org/licenses/mit-license.php)
'
' Required references:
'   None
'
' Required modules:
'   None
'


' Common constants.

    ' Date.
    Public Const MaxDateValue           As Date = #12/31/9999#
    Public Const MinDateValue           As Date = #1/1/100#
    Public Const MinDateValueMySql      As Date = #1/1/1000#
    Public Const MinDateValueSqlServer  As Date = #1/1/1753#
    Public Const ZeroDateValue          As Date = #12:00:00 AM#
    
    ' Maximum day value valid for a month.
    Public Const MaxDayValue            As Integer = 31
    ' Maximum day value valid for any month.
    Public Const MaxDayAllMonthsValue   As Integer = 28
    Public Const MinDayValue            As Integer = 1
    
    Public Const MaxWeekValue           As Integer = 53
    Public Const MinWeekValue           As Integer = 1
    Public Const MaxFortnightValue      As Integer = 27
    Public Const MinFortnightValue      As Integer = 1
    
    Public Const MaxMonthValue          As Integer = 12
    Public Const MinMonthValue          As Integer = 1
    Public Const SemimonthsPerMonth     As Integer = 2
    Public Const TertiamonthsPerMonth   As Integer = 3
    Public Const MaxSemimonthValue      As Integer = MaxMonthValue * SemimonthsPerMonth
    Public Const MinSemimonthValue      As Integer = MinMonthValue
    Public Const MaxTertiamonthValue    As Integer = MaxMonthValue * TertiamonthsPerMonth
    Public Const MinTertiamonthValue    As Integer = MinMonthValue
    Public Const DaysPerSemimonth       As Long = 30 / SemimonthsPerMonth
    Public Const DaysPerTertiamonth     As Long = 30 / TertiamonthsPerMonth
    Public Const FirstSemimonthday      As Long = MinDayValue
    Public Const SecondSemimonthday     As Long = MinDayValue + DaysPerSemimonth
    
    Public Const FirstWeekday           As Long = 1
    Public Const LastWeekday            As Long = 7
    Public Const DaysPerWeek            As Long = 7
    Public Const FortnightsPerWeek      As Long = 2
    Public Const FirstFortnightday      As Long = 1
    Public Const LastFortnightday       As Long = LastWeekday * FortnightsPerWeek
    Public Const DaysPerFortnight       As Long = DaysPerWeek * FortnightsPerWeek
    
    Public Const MaxWeekdayCountInMonth As Integer = 5
    Public Const MonthsPerYear          As Integer = 12
    Public Const SemimonthsPerYear      As Integer = MonthsPerYear * SemimonthsPerMonth
    Public Const TertiamonthsPerYear    As Integer = MonthsPerYear * TertiamonthsPerMonth
    
    Public Const YearsPerDecade         As Integer = 10
    Public Const YearsPerCentury        As Integer = 100
    Public Const YearsPerMillenium      As Integer = 1000
    
    ' Time.
    Public Const MaxTimeValue           As Date = #11:59:59 PM#
    Public Const MinTimeValue           As Date = #12:00:00 AM#
    
    Public Const HoursPerDay            As Long = 24
    Public Const MinutesPerHour         As Long = 60
    Public Const SecondsPerMinute       As Long = 60
    Public Const MillisecondsPerSecond  As Long = 10 ^ 3
    Public Const MicrosecondsPerSecond  As Long = 10 ^ 6
    Public Const NanosecondsPerSecond   As Long = 10 ^ 9
    Public Const MinutesPerDay          As Long = HoursPerDay * MinutesPerHour
    Public Const SecondsPerHour         As Long = MinutesPerHour * SecondsPerMinute
    Public Const SecondsPerDay          As Long = HoursPerDay * SecondsPerHour
    Public Const MillisecondsPerMinute  As Long = SecondsPerMinute * MillisecondsPerSecond
    Public Const MillisecondsPerDay     As Long = SecondsPerDay * MillisecondsPerSecond
    
    ' Millisecond count.
    Public Const MaxMillisecondCount    As Integer = 999
    Public Const MinMillisecondCount    As Integer = 0
    ' Millisecond values
    Public Const MaxMillisecondValue    As Date = #12:00:01 AM# / MillisecondsPerSecond * MaxMillisecondCount
    Public Const MinMillisecondValue    As Date = #12:00:00 AM#
    ' Ticks per millisecond and second.
    Public Const TicksPerMillisecond    As Long = 10 ^ 5
    Public Const TicksPerSecond         As Long = MillisecondsPerSecond * TicksPerMillisecond
    
    ' DateTime.
    Public Const MaxDateTimeValue       As Date = MaxDateValue + MaxTimeValue
    Public Const MinDateTimeValue       As Date = MinDateValue + MinTimeValue
    
    ' DateTimeMsec.
    Public Const MaxValue               As Date = MaxDateTimeValue + MaxMillisecondValue
    Public Const MinValue               As Date = MinDateTimeValue + MinMillisecondValue
    
    ' Numeric date value limits.
    Public Const MaxNumericDateValue    As Double = MaxDateTimeValue + MaxMillisecondValue
    Public Const MinNumericDateValue    As Double = MinDateValue - (MaxTimeValue + MaxMillisecondValue)
    
    ' Span.
    ' Interval with minimum one microsecond resolution.
    Public Const MaxMicrosecondDateValue    As Date = #5/18/1927#
    Public Const MinMicrosecondDateValue    As Date = #8/13/1872#
    ' Interval with minimum one nanosecond resolution.
    Public Const MaxNanosecondDateValue As Date = #1/9/1900#
    Public Const MinNanosecondDateValue As Date = #12/20/1899#
    ' Interval with minimum one tick resolution.
    Public Const MaxTickDateValue       As Date = #2:24:00 AM#
    Public Const MinTickDateValue       As Date = -#2:24:00 AM#
    ' Values.
    Public Const OneWeek                As Date = #1/6/1900#
    Public Const OneDay                 As Date = #12/31/1899#
    Public Const OneHour                As Date = #1:00:00 AM#
    Public Const OneMinute              As Date = #12:01:00 AM#
    Public Const OneSecond              As Date = #12:00:01 AM#
    Public Const OneMillisecond         As Date = OneSecond / MillisecondsPerSecond
    Public Const OneMicrosecond         As Date = OneSecond / MicrosecondsPerSecond
    Public Const OneNanosecond          As Date = OneSecond / NanosecondsPerSecond
    Public Const OneTick                As Date = OneSecond / TicksPerSecond
    
    ' Facebook flick.
    ' Reference: https://github.com/OculusVR/Flicks
    ' Interval with minimum one flick resolution.
    Public Const MaxFlickDateValue      As Date = MaxNanosecondDateValue
    Public Const MinFlickDateValue      As Date = MinNanosecondDateValue
    ' A flick is 1/705600000 second, about  1.4172335600907 ns.
    Public Const FlicksPerSecond        As Double = 705600000
    Public Const OneFlick               As Date = OneSecond / FlicksPerSecond
    
    ' Swatch Internet Time.
    ' Beats per day.
    Public Const BeatsPerDay            As Integer = 1000
    ' One .beat is 01:26.400.
    Public Const OneBeat                As Date = OneDay / BeatsPerDay
    
    ' Functions
    ' Maximum Number for DateAdd.
    Public Const MaxAddNumber           As Double = 2 ^ 31 - 1
    ' Decimal separator.
    Public Const DecimalSeparator       As String = "."
    ' ISO 8601 date separator.
    Public Const IsoDateSeparator       As String = "-"
    ' Invariant date separator.
    Public Const DateSeparator          As String = "/"
    ' Invariant time separator.
    Public Const TimeSeparator          As String = ":"
    ' Invariant millisecond separator: yyyy-mm-dd hh:nn:ss<separator>fff
    Public Const MillisecondSeparator   As String = "."
    ' Escape character for use in parameter Format of the format functions.
    Public Const EscapeCharacter        As String = "\"
    
    
' Enums.

    ' Enum for error values for use with Err.Raise.
    Public Enum DtError
        dtInvalidProcedureCallOrArgument = 5
        dtOverflow = 6
        dtTypeMismatch = 13
    End Enum
    
    ' Enum for selecting interval for calculations.
    ' If modified, be sure to modify [_First] and [_Last] as well.
    Public Enum DtInterval
        ' System value.
        [_First] = 0
        ' Native intervals.
        dtYear = 0
        dtQuarter = 1
        dtMonth = 2
        dtDayOfYear = 3
        dtDay = 4
        dtWeekday = 5
        dtWeek = 6
        dtHour = 7
        dtMinute = 8
        dtSecond = 9
        ' System value.
        [_LastNative] = 9
        ' System value.
        [_FirstExtended] = 10
        ' Extended date intervals.
        dtDimidiae = 10
        dtSemiyear = 10
        dtTertiayear = 11
        dtSextayear = 12
        dtFortnightday = 13
        dtFortnight = 14
        dtSemimonth = 15
        dtTertiamonth = 16
        dtDecade = 17
        dtCentury = 18
        dtMillenium = 19
        ' Extended time intervals.
        dtMillisecond = 20
        dtDecimalSecond = 21
        ' System value.
        [_Last] = 21
    End Enum
    
    ' Enum for count of months of intervals of full months.
    Public Enum DtIntervalMonths
        dtMonth = 1
        dtSextayear = 2
        dtQuarter = 3
        dtTertiayear = 4
        dtDimidiae = 6
        dtSemiyear = 6
        dtYear = 12
        dtDecade = 120
        dtCentury = 1200
        dtMillenium = 12000
    End Enum
    
    ' Enum for weekend days.
    Public Enum DtWeekendType
        dtMonday = vbMonday
        dtTuesday = vbTuesday
        dtWednesday = vbWednesday
        dtThursday = vbThursday
        dtFriday = vbFriday
        dtSaturday = vbSaturday
        dtSunday = vbSunday
        dtLongWeekend = 8
        dtShortWeekend = 9
        dtSabbath = 10
    End Enum
    
    ' Enum for TimeZone Daylight Saving Time.
    Public Enum TimeZoneId
        Unknown = 0
        Standard = 1
        Daylight = 2
        Invalid = &HFFFFFFFF
    End Enum
    
    
' User defined types.
'
    ' Type used in API calls.
    Public Type SystemTime
        wYear                           As Integer
        wMonth                          As Integer
        wDayOfWeek                      As Integer
        wDay                            As Integer
        wHour                           As Integer
        wMinute                         As Integer
        wSecond                         As Integer
        wMilliseconds                   As Integer
    End Type

    ' TimeZoneInformation holds information about a timezone.
    ' The two arrays are null-terminated strings, where each element
    ' holds the byte code for a character, and the last element is a
    ' null value, ASCII code 0.
    Public Type TimeZoneInformation
        Bias                            As Long
        StandardName(0 To (32 * 2 - 1)) As Byte     ' Unicode.
        StandardDate                    As SystemTime
        StandardBias                    As Long
        DaylightName(0 To (32 * 2 - 1)) As Byte     ' Unicode.
        DaylightDate                    As SystemTime
        DaylightBias                    As Long
    End Type

    ' Reference:
    '   https://msdn.microsoft.com/en-us/library/windows/desktop/ms724253(v=vs.85).aspx
    '
    ' Not used, for reference only.
    ' Complete dynamic timezone entry.
    ' Names must be Unicode arrays.
    Private Type DynamicTimeZoneInformation
        Bias                            As Long
        StandardName(0 To (32 * 2 - 1)) As Byte     ' Unicode.
        StandardDate                    As SystemTime
        StandardBias                    As Long
        DaylightName(0 To (32 * 2 - 1)) As Byte     ' Unicode.
        DaylightDate                    As SystemTime
        DaylightBias                    As Long
        TimeZoneKeyName(0 To 255)       As Byte     ' Unicode.
    End Type

    ' Type for holding separate date and time parts of a date value.
    Public Type DateTime
        ' Should always be a date at 00:00:00.000.
        Date As Date
        ' Should always be between 00:00:00 and 23:59:59 or, for
        ' milliseconds, between 00:00:00.000 and 23:59:59.999.
        Time As Date
    End Type
'

' Returns the count of months of a valid value which is a
' value that can be converted to DtInterval.
' Optionally, also returns the count for an extended value.
'
' An error is raised if an invalid value is passed.
'
' Examples:
'   Months = IntervalMonths("u", True)
'   Months -> 1200
'
'   Months = IntervalMonths(IntervalSetting(DtInterval.dtCentury, True), True)
'   Months -> 1200
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IntervalMonths( _
    ByVal Value As String, _
    Optional Extended As Boolean) _
    As Integer
    
    Dim Months  As Long
    
    If IsIntervalSetting(Value, False) Then
        Select Case IntervalValue(Value)
            Case DtInterval.dtYear
                Months = DtIntervalMonths.dtYear
            Case DtInterval.dtQuarter
                Months = DtIntervalMonths.dtQuarter
            Case DtInterval.dtMonth
                Months = DtIntervalMonths.dtMonth
        End Select
    ElseIf IsIntervalSetting(Value, Extended) Then
        Select Case IntervalValue(Value, True)
            Case DtInterval.dtDimidiae
                Months = DtIntervalMonths.dtDimidiae
            Case DtInterval.dtSemiyear
                Months = DtIntervalMonths.dtSemiyear
            Case DtInterval.dtTertiayear
                Months = DtIntervalMonths.dtTertiayear
            Case DtInterval.dtSextayear
                Months = DtIntervalMonths.dtSextayear
            Case DtInterval.dtDecade
                Months = DtIntervalMonths.dtDecade
            Case DtInterval.dtCentury
                Months = DtIntervalMonths.dtCentury
            Case DtInterval.dtMillenium
                Months = DtIntervalMonths.dtMillenium
        End Select
    Else
        Err.Raise DtError.dtInvalidProcedureCallOrArgument
        Exit Function
    End If
    
    IntervalMonths = Months
    
End Function

    
' Returns the interval setting from a value of DtInterval for use as
' the Interval parameter of DateAdd, DateDiff, and DatePart.
' Optionally, returns custom (extended) values for Interval for use in
' DateAddExt, DateDiffExt, and DatePartExt.
'
' If an invalid value is passed, an error is raised.
'
' 2019-10-11. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IntervalSetting( _
    ByVal Interval As DtInterval, _
    Optional ByVal Extended As Boolean) _
    As String

    Dim Symbol  As String
    
    Select Case Interval
        ' Native values.
        Case DtInterval.dtYear
            Symbol = "yyyy"
        Case DtInterval.dtQuarter
            Symbol = "q"
        Case DtInterval.dtMonth
            Symbol = "m"
        Case DtInterval.dtDayOfYear
            Symbol = "y"
        Case DtInterval.dtDay
            Symbol = "d"
        Case DtInterval.dtWeekday
            Symbol = "w"
        Case DtInterval.dtWeek
            Symbol = "ww"
        Case DtInterval.dtHour
            Symbol = "h"
        Case DtInterval.dtMinute
            Symbol = "n"
        Case DtInterval.dtSecond
            Symbol = "s"
        ' Extended values.
        Case DtInterval.[_FirstExtended] To DtInterval.[_Last]
            If Extended = True Then
                Select Case Interval
                    Case DtInterval.dtDimidiae, DtInterval.dtSemiyear
                        Symbol = "i"
                    Case DtInterval.dtTertiayear
                        Symbol = "r"
                    Case DtInterval.dtSextayear
                        Symbol = "g"
                    Case DtInterval.dtFortnightday
                        Symbol = "v"
                    Case DtInterval.dtFortnight
                        Symbol = "vv"
                    Case DtInterval.dtSemimonth
                        Symbol = "e"
                    Case DtInterval.dtTertiamonth
                        Symbol = "t"
                    Case DtInterval.dtDecade
                        Symbol = "x"
                    Case DtInterval.dtCentury
                        Symbol = "u"
                    Case DtInterval.dtMillenium
                        Symbol = "k"
                    Case DtInterval.dtMillisecond
                        Symbol = "f"
                    Case DtInterval.dtDecimalSecond
                        Symbol = "l"
                End Select
            End If
    End Select
    
    If Symbol = "" Then
        Err.Raise DtError.dtInvalidProcedureCallOrArgument
        Exit Function
    End If
    
    IntervalSetting = Symbol
    
End Function

' Returns the DtInterval of a valid value which is a
' value that can be converted to DtInterval.
' Optionally, also validates an extended value.
'
' The case of Value will be ignored.
'
' An error is raised if an invalid value is passed.
'
' 2021-01-06. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IntervalValue( _
    ByVal Value As String, _
    Optional ByVal Extended As Boolean) _
    As DtInterval
    
    Dim Interval    As DtInterval
    Dim Symbol      As String
    Dim Result      As Boolean
    
    ' Exit with True if Value is a valid interval setting.
    For Interval = DtInterval.[_First] To DtInterval.[_Last]
        Symbol = IntervalSetting(Interval, Extended)
        If LCase(Value) = Symbol Then
            Result = True
            Exit For
        End If
    Next
    
    If Result = False Then
        Err.Raise DtError.dtInvalidProcedureCallOrArgument
        Exit Function
    End If
    
    IntervalValue = Interval
    
End Function

    
' Returns True if Interval is passed a valid value
' of DtInterval.
' Optionally, also returns True for an extended value.
'
' 2019-11-10. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsInterval( _
    ByVal Interval As DtInterval, _
    Optional ByVal Extended As Boolean) _
    As Boolean
    
    Dim Result  As Boolean

    Select Case Interval
        Case _
                DtInterval.dtDay, _
                DtInterval.dtDayOfYear, _
                DtInterval.dtHour, _
                DtInterval.dtMinute, _
                DtInterval.dtMonth, _
                DtInterval.dtQuarter, _
                DtInterval.dtSecond, _
                DtInterval.dtWeek, _
                DtInterval.dtWeekday, _
                DtInterval.dtYear
            Result = True
        Case _
                DtInterval.dtDimidiae, _
                DtInterval.dtMillisecond, _
                DtInterval.dtSemiyear, _
                DtInterval.dtSextayear, _
                DtInterval.dtTertiayear, _
                DtInterval.dtFortnight, _
                DtInterval.dtFortnightday, _
                DtInterval.dtSemimonth, _
                DtInterval.dtTertiamonth, _
                DtInterval.dtDecade, _
                DtInterval.dtCentury, _
                DtInterval.dtMillenium
            Result = Extended
    End Select

    IsInterval = Result
    
End Function

' Returns True if Interval is passed a valid value
' of DtInterval for intervals of one day or higher.
' Optionally, also returns True for an extended value.
'
' 2019-10-11. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsIntervalDate( _
    ByVal Interval As DtInterval, _
    Optional ByVal Extended As Boolean) _
    As Boolean
    
    Dim Result  As Boolean

    Select Case Interval
        Case _
                DtInterval.dtDay, _
                DtInterval.dtDayOfYear, _
                DtInterval.dtMonth, _
                DtInterval.dtQuarter, _
                DtInterval.dtWeek, _
                DtInterval.dtWeekday, _
                DtInterval.dtYear
            Result = True
        Case _
                DtInterval.dtDimidiae, _
                DtInterval.dtSemiyear, _
                DtInterval.dtSextayear, _
                DtInterval.dtTertiayear, _
                DtInterval.dtFortnight, _
                DtInterval.dtFortnightday, _
                DtInterval.dtSemimonth, _
                DtInterval.dtTertiamonth, _
                DtInterval.dtDecade, _
                DtInterval.dtCentury, _
                DtInterval.dtMillenium
            Result = Extended
    End Select

    IsIntervalDate = Result
    
End Function

' Returns True if the passed Value is a valid setting for
' parameter Interval in DateAdd, DateDiff, and DatePart.
' Optionally, validates custom (extended) values accepted
' by DateAddExt, DateDiffExt, and DatePartExt.
'
' The case of Value will be ignored.
'
' 2022-01-29. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsIntervalSetting( _
    ByVal Value As String, _
    Optional ByVal Extended As Boolean) _
    As Boolean
    
    Dim Interval        As DtInterval
    Dim LastInterval    As DtInterval
    Dim Symbol          As String
    Dim Result          As Boolean
    
    If Extended = False Then
        LastInterval = DtInterval.[_LastNative]
    Else
        LastInterval = DtInterval.[_Last]
    End If
    
    ' Exit with True if Value is a valid interval setting.
    If Value <> "" Then
        For Interval = DtInterval.[_First] To LastInterval
            Symbol = IntervalSetting(Interval, Extended)
            If LCase(Value) = Symbol Then
                Result = True
                Exit For
            End If
        Next
    End If
    
    IsIntervalSetting = Result

End Function

' Returns True if Interval is passed a valid value
' of DtInterval for intervals of less than one day.
' Optionally, also returns True for an extended value.
'
' 2015-11-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsIntervalTime( _
    ByVal Interval As DtInterval, _
    Optional ByVal Extended As Boolean) _
    As Boolean
    
    Dim Result  As Boolean

    Select Case Interval
        Case _
                DtInterval.dtHour, _
                DtInterval.dtMinute, _
                DtInterval.dtSecond
            Result = True
        Case _
                DtInterval.dtMillisecond, _
                DtInterval.dtDecimalSecond
            Result = Extended
    End Select

    IsIntervalTime = Result
    
End Function


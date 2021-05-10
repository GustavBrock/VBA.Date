Attribute VB_Name = "DateSpan"
Option Explicit
'
' DateSpan
' Version 1.6.2
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
'


' Date conversion.
'
' Name              Epoch                Calculation VBD                Calculation JD               Calculation RJD             Value at 2016-01-16 06:24:00
' ----------------  -------------------  ----------------------------  ----------------------------  --------------------------  ----------------------------
' Visual Basic VBD  AD  100-01-01 00:00  0                             JD - 2415018.5                RJD - 15018.5               42385.2666666667
' Julian Date JD    BC 4713-01-01 12:00  VBD + 2415018.5               0                             RJD + 2400000               2457403.76666667
' Reduced JD        AD 1858-11-16 12:00  VBD + 15018.5                 JD - 2400000                  0                           57403.7666666667 [7][8]
' Modified JD       AD 1858-11-17 00:00  VBD + 15018                   JD - 2400000.5                RJD - 0.5                   57403.2666666667 Introduced by SAO in 1957
' Truncated JD      AD 1968-05-24 00:00  Int(VBD - 24982)              Int(JD - 2440000.5)           Int(RJD - 40000.5)          17403            Introduced by NASA in 1979
' Dublin JD         AD 1899-12-31 12:00  VBD - 1.5                     JD - 2415020                  RJD - 15020                 42383.7666666667 Introduced by the IAU in 1955
' Lilian Date       AD 1582-10-15 00:00  Int(VBD + 115858)             Int(JD - 2299159.5)           Int(RJD + 100841.5)         158243           Count of days of the Gregorian calendar
' Rata Die          BC    1-12-31 00:00  Int(VBD + 693594)             Int(JD - 1721424.5)           Int(RJD + 678576.5)         735979           Count of days of the Common Era
' dotNet            AD    1-01-01 00:00  (VBD + 693593) × 86400000     (JD - 1721425.5) × 86400000   (RJD + 678575.5) × 86400000 63588522240000   Count of days from 1-1-1 in milliseconds
' Unix Time         AD 1970-01-01 00:00  (VBD - 25569) × 86400         (JD - 2440587.5) × 86400      (RJD - 40587.5) × 86400     1452925440       Count of seconds[10]
' Mars Sol Date MSD AD 1873-12-29 12:00  (VBD + 9496.5) / 1.0274912510 (JD - 2405522) / 1.0274912510 (RJD - 5522) / 1.0274912510 50493.63356      Count of Martian days
' Mars Sol Date MSD updated                                            (JD - 2405522.0025054) / 1.027491251                      50493.631438

' Day offset from Visual Basic numerical zero date value (1899-12-30 00:00:00.000).
    ' Julian Date.
    Private Const JdOffset          As Double = 2415018.5
    ' Reduced Julian Date.
    Private Const RjdOffset         As Double = 15018.5
    ' Modified Julian Date.
    Private Const MjdOffset         As Double = 15018
    ' Truncated Julian Date.
    Private Const TjdOffset         As Long = -24982
    ' Dublin Julian Date.
    Private Const DjdOffset         As Double = -1.5
    ' Lilian Date.
    Private Const LdOffset          As Long = 115858
    ' Rata Die.
    Private Const RdOffset          As Long = 693594
    ' dotNet.
    Private Const DnOffset          As Long = 693593
    ' Unix Time.
    Private Const UtOffset          As Long = -25569
    ' Mars Sol Date.
    Private Const MsdOffset         As Double = 9496.5
    
' Epochs. For information only.
    ' VBD Visual Basic Date Epoch.
    Private Const VbdEpoch          As Date = #1/1/100#
    ' RJD Reduced Julian Date Epoch.
    Private Const RjdEpoch          As Date = #11/16/1858 12:00:00 PM#
    ' MJD Modified Julian Date Epoch.
    Private Const MjdEpoch          As Date = #11/17/1858#
    ' TJD Truncated Julian Date Epoch.
    Private Const TjdEpoch          As Date = #5/24/1968#
    ' DJD Dublin Julian Date Epoch.
    Private Const DjdEpoch          As Date = #12/31/1899 12:00:00 PM#
    ' LJD Lilian Julian Date Epoch.
    Private Const LjdEpoch          As Date = #10/15/1582#
    ' Unix Time Epoch.
    Private Const UnixEpoch         As Date = #1/1/1970#
    ' MSD Mars Sol Date Epoch traditional.
    Private Const MsdEpoch          As Date = #12/29/1873 12:00:00 PM#
    ' MSD Mars Sol Date Epoch for zero value.
    ' 1873-12-29 12:03:04.283
    Private Const MsdEpochZero      As Date = #12/29/1873 12:03:04 PM#
    
' Martian time.
    ' Offset between JD and MSD.
    Private Const MsdJdOffset       As Double = 2405522.0025054
    ' Leap seconds as of February 2016.
    ' Difference between TAI, International Atomic Time, and UTC, Coordinated Universal Time.
    Private Const TaiLeapSeconds    As Double = 32.184
    ' Relation between Sol and Day.
    Private Const SolDayFactor      As Double = 1.027491251
    ' One Mars Sol is averagely 24:39:35.244 hours or 88775.2440864 seconds.
    Private Const MsdSecondsPerSol  As Double = 88775.2440864

' Calculates the .beats for the "Swatch Internet Time" from
' a date/time value.
' A such .beat is 1/1000 of a day or 1 minute 26.4 seconds,
' thus the count of .beats is between 0 and 999.
'
' The result is by default rounded by +/- half a .beat to
' the nearest integer .beat.
' Optionally, by passing parameter RoundSeconds as False,
' deciseconds will be respected: 4, 8, 2, 6, or 0.
'
' If .beats and times are converted back and forth using the
' functions Beat and DateBeat, parameter RoundSeconds must be
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
Public Function Beat( _
    ByVal Date1 As Date, _
    Optional ByVal RoundSeconds As Boolean = True) _
    As Integer
    
    Dim Slots   As Variant
    Dim Beats   As Integer
    
    ' Extract the time part using Fix to allow a date of 100-01-01,
    ' convert to Decimal to prevent floating point bit errors,
    ' turn numerial negative values (of dates earlier than 1899-12-30)
    ' positive with Abs, and round down to obtain a correct count.
    ' See examples above.
    Slots = Abs(CDec(Date1) - Fix(CDec(Date1))) * BeatsPerDay
    
    If RoundSeconds = True Then
        ' Round by 4/5 for split by times of +/- 0.5 beat.
        Beats = CInt(Slots) Mod BeatsPerDay
    Else
        ' Round down for exact split by the decisecond.
        Beats = Int(Slots)
    End If
    
    Beat = Beats
    
End Function

' Calculates the time from a count of .beats of the
' "Swatch Internet Time".
' A such .beat is 1/1000 of a day or 1 minute 26.4 seconds,
' thus the count of .beats is between 0 and 999.
'
' The result is by default rounded to the second.
' Optionally, by passing parameter RoundSeconds as False,
' deciseconds will be preserved: 4, 8, 2, 6, or 0.
'
' If .beats and times are converted back and forth using the
' functions Beat and DateBeat, parameter RoundSeconds must be
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
Public Function DateBeat( _
    ByVal Beats As Long, _
    Optional ByVal RoundSeconds As Boolean = True) _
    As Date
    
    Dim TimePart    As Double
    Dim TimeValue   As Date
    
    ' Limit the count of .beats to 0 to 999 and
    ' convert to a time value.
    TimePart = (Beats Mod BeatsPerDay) / BeatsPerDay
    If RoundSeconds = True Then
        ' Round deciseconds to nearest integer second.
        TimeValue = CDate(Int((TimePart * SecondsPerDay + 0.5)) / SecondsPerDay)
    Else
        ' Convert exactly respecting to deciseconds.
        TimeValue = CDate(TimePart)
    End If
    
    DateBeat = TimeValue

End Function

' Returns the date of a specified dotNet DateTime value with a resolution of 1 ms.
' DotNet can be any value that will return a valid VBA Date value.
'
' Minimum value:   31241376000000000
'   ->  100-01-01 00:00:00.000
' Maximum value: 3155378975999990000
'   -> 9999-12-31 23:59:59.999
'
' 2016-02-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateDotNet( _
    ByVal DotNet As Variant) _
    As Date

    Dim Timespan    As Variant
    Dim ResultDate  As Date
    
    Timespan = CDec(DotNet) / TicksPerMillisecond / MillisecondsPerDay - CDec(DnOffset)
    ResultDate = DateFromTimespan(Timespan)
    
    DateDotNet = ResultDate
  
End Function

' Returns the date of a specified Dublin Date with a resolution of 1 ms.
' DublinDate can be any value that will return a valid VBA Date value.
'
' Minimum value: -657435.5
'   ->  100-01-01 00:00:00.000
' Maximum value: 2958464.49999999
'   -> 9999-12-31 23:59:59.999
'
' 2016-02-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateDublin( _
    ByVal DublinDate As Variant) _
    As Date
    
    Dim Timespan    As Variant
    Dim ResultDate  As Date
    
    Timespan = CDec(DublinDate) - CDec(DjdOffset)
    ResultDate = DateFromTimespan(Timespan)
    
    DateDublin = ResultDate
    
End Function

' Returns the date of a specified Julian Date with a resolution of 1 ms.
' JulianDate can be any value that will return a valid VBA Date value.
'
' Minimum value: 1757584.5
'   ->  100-01-01 00:00:00.000
' Maximum value: 5373484.49999999
'   -> 9999-12-31 23:59:59.999
'
' 2016-02-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateJulian( _
    ByVal JulianDate As Variant) _
    As Date
    
    Dim Timespan    As Variant
    Dim ResultDate  As Date
    
    Timespan = CDec(JulianDate) - CDec(JdOffset)
    ResultDate = DateFromTimespan(Timespan)
    
    DateJulian = ResultDate
    
End Function

' Returns the date of a specified Lilian Date with a resolution of 1 ms.
' LilianDate can be any value that will return a valid VBA Date value.
'
' Minimum value: -541576
'   ->  100-01-01 00:00:00.000
' Maximum value: 3074323.99999999
'   -> 9999-12-31 23:59:59.999
'
' 2016-02-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateLilian( _
    ByVal LilianDate As Variant) _
    As Date
    
    Dim Timespan    As Variant
    Dim ResultDate  As Date
    
    Timespan = CDec(LilianDate) - CDec(LdOffset)
    ResultDate = DateFromTimespan(Timespan)
    
    DateLilian = ResultDate
    
End Function

' Returns the date of a specified Mars Sol with a resolution of 1 ms.
' MarsSol can be any value that will return a valid VBA Date value.
'
' Minimum value:  -630601.47860363630483117369142
'   ->  100-01-01 00:00:00.000
' Maximum value:  2888552.5740277942531365903824
'   -> 9999-12-31 23:59:59.999
'
' 2016-02-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateMarsSol( _
    ByVal MarsSol As Variant) _
    As Date
    
    Dim JulianDate  As Variant
    Dim Timespan    As Variant
    Dim ResultDate  As Date
    
    JulianDate = MarsSol * SolDayFactor - CDec(TaiLeapSeconds / SecondsPerDay) + MsdJdOffset
    Timespan = CDec(JulianDate) - CDec(JdOffset)
    ResultDate = DateFromTimespan(Timespan)
    
    DateMarsSol = ResultDate
    
End Function

' Returns the date of a specified Modified Julian Date with a resolution of 1 ms.
' ModifiedJulianDate can be any value that will return a valid VBA Date value.
'
' Minimum value: -642416
'   ->  100-01-01 00:00:00.000
' Maximum value: 2973483.99999999
'   -> 9999-12-31 23:59:59.999
'
' 2016-02-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateModifiedJulian( _
    ByVal ModifiedJulianDate As Variant) _
    As Date
    
    Dim Timespan    As Variant
    Dim ResultDate  As Date
    
    Timespan = CDec(ModifiedJulianDate) - CDec(MjdOffset)
    ResultDate = DateFromTimespan(Timespan)
    
    DateModifiedJulian = ResultDate

End Function

' Returns the date of a specified Rata Die with a resolution of 1 ms.
' RataDie can be any value that will return a valid VBA Date value.
'
' Minimum value:   36160
'   ->  100-01-01 00:00:00.000
' Maximum value: 3652059.99999999
'   -> 9999-12-31 23:59:59.999
'
' 2016-02-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateRataDie( _
    ByVal RataDie As Variant) _
    As Date

    Dim Timespan    As Variant
    Dim ResultDate  As Date
    
    Timespan = CDec(RataDie) - CDec(RdOffset)
    ResultDate = DateFromTimespan(Timespan)
    
    DateRataDie = ResultDate
  
End Function

' Returns the date of a specified Reduced Julian Date with a resolution of 1 ms.
' ReducedJulianDate can be any value that will return a valid VBA Date value.
'
' Minimum value: -642415.5
'   ->  100-01-01 00:00:00.000
' Maximum value: 2973484.49999999
'   -> 9999-12-31 23:59:59.999
'
' 2016-02-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateReducedJulian( _
    ByVal ReducedJulianDate As Variant) _
    As Date
    
    Dim Timespan    As Variant
    Dim ResultDate  As Date
    
    Timespan = CDec(ReducedJulianDate) - CDec(RjdOffset)
    ResultDate = DateFromTimespan(Timespan)
    
    DateReducedJulian = ResultDate
    
End Function

' Returns the date of a specified Truncated Julian Date with a resolution of 1 ms.
' TruncatedJulianDate can be any value that will return a valid VBA Date value.
'
' Minimum value: -682416
'   ->  100-01-01 00:00:00.000
' Maximum value: 2933483.99999999
'   -> 9999-12-31 23:59:59.999
'
' 2016-02-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateTruncatedJulian( _
    ByVal TruncatedJulianDate As Variant) _
    As Date
    
    Dim Timespan    As Variant
    Dim ResultDate  As Date
    
    Timespan = CDec(TruncatedJulianDate) - CDec(TjdOffset)
    ResultDate = DateFromTimespan(Timespan)
    
    DateTruncatedJulian = ResultDate

End Function

' Returns the date of a specified Unix Time with a resolution of 1 ms.
' UnixDate can be any value that will return a valid VBA Date value.
'
' Minimum value:  -59011459200
'   ->  100-01-01 00:00:00.000
' Maximum value:  253402300799.999
'   -> 9999-12-31 23:59:59.999
'
' 2016-02-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateUnix( _
    ByVal UnixDate As Variant) _
    As Date
    
    Dim Timespan    As Variant
    Dim ResultDate  As Date
    
    Timespan = (CDec(UnixDate) / SecondsPerDay) - CDec(UtOffset)
    ResultDate = DateFromTimespan(Timespan)
    
    DateUnix = ResultDate
    
End Function

' Returns the dotNet DateTime value in ticks for a specified date.
' UtcDate can be any Date value of VBA with a resolution of one millisecond.
'
' Examples:
'    100-01-01 00:00:00.000 ->   31241376000000000
'    100-01-01 00:00:00.001 ->   31241376000010000
'    100-01-01 00:00:00.002 ->   31241376000020000
'   1899-12-30 00:00:00.000 ->  599264352000000000
'   2018-08-18 03:24:47.000 ->  636701594870000000
'   2018-08-18 18:24:47.000 ->  636702134870000000
'   9999-12-31 23:59:59.000 -> 3155378975990000000
'   9999-12-31 23:59:59.998 -> 3155378975999980000
'   9999-12-31 23:59:59.999 -> 3155378975999990000
'
' 2016-02-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DotNet( _
    ByVal UtcDate As Date) _
    As Variant

    Dim Result  As Variant
    
    Result = Int((CDec(DateToTimespan(UtcDate) + CDec(DnOffset)) * MillisecondsPerDay + 0.5)) * TicksPerMillisecond
    
    DotNet = Result

End Function

' Returns the dotNet time in ticks rounded to 00:00:000 of the day for a specified date.
' UtcDate can be any Date value of VBA.
'
' Examples:
'    100-01-01 00:00:00.000 ->   31241376000000000
'    100-01-01 00:00:00.001 ->   31241376000000000
'    100-01-01 00:00:00.002 ->   31241376000000000
'   1899-12-30 00:00:00.000 ->  599264352000000000
'   2018-08-18 03:24:47.000 ->  636701472000000000
'   2018-08-18 18:24:47.000 ->  636701472000000000
'   9999-12-31 23:59:59.000 -> 3155378112000000000
'   9999-12-31 23:59:59.998 -> 3155378112000000000
'   9999-12-31 23:59:59.999 -> 3155378112000000000
'
' 2016-02-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DotNetDay( _
    ByVal UtcDate As Date) _
    As Variant
    
    Dim Result  As Variant
    
    Result = Int(DotNet(UtcDate) / SecondsPerDay / MillisecondsPerSecond / TicksPerMillisecond) _
        * SecondsPerDay * MillisecondsPerSecond * TicksPerMillisecond
    
    DotNetDay = Result
    
End Function

' Returns the Dublin Date for a specified date.
' UtcDate can be any Date value of VBA with a resolution of one millisecond.
'
' Examples:
'    100-01-01 00:00:00.000 -> -657435.5
'    100-01-01 00:00:00.001 -> -657435.49999998842592592592593
'    100-01-01 00:00:00.002 -> -657435.49999997685185185185185
'   1899-12-30 00:00:00.000 ->   15018.5
'   1858-11-16 12:00:00.000 ->       0
'   2018-08-18 03:24:47.000 ->   43328.642210648148148148148148
'   2018-08-18 18:24:47.000 ->   43329.267210648148148148148148
'   9999-12-31 23:59:59.000 -> 2958464.4999884259259259259259
'   9999-12-31 23:59:59.998 -> 2958464.4999999768518518518519
'   9999-12-31 23:59:59.999 -> 2958464.4999999884259259259259
'
' 2016-02-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DublinDate( _
    ByVal UtcDate As Date) _
    As Variant
    
    Dim Result  As Variant
    
    Result = Int(CDec(DateToTimespan(UtcDate) * MillisecondsPerDay) + 0.5) / MillisecondsPerDay + CDec(DjdOffset)
    
    DublinDate = Result

End Function

' Returns the Dublin Day for a specified date.
' UtcDate can be any Date value of VBA.
'
' Examples:
'    100-01-01 00:00:00.000 ->  -657436
'    100-01-01 00:00:00.001 ->  -657436
'    100-01-01 00:00:00.002 ->  -657436
'   1899-12-30 00:00:00.000 ->       -2
'   1899-12-31 12:00:00.000 ->        0
'   2018-08-18 03:24:47.000 ->    43328
'   2018-08-18 18:24:47.000 ->    43329
'   9999-12-31 23:59:59.000 ->  2958464
'   9999-12-31 23:59:59.998 ->  2958464
'   9999-12-31 23:59:59.999 ->  2958464
'
' 2016-02-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DublinDay( _
    ByVal UtcDate As Date) _
    As Long
    
    Dim Result  As Long
    
    Result = Int(DublinDate(UtcDate))
    
    DublinDay = Result
    
End Function

        
' Returns the Julian Date for a specified date.
' UtcDate can be any Date value of VBA with a resolution of one millisecond.
'
' Examples:
'    100-01-01 00:00:00.000 ->  1757584.5
'    100-01-01 00:00:00.001 ->  1757584.5000000115740740740741
'    100-01-01 00:00:00.002 ->  1757584.5000000231481481481482
'   1899-12-30 00:00:00.000 ->  2415018.5
'   2018-08-18 03:24:47.000 ->  2458348.6422106481481481481481
'   2018-08-18 18:24:47.000 ->  2458349.2672106481481481481481
'   9999-12-31 23:59:59.000 ->  5373484.4999884259259259259259
'   9999-12-31 23:59:59.998 ->  5373484.4999999768518518518519
'   9999-12-31 23:59:59.999 ->  5373484.4999999884259259259259
'
' 2016-02-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function JulianDate( _
    ByVal UtcDate As Date) _
    As Variant
    
    Dim Result  As Variant
    
    Result = Int(CDec(DateToTimespan(UtcDate) * MillisecondsPerDay) + 0.5) / MillisecondsPerDay + CDec(JdOffset)
    
    JulianDate = Result
    
End Function

' Returns the Julian Day for a specified date.
' UtcDate can be any Date value of VBA.
'
' Examples:
'    100-01-01 00:00:00.000 ->  1757584
'    100-01-01 00:00:00.001 ->  1757584
'    100-01-01 00:00:00.002 ->  1757584
'   1899-12-30 00:00:00.000 ->  2458349
'   2018-08-18 03:24:47.000 ->  2458348
'   2018-08-18 18:24:47.000 ->  2458349
'   9999-12-31 23:59:59.000 ->  5373484
'   9999-12-31 23:59:59.998 ->  5373484
'   9999-12-31 23:59:59.999 ->  5373484
'
' 2016-02-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function JulianDay( _
    ByVal UtcDate As Date) _
    As Long
    
    Dim Result  As Long
    
    Result = Int(JulianDate(UtcDate))
    
    JulianDay = Result
    
End Function

' Returns the Lilian Date for a specified date.
' UtcDate can be any Date value of VBA with a resolution of one millisecond.
'
' Examples:
'    100-01-01 00:00:00.000 -> -541576
'    100-01-01 00:00:00.001 -> -541575.99999998842592592592593
'    100-01-01 00:00:00.002 -> -541575.99999997685185185185185
'   1899-12-30 00:00:00.000 ->  115858
'   1582-10-15 00:00:00.000 ->       0
'   2018-08-18 03:24:47.000 ->  159188.14221064814814814814815
'   2018-08-18 18:24:47.000 ->  159188.76721064814814814814815
'   9999-12-31 23:59:59.000 -> 3074323.9999884259259259259259
'   9999-12-31 23:59:59.998 -> 3074323.9999999768518518518519
'   9999-12-31 23:59:59.999 -> 3074323.9999999884259259259259
'
' 2016-02-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function LilianDate( _
    ByVal UtcDate As Date) _
    As Variant
    
    Dim Result  As Variant
    
    Result = Int(CDec(DateToTimespan(UtcDate) * MillisecondsPerDay) + 0.5) / MillisecondsPerDay + CDec(LdOffset)
    
    LilianDate = Result

End Function

' Returns the Lilian Day for a specified date.
' UtcDate can be any Date value of VBA.
'
' Examples:
'    100-01-01 00:00:00.000 ->  -541576
'    100-01-01 00:00:00.001 ->  -541576
'    100-01-01 00:00:00.002 ->  -541576
'   1582-10-15 00:00:00.000 ->        0
'   1899-12-30 00:00:00.000 ->   115858
'   2018-08-18 03:24:47.000 ->   159188
'   2018-08-18 18:24:47.000 ->   159188
'   9999-12-31 23:59:59.000 ->  3074323
'   9999-12-31 23:59:59.998 ->  3074323
'   9999-12-31 23:59:59.999 ->  3074323
'
' 2016-02-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function LilianDay( _
    ByVal UtcDate As Date) _
    As Long
    
    Dim Result  As Long
    
    Result = Int(LilianDate(UtcDate))
    
    LilianDay = Result
    
End Function

' Returns the Mars Sol for a specified date.
' UtcDate can be any Date value of VBA.
'
' Examples:
'    100-01-01 00:00:00.000 ->  -630602
'    100-01-01 00:00:00.001 ->  -630602
'    100-01-01 00:00:00.002 ->  -630602
'   1873-12-29 12:03:04.283 ->        0
'   1899-12-30 00:00:00.000 ->     9242
'   2018-08-18 03:24:47.000 ->    51413
'   2018-08-18 18:24:47.000 ->    51413
'   9999-12-31 23:59:59.000 ->  2888552
'   9999-12-31 23:59:59.998 ->  2888552
'   9999-12-31 23:59:59.999 ->  2888552
'
' 2016-02-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function MarsSol( _
    ByVal UtcDate As Date) _
    As Long
    
    Dim Result  As Long
    
    Result = Int(MarsSolDate(UtcDate))
    
    MarsSol = Result
    
End Function

' Returns the Mars Sol for a specified date.
' UtcDate can be any Date value of VBA with a resolution of one millisecond.
'
' Examples:
'    100-01-01 00:00:00.000 ->  -630601.47860363630483117369142
'    100-01-01 00:00:00.001 ->  -630601.47860362504042959089478
'    100-01-01 00:00:00.002 ->  -630601.47860361377602800809814
'   1873-12-29 12:03:04.283 ->        0.0000000049563366964305178303
'   1899-12-30 00:00:00.000 ->     9242.4123882880633890672418
'   2018-08-18 03:24:47.000 ->    51413.226172324992525068369755
'   2018-08-18 18:24:47.000 ->    51413.8344500104635422809533
'   9999-12-31 23:59:59.000 ->  2888552.574016541115955376564
'   9999-12-31 23:59:59.998 ->  2888552.5740277829887350075859
'   9999-12-31 23:59:59.999 ->  2888552.5740277942531365903824
'
' 2016-02-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function MarsSolDate( _
    ByVal UtcDate As Date) _
    As Variant
    
    Dim Result  As Variant
    
    Result = (JulianDate(UtcDate) + CDec(TaiLeapSeconds / SecondsPerDay) - MsdJdOffset) / SolDayFactor
    
    MarsSolDate = Result
        
End Function

' Returns the time part of a specified Mars Sol.
'
' 2016-02-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function MarsTime( _
    ByVal MarsSol As Variant) _
    As Date
    
    Dim ResultTime  As Date
    
    ResultTime = CDate(MarsSol - Fix(MarsSol))
    
    MarsTime = ResultTime
    
End Function

' Returns the Modified Julian Date for a specified date.
' UtcDate can be any Date value of VBA with a resolution of one millisecond.
'
' Examples:
'    100-01-01 00:00:00.000 -> -642416
'    100-01-01 00:00:00.001 -> -642415.99999998842592592592593
'    100-01-01 00:00:00.002 -> -642415.99999997685185185185185
'   1899-12-30 00:00:00.000 ->   15018
'   1858-11-17 00:00:00.000 ->       0
'   2018-08-18 03:24:47.000 ->   58348.142210648148148148148148
'   2018-08-18 18:24:47.000 ->   58348.767210648148148148148148
'   9999-12-31 23:59:59.000 -> 2973483.9999884259259259259259
'   9999-12-31 23:59:59.998 -> 2973483.9999999768518518518519
'   9999-12-31 23:59:59.999 -> 2973483.9999999884259259259259
'
' 2016-02-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function ModifiedJulianDate( _
    ByVal UtcDate As Date) _
    As Variant
    
    Dim Result  As Variant
    
    Result = Int(CDec(DateToTimespan(UtcDate) * MillisecondsPerDay) + 0.5) / MillisecondsPerDay + CDec(MjdOffset)
    
    ModifiedJulianDate = Result

End Function

' Returns the Modified Julian Day for a specified date.
' UtcDate can be any Date value of VBA.
'
' Examples:
'    100-01-01 00:00:00.000 ->  -642416
'    100-01-01 00:00:00.001 ->  -642416
'    100-01-01 00:00:00.002 ->  -642416
'   1858-11-17 00:00:00.000 ->        0
'   1899-12-30 00:00:00.000 ->    15018
'   2018-08-18 03:24:47.000 ->    58348
'   2018-08-18 18:24:47.000 ->    58348
'   9999-12-31 23:59:59.000 ->  2973483
'   9999-12-31 23:59:59.998 ->  2973483
'   9999-12-31 23:59:59.999 ->  2973483
'
' 2016-02-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function ModifiedJulianDay( _
    ByVal UtcDate As Date) _
    As Long
    
    Dim Result  As Long
    
    Result = Int(ModifiedJulianDate(UtcDate))
    
    ModifiedJulianDay = Result
    
End Function

' Returns the ordinal day for a specified date.
' This is sometimes (wrongly) named the Julian day.
' Date1 can be any Date value of VBA.
'
' Examples:
'    100-01-01 ->     1
'   1899-12-30 ->   364
'   1980-03-01 ->    61
'   1981-03-01 ->    60
'   2000-12-31 ->   366
'   9999-12-31 ->   365
'
' 2021-04-02. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function OrdinalDay( _
    ByVal Date1 As Date) _
    As Integer
    
    Dim Result  As Integer
    
    Result = DatePart("y", Date1)
    
    OrdinalDay = Result
    
End Function

' Returns the Rata Die for a specified date.
' UtcDate can be any Date value of VBA with a resolution of one millisecond.
'
' Examples:
'    100-01-01 00:00:00.000 ->   36160
'    100-01-01 00:00:00.001 ->   36160.00000001157407407407407
'    100-01-01 00:00:00.002 ->   36160.00000002314814814814815
'   1899-12-30 00:00:00.000 ->  693594
'   2018-08-18 03:24:47.000 ->  736924.14221064814814814814815
'   2018-08-18 18:24:47.000 ->  736924.76721064814814814814815
'   9999-12-31 23:59:59.000 -> 3652059.9999884259259259259259
'   9999-12-31 23:59:59.998 -> 3652059.9999999768518518518519
'   9999-12-31 23:59:59.999 -> 3652059.9999999884259259259259
'
' 2016-02-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function RataDie( _
    ByVal UtcDate As Date) _
    As Variant

    Dim Result  As Variant
    
    Result = Int(CDec(DateToTimespan(UtcDate) * MillisecondsPerDay) + 0.5) / MillisecondsPerDay + CDec(RdOffset)
    
    RataDie = Result

End Function

' Returns the Rata Die Day for a specified date.
' UtcDate can be any Date value of VBA.
'
' Examples:
'    100-01-01 00:00:00.000 ->    36160
'    100-01-01 00:00:00.001 ->    36160
'    100-01-01 00:00:00.002 ->    36160
'   1899-12-30 00:00:00.000 ->   693594
'   2018-08-18 03:24:47.000 ->   736924
'   2018-08-18 18:24:47.000 ->   736924
'   9999-12-31 23:59:59.000 ->  3652059
'   9999-12-31 23:59:59.998 ->  3652059
'   9999-12-31 23:59:59.999 ->  3652059
'
' 2016-02-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function RataDieDay( _
    ByVal UtcDate As Date) _
    As Long
    
    Dim Result  As Long
    
    Result = Int(RataDie(UtcDate))
    
    RataDieDay = Result
    
End Function

' Returns the Reduced Julian Date for a specified date.
' UtcDate can be any Date value of VBA with a resolution of one millisecond.
'
' Examples:
'    100-01-01 00:00:00.000 -> -642415.5
'    100-01-01 00:00:00.001 -> -642415.49999998842592592592593
'    100-01-01 00:00:00.002 -> -642415.49999997685185185185185
'   1899-12-30 00:00:00.000 ->   15018.5
'   1858-11-16 12:00:00.000 ->       0
'   2018-08-18 03:24:47.000 ->   58348.642210648148148148148148
'   2018-08-18 18:24:47.000 ->   58349.267210648148148148148148
'   9999-12-31 23:59:59.000 -> 2973484.4999884259259259259259
'   9999-12-31 23:59:59.998 -> 2973484.4999999768518518518519
'   9999-12-31 23:59:59.999 -> 2973484.4999999884259259259259
'
' 2016-02-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function ReducedJulianDate( _
    ByVal UtcDate As Date) _
    As Variant
    
    Dim Result  As Variant
    
    Result = Int(CDec(DateToTimespan(UtcDate) * MillisecondsPerDay) + 0.5) / MillisecondsPerDay + CDec(RjdOffset)
    
    ReducedJulianDate = Result

End Function

' Returns the Reduced Julian Day for a specified date.
' UtcDate can be any Date value of VBA.
'
' Examples:
'    100-01-01 00:00:00.000 ->  -642416
'    100-01-01 00:00:00.001 ->  -642416
'    100-01-01 00:00:00.002 ->  -642416
'   1858-11-16 12:00:00.000 ->        0
'   1899-12-30 00:00:00.000 ->    15018
'   2018-08-18 03:24:47.000 ->    58348
'   2018-08-18 18:24:47.000 ->    58349
'   9999-12-31 23:59:59.000 ->  2973484
'   9999-12-31 23:59:59.998 ->  2973484
'   9999-12-31 23:59:59.999 ->  2973484
'
' 2016-02-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function ReducedJulianDay( _
    ByVal UtcDate As Date) _
    As Long
    
    Dim Result  As Long
    
    Result = Int(ReducedJulianDate(UtcDate))
    
    ReducedJulianDay = Result
    
End Function

' Returns the Truncated Julian Date for a specified date.
' UtcDate can be any Date value of VBA with a resolution of one millisecond.
'
' Examples:
'    100-01-01 00:00:00.000 -> -682416
'    100-01-01 00:00:00.001 -> -682415.99999998842592592592593
'    100-01-01 00:00:00.002 -> -682415.99999997685185185185185
'   1899-12-30 00:00:00.000 ->  -24982
'   1968-05-24 00:00:00.000 ->       0
'   2018-08-18 03:24:47.000 ->   18348.142210648148148148148148
'   2018-08-18 18:24:47.000 ->   18348.767210648148148148148148
'   9999-12-31 23:59:59.000 -> 2933483.9999884259259259259259
'   9999-12-31 23:59:59.998 -> 2933483.9999999768518518518519
'   9999-12-31 23:59:59.999 -> 2933483.9999999884259259259259
'
' 2016-02-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function TruncatedJulianDate( _
    ByVal UtcDate As Date) _
    As Variant
    
    Dim Result  As Variant
    
    Result = Int(CDec(DateToTimespan(UtcDate) * MillisecondsPerDay) + 0.5) / MillisecondsPerDay + CDec(TjdOffset)
    
    TruncatedJulianDate = Result

End Function

' Returns the Truncated Julian Day for a specified date.
' UtcDate can be any Date value of VBA.
'
' Examples:
'    100-01-01 00:00:00.000 ->  -682416
'    100-01-01 00:00:00.001 ->  -682416
'    100-01-01 00:00:00.002 ->  -682416
'   1899-12-30 00:00:00.000 ->   -24982
'   1968-05-24 00:00:00.000 ->        0
'   2018-08-18 03:24:47.000 ->    18348
'   2018-08-18 18:24:47.000 ->    18348
'   9999-12-31 23:59:59.000 ->  2933483
'   9999-12-31 23:59:59.998 ->  2933483
'   9999-12-31 23:59:59.999 ->  2933483
'
' 2016-02-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function TruncatedJulianDay( _
    ByVal UtcDate As Date) _
    As Long
    
    Dim Result  As Long
    
    Result = Int(TruncatedJulianDate(UtcDate))
    
    TruncatedJulianDay = Result
    
End Function

' Returns the Unix Time in seconds for a specified date.
' UtcDate can be any Date value of VBA with a resolution of one millisecond.
'
' Examples:
'    100-01-01 00:00:00.000 ->  -59011459200
'    100-01-01 00:00:00.001 ->  -59011459199.999
'    100-01-01 00:00:00.002 ->  -59011459199.998
'   1899-12-30 00:00:00.000 ->   -2209161600
'   1970-01-01 00:00:00.000 ->             0
'   2018-08-18 03:24:47.000 ->    1534562687
'   2018-08-18 18:24:47.000 ->    1534616687
'   9999-12-31 23:59:59.000 ->  253402300799
'   9999-12-31 23:59:59.998 ->  253402300799.998
'   9999-12-31 23:59:59.999 ->  253402300799.999
'
' 2016-02-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function UnixDate( _
    ByVal UtcDate As Date) _
    As Variant
    
    Dim Result  As Variant
    
    Result = Int((CDec(DateToTimespan(UtcDate) + CDec(UtOffset)) * MillisecondsPerDay + 0.5)) / MillisecondsPerSecond
    
    UnixDate = Result
    
End Function

' Returns the Unix Time rounded to 00:00:000 of the day for a specified date.
' UtcDate can be any Date value of VBA with a resolution of one millisecond.
'
' Examples:
'    100-01-01 00:00:00.000 ->  -59011459200
'    100-01-01 00:00:00.001 ->  -59011459200
'    100-01-01 00:00:00.002 ->  -59011459200
'   1899-12-30 00:00:00.000 ->   -2209161600
'   1970-01-01 00:00:00.000 ->             0
'   2018-08-18 03:24:47.000 ->    1534550400
'   2018-08-18 18:24:47.000 ->    1534550400
'   9999-12-31 23:59:59.000 ->  253402214400
'   9999-12-31 23:59:59.998 ->  253402214400
'   9999-12-31 23:59:59.999 ->  253402214400
'
' 2016-02-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function UnixDay( _
    ByVal UtcDate As Date) _
    As Variant
    
    Dim Result  As Variant
    
    Result = Int(UnixDate(UtcDate) / SecondsPerDay) * SecondsPerDay
    
    UnixDay = Result
    
End Function


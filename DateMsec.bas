Attribute VB_Name = "DateMsec"
Option Explicit
'
' DateMsec
' Version 1.3.1
'
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.Date
'
' Functions to generate and handle date/time with millisecond accuracy in VBA.
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
'


' Note:
' Dates in VBA follow a pseudo Gregorian calender from year 100 to
' the introduction of the Gregorian calendar in October 1582:
'
' http://www.ghgrb.ch/index.php/genealogie/erklaerung-der-kalender
'
'
' General note:
'   The numeric value of date
'     100-1-1 23:59:59.999
'   is lower than that of date
'     100-1-1 00:00:00.000
'
'
' SQL methods.
' To extract the millisecond of a date value and have this rounded down to integer second
' using SQL only while sorting date values prior to 1899-12-30 correctly:
'
'   SELECT
'     [DateTimeMs],
'     Fix([DateTimeMs]*24*60*60)/(24*60*60) AS RoundSecSQL,
'     (([DateTimeMs]-Fix([DateTimeMs]))*24*60*60*1000)*Sgn([DateTimeMs]) Mod 1000 AS MsecSQL
'   FROM
'     TestTimeMsec
'   ORDER BY
'     Fix([DateTimeMs]),
'     Abs([DateTimeMs]);


' Enums.
'
    Public Enum DtMsecResult
        dtMillisecondOnly = 0
        dtTimeMillisecond = 1
        dtDateTimeMillisecond = 2
    End Enum
    

' Declarations.
'
Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" ( _
    ByRef lpSystemTime As SystemTime)
    
' Retrieves the current time zone settings from Windows.
Private Declare PtrSafe Function GetTimeZoneInformation Lib "Kernel32.dll" ( _
    ByRef lpTimeZoneInformation As TimeZoneInformation) _
    As Long
    
Private Declare PtrSafe Function TimeBeginPeriod Lib "winmm.dll" Alias "timeBeginPeriod" ( _
    ByVal uPeriod As Long) _
    As Long

Private Declare PtrSafe Function TimeGetTime Lib "winmm.dll" Alias "timeGetTime" () _
    As Long

' Converts a date/time expression including milliseconds
' to a date/time value.
' Note: Will raise error 13 if the cleaned Expression
' does not represent a valid date expression.
'
' 2016-09-14. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function CDateMsec( _
    ByVal Expression As Variant) _
    As Date

    Dim DateText    As String
    Dim DatePart    As Date
    Dim MsecPart    As Date
    Dim Result      As Date
  
    ' First try IsDate(Expression) or IsNumeric(Expression).
    ' If success, simply use CDate.
    If IsDate(Expression) Then
        ' Convert a date expression to a date value.
        Result = CDate(Expression)
    ElseIf IsNumeric(Expression) Then
        ' Convert a numeric value to a date value.
        Result = CDate(CDbl(Expression))
    Else
        ' Convert Expression to a string value.
        DateText = CStr(Expression)
        ' Try to convert DateText to a normal date string expression only
        ' by stripping a millisecond part.
        MsecPart = ExtractMsec(DateText)
        ' ExtractMsec returned a cleaned DateText.
        
        ' If DateText represents a date value, convert DateText to
        ' a date value and use this as base date for MsecSerial.
        ' Using CDate will, as usual, raise an error if DateText
        ' is not a valid date expression.
        DatePart = CDate(DateText)
        
        ' Return the combined date part and millisecond part.
        Result = MsecSerial(Millisecond(MsecPart), DatePart)
    End If
  
    CDateMsec = Result

End Function

' Converts an integer count of milliseconds to a three-character string.
' Maximum value is 999.
'
' Examples:
'   MsText = CStrMillisecond(87)
'   ' MsText -> 087
'   MsText = CStrMillisecond(2056)
'   ' MsText -> 056
'
' 2016-12-10. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function CStrMillisecond( _
    ByVal Expression As Integer) _
    As String
    
    Dim Result As String
    
    Result = Right("00" & LTrim(Str(Expression)), 3)
    
    CStrMillisecond = Result

End Function

' Converts a date/time expression including milliseconds
' to a date/time value.
' As for CVDate, also Null is accepted and is returned as Null.
' Note: Will raise error 13 if the cleaned Expression is not Null
' and does not represent a valid date expression.
'
' 2016-09-14. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function CVDateMsec( _
    ByVal Expression As Variant) _
    As Variant

    Dim Result  As Variant
    
    On Error GoTo Err_CVDateMsec
    
    If IsNull(Expression) Then
        Result = Null
    Else
        Result = CDateMsec(Expression)
    End If
  
    CVDateMsec = Result

Exit_CVDateMsec:
    Exit Function
  
Err_CVDateMsec:
    Err.Raise Err.Number
    Resume Exit_CVDateMsec

End Function

' Adds milliseconds or decimal seconds - as well as
' all the standard date/time intervals - to Date1.
'
' Interval "f" will add milliseconds to Date1.
' Input range is any value between
'   -312,413,759,999,999 and 312,413,759,999,999
' that will result in a valid date value between
'   100-1-1 00:00:00.000 and 9999-12-31 23:59:59.999
'
' Interval "l" will add decimal seconds to Date1.
' Input range is any value between
'   -312,413,759,999.999 and 312,413,759,999.999
' that will result in a valid date value between
'   100-1-1 00:00:00.000 and 9999-12-31 23:59:59.999
'
' Any other Interval is handled by DateAdd as usual.
'
' 2016-09-14. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateAddMsec( _
    ByVal Interval As String, _
    ByVal Number As Double, _
    ByVal Date1 As Date) _
    As Date
    
    Dim Result  As Date
  
    Select Case IntervalValue(Interval, True)
        Case DtInterval.dtMillisecond
            Result = MsecSerial(Number, Date1)
        Case DtInterval.dtSecond, DtInterval.dtDecimalSecond
            Result = MsecSerial(Number * MillisecondsPerSecond, Date1)
        Case DtInterval.dtMinute
            Result = MsecSerial(Number * MillisecondsPerMinute, Date1)
        Case Else
            Result = DateAdd(Interval, Number, Date1)
    End Select
  
    DateAddMsec = Result

End Function

' Will calculate the difference in milliseconds or decimal seconds
' from Date1 to Date2 as well as all the standard date/time diffs.
'
' Date1 and Date2 can be any valid date value between
'   100-1-1 00:00:00.000 and 9999-12-31 23:59:59.999
'
' Interval "f" will return the difference in milliseconds.
' Interval "l" will return the difference in decimal seconds.
' Any other Interval is handled by DateDiff as usual.
'
' 2016-09-14. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateDiffMsec( _
    ByVal Interval As String, _
    ByVal Date1 As Date, _
    ByVal Date2 As Date, _
    Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbSunday, _
    Optional ByVal FirstWeekOfYear As VbFirstWeekOfYear = vbFirstJan1) _
    As Double
  
    Dim Milliseconds    As Double
  
    Select Case IntervalValue(Interval, True)
        Case DtInterval.dtMillisecond
            Milliseconds = MsecDiff(Date1, Date2)
        Case DtInterval.dtDecimalSecond
            Milliseconds = MsecDiff(Date1, Date2) / MillisecondsPerSecond
        Case Else
            Milliseconds = DateDiff(Interval, Date1, Date2, FirstDayOfWeek, FirstWeekOfYear)
    End Select
  
    DateDiffMsec = Milliseconds
  
End Function

' Rounds off Date1 to the second and optionally adds the specified
' millisecond part up to and including 999 milliseconds.
'
' Typical usage:
'   Sequentialize a series of date values identical by the second.
'
'   For each element in <collection of identical date values>
'       Date1 = <read date value>
'       Date1 = DateMsecSet(i, Date1)
'       <write date value> = Date1
'       i = i + 1
'   Next
'
' 2019-11-04. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateMsecSet( _
    ByVal Number As Integer, _
    ByVal Date1 As Date) _
    As Date
  
    Dim Result As Date
    
    ' Round off a millisecond part.
    RoundOffMilliseconds Date1
    If Number <= 0 Then
        ' No milliseconds.
        Result = Date1
    Else
        ' Add the count of milliseconds from 0 up to 999.
        If Number > MaxMillisecondCount Then
            Number = MaxMillisecondCount
        End If
        Result = DateAddMsec(IntervalSetting(DtInterval.dtMillisecond, True), Number, Date1)
    End If
    
    DateMsecSet = Result
  
End Function

' Extracts milliseconds or decimal seconds - as well as all
' the standard date/time parts - from Date1.
'
' Interval "f" will return millisecond of Date1.
' Interval "l" will return decimal seconds of Date1.
' Any other Interval is handled by DatePart as usual.
'
' Note that DatePartExt should be used for finding seconds
' if Date1 is the last date (9999-12-31).
'
' 2016-09-14. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DatePartMsec( _
    ByVal Interval As String, _
    ByVal Date1 As Date, _
    Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbSunday, _
    Optional ByVal FirstWeekOfYear As VbFirstWeekOfYear = vbFirstJan1) _
    As Double

    Dim Result  As Double
  
    Select Case IntervalValue(Interval, True)
        Case DtInterval.dtMillisecond
            Result = Millisecond(Date1)
        Case DtInterval.dtDecimalSecond
            Result = DecimalSecond(Date1)
        Case Else
            Result = DatePart(Interval, Date1, FirstDayOfWeek, FirstWeekOfYear)
    End Select
  
    DatePartMsec = Result

End Function

' Returns Date1 rounded to the nearest millisecond approximately by 4/5.
' The dividing point for up/down rounding may vary between 0.3 and 0.7ms
' due to the limited resolution of data type Double.
'
' If RoundSqlServer is True, milliseconds are rounded by 3.333ms to match
' the rounding of the Datetime data type of SQL Server - to 0, 3 or 7 as the
' least significant digit:
'
' Msec SqlServer
'   0    0
'   1    0
'   2    3
'   3    3
'   4    3
'   5    7
'   6    7
'   7    7
'   8    7
'   9   10
'  10   10
'  11   10
'  12   13
'  13   13
'  14   13
'  15   17
'  16   17
'  17   17
'  18   17
'  19   20
' ...
' 990  990
' 991  990
' 992  993
' 993  993
' 994  993
' 995  997
' 996  997
' 997  997
' 998  997
' 999 1000
'
' If RoundSqlServer is True and if RoundSecondUp is True, 999ms will be
' rounded up to 1000ms - the next second - which may not be what you wish.
' If RoundSecondUp is False, 999ms will be rounded down to 997ms:
'
' 994  993
' 995  997
' 996  997
' 997  997
' 998  997
' 999  997
'
' If RoundSqlServer is False, RoundSecondUp is ignored.
'
' 2016-09-14. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateRoundMillisecond( _
    ByVal Date1 As Date, _
    Optional RoundSqlServer As Boolean, _
    Optional RoundSecondUp As Boolean) _
    As Date
  
    Dim Milliseconds    As Integer
    Dim MsecValue       As Date
    Dim Result          As Date
  
    ' Retrieve the millisecond part of Date1.
    Milliseconds = Millisecond(Date1)
    If RoundSqlServer = True Then
        ' Perform special rounding to match data type datetime of SQL Server.
        Milliseconds = (Milliseconds \ 10) * 10 + Choose(Milliseconds Mod 10 + 1, 0, 0, 3, 3, 3, 7, 7, 7, 7, 10)
        If RoundSecondUp = False Then
            If Milliseconds = 1000 Then
                Milliseconds = 997
            End If
        End If
    End If
    
    ' Round Date1 down to the second.
    Call RoundOffMilliseconds(Date1)
    ' Get milliseconds as date value.
    MsecValue = MsecSerial(Milliseconds)
    ' Add milliseconds to rounded date.
    Result = DateFromTimespan(DateToTimespan(Date1) + DateToTimespan(MsecValue))
  
    DateRoundMillisecond = Result
  
End Function

' Returns Date1 rounded off to the second by
' removing a millisecond portion.
'
' 2016-09-14. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateRoundOffMilliseconds( _
    ByVal Date1 As Date) _
    As Date
    
    Call RoundOffMilliseconds(Date1)
  
    DateRoundOffMilliseconds = Date1
  
End Function

' Cleans Expression for milliseconds and a time part
' and, if possible, returns the date value of Expression.
'
' Note:
'   Will raise error 13 if the cleaned Expression
'   does not represent a valid date value.
'
' 2016-09-17. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateValueMsec( _
    ByVal Expression As String) _
    As Date
  
    Dim Result As Date
  
    On Error GoTo Err_DateValueMsec
  
    ' Strip a millisecond part from Expression.
    ExtractMsec Expression
    ' Convert Expression to a date value with no time part.
    Result = DateValue(Expression)
    
    DateValueMsec = Result

Exit_DateValueMsec:
    Exit Function
    
Err_DateValueMsec:
    Err.Raise Err.Number
    Resume Exit_DateValueMsec
  
End Function

' Returns the second and the millisecond from Date1 as a decimal value.
'
' 2017-11-27. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DecimalSecond( _
    ByVal Date1 As Date) _
    As Double
    
    Dim Seconds         As Integer
    Dim Milliseconds    As Integer
    Dim TotalSeconds    As Double
    
    ' Get milliseconds of Date1.
    Milliseconds = Millisecond(Date1)
    ' Round off Date1 to the second.
    Call RoundOffMilliseconds(Date1)
    ' Get the rounded count of seconds.
    Seconds = Second(Date1 - Fix(Date1))
    ' Calculate seconds and milliseconds as decimal seconds.
    TotalSeconds = Seconds + Milliseconds / MillisecondsPerSecond
    
    DecimalSecond = TotalSeconds
  
End Function

' Returns millisecond date/time value from the last digits of a Expression.
'
' Note:
'   Returns ByRef Expression without millisecond part.
'   To pass Expression ByVal, call the function like this:
'       Result = ExtractMsec((Expression))
'   or use MsecValueMsec(Expression).
'
' Examples:
'   "01:13"             ->   0 milliseconds
'   "09:25.17"          ->   0 milliseconds
'   "11:45:27"          ->   0 milliseconds
'   "08:33:12 AM 60"    ->   0 milliseconds
'   "18:23:22.322"      -> 322 milliseconds
'   "18:23:22.322 ms"   -> 322 milliseconds
'   "18:23:22-322"      ->   0 milliseconds
'   "08:33:42.87391"    -> 873 milliseconds
'   "08:33:42.87.391"   -> 391 milliseconds
'   "45.078"            ->  78 milliseconds
'   "55.6"              -> 600 milliseconds
'   ".04758"            ->  47 milliseconds
'   "822327"            ->   0 milliseconds
'   "73.Mil"            ->   0 milliseconds
'
' 2016-09-16. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function ExtractMsec( _
    ByRef Expression As String) _
    As Date

    ' Only two parts: date/time and milliseconds.
    Const PartCount As Long = 2
    
    Dim Parts       As Variant
    Dim Result      As Date
    
    If IsDate(Expression) Then
        ' Expression represents a valid time expression, thus it contains no milliseconds.
        ' Nothing to do.
    Else
        Parts = Split(StrReverse(RTrim(Expression)), MillisecondSeparator, PartCount)
        If UBound(Parts) = 0 Then
            ' No millisecond part.
        Else
            ' Return the millisecond part as a date value.
            Result = MsecSerial(Val(Right(Left(StrReverse(Parts(LBound(Parts))) & "00", 3), 3)))
            ' Return Expression with stripped millisecond part.
            Expression = StrReverse(Parts(UBound(Parts)))
        End If
    End If
    
    ExtractMsec = Result

End Function

' Formats a millisecond part of a date expression.
'
' Of the Format expression, only the first millisecond format symbol
' is replaced. Thus, the returned format expression can be processed
' by a following process.
' Use FormatExt if several symbols must be processed in one go.
'
' Parameter SecondRoundUp is returned as True if the millisecond value
' is rounded up to 1000, thus being displayed as ".0", or ".00".
'
' Examples from FormatMillisecond(x, y, SecondRoundUp):
'   x = "2016-09-22 16:20:18.007"
'       y = "hh:ss.fff"     -> "hh:ss.007"  SecondRoundUp = False
'       y = "hh:ss.ff#"     -> "hh:ss.007"  SecondRoundUp = False
'       y = "hh:ss.f##"     -> "hh:ss.007"  SecondRoundUp = False
'       y = "hh:ss.ff"      -> "hh:ss.01"   SecondRoundUp = False
'       y = "hh:ss.f#"      -> "hh:ss.01"   SecondRoundUp = False
'       y = "hh:ss.f"       -> "hh:ss.0"    SecondRoundUp = False
'       y = "Now f milliseconds"        ->  "Now 7 milliseconds"
'       y = "Now ff milliseconds"       ->  "Now 07 milliseconds"
'       y = "Now fff milliseconds"      ->  "Now 007 milliseconds"
'   x = "2016-09-22 16:20:18.045"
'       y = "hh:ss.fff"     -> "hh:ss.045"  SecondRoundUp = False
'       y = "hh:ss.ff#"     -> "hh:ss.045"  SecondRoundUp = False
'       y = "hh:ss.f##"     -> "hh:ss.045"  SecondRoundUp = False
'       y = "hh:ss.ff"      -> "hh:ss.05"   SecondRoundUp = False
'       y = "hh:ss.f#"      -> "hh:ss.05"   SecondRoundUp = False
'       y = "hh:ss.f"       -> "hh:ss.0"    SecondRoundUp = False
'       y = "Now f milliseconds"        ->  "Now 45 milliseconds"
'       y = "Now ff milliseconds"       ->  "Now 45 milliseconds"
'       y = "Now fff milliseconds"      ->  "Now 045 milliseconds"
'   x = "2016-09-22 16:20:18.400"
'       y = "hh:ss.fff"     -> "hh:ss.400"  SecondRoundUp = False
'       y = "hh:ss.ff#"     -> "hh:ss.40"   SecondRoundUp = False
'       y = "hh:ss.f##"     -> "hh:ss.4"    SecondRoundUp = False
'       y = "hh:ss.ff"      -> "hh:ss.40"   SecondRoundUp = False
'       y = "hh:ss.f#"      -> "hh:ss.4"    SecondRoundUp = False
'       y = "hh:ss.f"       -> "hh:ss.4"    SecondRoundUp = False
'       y = "Now f milliseconds"        ->  "Now 400 milliseconds"
'       y = "Now ff milliseconds"       ->  "Now 400 milliseconds"
'       y = "Now fff milliseconds"      ->  "Now 400 milliseconds"
'   x = "2016-09-22 16:59:37.985"
'       y = "hh:ss.fff"     -> "hh:ss.985"  SecondRoundUp = False
'       y = "hh:ss.ff#"     -> "hh:ss.985"  SecondRoundUp = False
'       y = "hh:ss.f##"     -> "hh:ss.985"  SecondRoundUp = False
'       y = "hh:ss.ff"      -> "hh:ss.99"   SecondRoundUp = False
'       y = "hh:ss.f#"      -> "hh:ss.99"   SecondRoundUp = False
'       y = "hh:ss.f"       -> "hh:ss.0"    SecondRoundUp = True
'       y = "Now f milliseconds"        ->  "Now 985 milliseconds"
'       y = "Now ff milliseconds"       ->  "Now 985 milliseconds"
'       y = "Now fff milliseconds"      ->  "Now 985 milliseconds"
'
' 2016-12-11. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function FormatMillisecond( _
    ByVal Expression As Variant, _
    ByVal Format As String, _
    Optional ByRef SecondRoundUp As Boolean) _
    As String
    
    ' Symbol for optional trailing milliseconds: yyyy-mm-dd hh:nn:ss.f#[#]
    Const MsecOptional  As String = "#"
    Const DummySymbol   As String = "" ' Chr(15)
    
    Dim LocalFormat     As String
    Dim MsecSymbol      As String
    Dim DateValue       As Date
    Dim Milliseconds    As Integer
    Dim Length          As Long
    Dim Symbol          As String
    Dim SymbolLength    As Long
    Dim Digits          As Integer
    Dim MsecText        As String
    Dim Result          As String
    
    ' Replace escaped interval characters with a temporary symbol.
    LocalFormat = Replace(Format, _
        EscapeCharacter & IntervalSetting(DtInterval.dtMillisecond, True), _
        EscapeCharacter & DummySymbol)
    ' Replace escaped millisecond symbol: \.f[f][f] with clean millisecond symbol.
    LocalFormat = Replace(LocalFormat, _
        EscapeCharacter & MillisecondSeparator & IntervalSetting(DtInterval.dtMillisecond, True), _
        MillisecondSeparator & IntervalSetting(DtInterval.dtMillisecond, True))
    
    If InStr(LocalFormat, IntervalSetting(DtInterval.dtMillisecond, True)) > 0 Then
        ' LocalFormat contains a placeholder for milliseconds.
        If IsDateMsec(Expression) Then
            ' Expression can be formatted as date, time, and milliseconds.
            ' Convert Expression to a true Date value.
            DateValue = CDateMsec(Expression)
            ' Build full millisecond format symbol (".fff").
            MsecSymbol = MillisecondSeparator & String(3, IntervalSetting(DtInterval.dtMillisecond, True))
            SymbolLength = Len(MsecSymbol)
            
            ' Locate first occurrency of a millisecond symbol: .f[f][f][#][#]
            ' Loop from ".fff" to ".f".
            For Length = SymbolLength To 2 Step -1
                ' Shorten Symbol to ".fff", ".ff", or ".f".
                Symbol = Mid(MsecSymbol, 1, Length)
                If UBound(Split(LocalFormat, Symbol)) > 0 Then
                    ' This symbol is found.
                    ' Replace it with a formatted string of milliseconds.
                    ' Length determines how the count of milliseconds should be rounded.
                    Digits = Length - 1
                    ' Check if further significant digits are requested.
                    While Mid(Split(LocalFormat, Symbol, 2)(1), 1, 1) = MsecOptional And Digits < SymbolLength - 1
                        Symbol = Symbol & MsecOptional
                        Digits = Digits + 1
                    Wend
                    ' Extract and round the count of milliseconds in DateValue.
                    Milliseconds = RoundMillisecondMid(Millisecond(DateValue), Digits)
                    If Milliseconds >= MillisecondsPerSecond Then
                        ' A high count of milliseconds, say 990, was rounded up to 1000.
                        Milliseconds = Milliseconds Mod MillisecondsPerSecond
                        ' Return (ByRef), that a displayed value of seconds should be rounded up by 1.
                        SecondRoundUp = True
                    End If
                    ' Format decimal Milliseconds as ".0[0][0]".
                    MsecText = LTrim(Str(CDbl(VBA.Format(Milliseconds / MillisecondsPerSecond, DecimalSeparator & Replace(Mid(Symbol, 2), IntervalSetting(DtInterval.dtMillisecond, True), "0")))))
                    ' Prefix with the separator - also for a clean zero.
                    MsecText = MillisecondSeparator & Replace(MsecText, DecimalSeparator, "")
                    If Length > Len(MsecText) Then
                        MsecText = MsecText & String(Length - Len(MsecText), "0")
                    End If
                    ' Replace symbols with values.
                    Result = _
                        Split(LocalFormat, Symbol, 2)(0) & _
                        MsecText & _
                        Split(LocalFormat, Symbol, 2)(1)
                End If
                ' Replace only the first occurrence of the symbol.
                If Result <> "" Then
                    Exit For
                End If
            Next
            
            If Result = "" Then
                ' A symbol of .f[f][f] was not found.
                ' Locate first occurrency of a millisecond symbol: f[f][f]
                ' Loop from "fff" to "f".
                For Length = SymbolLength - 1 To 2 - 1 Step -1
                    ' Shorten Symbol to "fff", "ff", or "f".
                    Symbol = Mid(MsecSymbol, 2, Length)
                    If UBound(Split(LocalFormat, Symbol)) > 0 Then
                        ' This symbol is found.
                        ' Replace it with a formatted string of milliseconds.
                        ' Length determines how the count of milliseconds should be rounded.
                        Digits = Length
                        ' Extract and round the count of milliseconds in DateValue.
                        Milliseconds = Millisecond(DateValue)
                        ' Format Milliseconds as "000", "00", or "0".
                        MsecText = VBA.Format(Milliseconds, String(Digits, "0"))
                        ' Replace symbols with values.
                        Result = _
                            Split(LocalFormat, Symbol, 2)(0) & _
                            MsecText & _
                            Split(LocalFormat, Symbol, 2)(1)
                    End If
                    ' Replace only the first occurrence of the symbol.
                    If Result <> "" Then
                        Exit For
                    End If
                Next
            End If
        End If
        ' Restore from temporary symbols.
        Result = Replace(Result, EscapeCharacter & DummySymbol, EscapeCharacter & IntervalSetting(DtInterval.dtMillisecond, True))
    Else
        ' No milliseconds. Return original format expression.
        Result = Format
    End If
    
    FormatMillisecond = Result
     
End Function

' Checks an expression if it represents a date/time value
' with or without a millisecond part.
'
' 2016-09-17. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsDateMsec( _
    ByVal Expression As Variant) _
    As Boolean
  
    Dim DateText    As String
    Dim Result      As Boolean

    On Error GoTo Err_IsDateMsec
    
    ' First try IsDate(Expression). If success, we are done.
    Result = IsDate(Expression)
    If Result = False And Not IsNull(Expression) Then
        ' Try to convert Expression to a string and strip a millisecond part.
        DateText = CStr(Expression)
        ExtractMsec DateText
        ' ExtractMsec returned a cleaned string expression.
        ' Validate this.
        Result = IsDate(DateText)
    End If

    IsDateMsec = Result
  
Exit_IsDateMsec:
    Exit Function
    
Err_IsDateMsec:
    Err.Clear
    Resume Exit_IsDateMsec

End Function

' Obtain the current local timezone bias.
' IF DST is active, the bias will include the daylight bias.
'
' Note:
'   For Public use, use the function from project VBA.Timezone-Windows:
'   WtziBase.LocalBiasTimezonePresent
'
' 2016-09-12. Gustav Brock, Cactus Data ApS, CPH.
'
Private Function LocalTimeZoneBias() As Long

    Dim TzInfo  As TimeZoneInformation
    Dim TzId    As Long
    Dim Bias    As Long
    
    TzId = GetTimeZoneInformation(TzInfo)
    
    Select Case TzId
        Case TimeZoneId.Standard, TimeZoneId.Daylight
            Bias = TzInfo.Bias
            If TzId = TimeZoneId.Daylight Then
                Bias = Bias + TzInfo.DaylightBias
            End If
    End Select
    
    LocalTimeZoneBias = Bias
   
End Function

' Returns the millisecond part from Date1.
'
' 2016-09-15. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function Millisecond( _
    ByVal Date1 As Date) _
    As Integer

    Dim Milliseconds    As Integer
    
    ' Remove date part from date/time value and extract count of milliseconds.
    ' Note the use of CDec() to prevent bit errors for very large date values.
    Milliseconds = Abs(Date1 - CDec(Fix(Date1))) * MillisecondsPerDay Mod MillisecondsPerSecond
    
    Millisecond = Milliseconds

End Function

       
' This is the core timer function for milliseconds in VBA.
' Controlled by parameter ResultType, it generates one of these current values:
'
'   Date and time and milliseconds
'   Time and milliseconds
'   Milliseconds only (default)
'
' with millisecond resolution.
'
' 2016-09-12. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function Msec( _
    Optional ByVal ResultType As DtMsecResult = DtMsecResult.dtMillisecondOnly) _
    As Date

    Static SysTime      As SystemTime
    Static MsecInit     As Long

    Dim Result          As DateTime
    Dim MsecValue       As Date
    Dim DateValue       As Date
    Dim Milliseconds    As Integer
    Dim TimezoneBias    As Long
    Dim MsecCount       As Long
    Dim MsecCurrent     As Long
    Dim MsecOffset      As Long
  
    ' Set resolution of timer to 1 ms.
    TimeBeginPeriod 1
    
    MsecCurrent = TimeGetTime()
  
    If MsecInit = 0 Or MsecCurrent < MsecInit Then
        ' Initialize.
        ' Get bias for local time zone respecting
        ' current setting for daylight savings.
        TimezoneBias = LocalTimeZoneBias()
        ' Get current UTC system time.
        Call GetSystemTime(SysTime)
        Milliseconds = SysTime.wMilliseconds
        ' Repeat until GetSystemTime retrieves next count of milliseconds.
        ' Then retrieve and store count of milliseconds from launch.
        Do
            Call GetSystemTime(SysTime)
        Loop Until SysTime.wMilliseconds <> Milliseconds
        MsecInit = TimeGetTime()
        ' Adjust UTC to local system time by correcting for time zone bias.
        SysTime.wMinute = SysTime.wMinute - TimezoneBias
        ' Note: SysTime may now contain an invalid (zero or negative) minute count.
        ' However, the minute count is acceptable by TimeSerial().
    Else
        ' Retrieve offset from initial time to current time.
        MsecOffset = MsecCurrent - MsecInit
    End If
    
    With SysTime
        ' Now, current system time equals initial system time corrected for
        ' time zone bias.
        MsecCount = (.wMilliseconds + MsecOffset)
        Select Case ResultType
            Case DtMsecResult.dtTimeMillisecond, DtMsecResult.dtDateTimeMillisecond
                ' Calculate the time to add as a date/time value with millisecond resolution.
                MsecValue = MsecCount / MillisecondsPerSecond / SecondsPerDay
                ' Add to this the current system time.
                Result.Time = MsecValue + TimeSerial(.wHour, .wMinute, .wSecond)
                If ResultType = DtMsecResult.dtDateTimeMillisecond Then
                    ' Include the current system date.
                    Result.Date = DateSerial(.wYear, .wMonth, .wDay)
                End If
            Case Else
                ' Calculate the millisecond part as a date value with millisecond resolution.
                MsecValue = (MsecCount Mod MillisecondsPerSecond) / MillisecondsPerSecond / SecondsPerDay
                ' Return the millisecond part only.
                Result.Time = MsecValue
        End Select
    End With
    
    DateValue = JoinDateTime(Result)
    
    Msec = DateValue
  
End Function

' Returns the difference in milliseconds between Date1 and Date2.
' Accepts any valid Date value including milliseconds from
'   100-01-01 00:00:00.000 to 9999-12-31 23:59:59.999 or reverse
' which will return from
'   -312,413,759,999,999 to 312,413,759,999,999
'
' 2016-09-14. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function MsecDiff( _
    ByVal Date1 As Date, _
    ByVal Date2 As Date) _
    As Double

    Dim Milliseconds As Double
  
    ' Convert native date values to linear date values.
    Call ConvDateToTimespan(Date1)
    Call ConvDateToTimespan(Date2)
    ' Convert to milliseconds and find the difference.
    Milliseconds = CDec(Date2 * MillisecondsPerDay) - CDec(Date1 * MillisecondsPerDay)
    
    MsecDiff = Milliseconds
  
End Function

' Returns the date/time value of Millisecond rounded to integer milliseconds.
' Typical usage:
'   MsecValue = MsecSerial(milliseconds)
'   DateTimeMsecValue = MsecSerial(milliseconds, DateTimeValue)
'
' Values of Millisecond beyond +/-999 are still converted to valid date values.
' Accepts, with no Date1, any input in the interval:
'   -56,802,297,600,000 to 255,611,462,399,999
' Possible return value is any Date value from:
'   100-1-1 00:00:00.000 to 9999-12-31 23:59:59.999
'
' If a Date1 is specified, Millisecond will be added to this, and the
'   acceptable input range is shifted accordingly.
' Min. Date1:
'    100-01-01 00:00:00.000. Will accept Millisecond of
'   0 to 312,413,759,999,999
' Max. Date1:
'   9999-12-31 23:59:59.999. Will accept Millisecond of
'   -312.413.759.999.999 to 0
'
' Resulting return dates must be within the limits of datatype Date:
'   100-1-1 00:00:00.000 to 9999-12-31 23:59:59.999
'
' 2016-09-12. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function MsecSerial( _
    ByVal Millisecond As Double, _
    Optional ByVal Date1 As Date) _
    As Date
  
    Dim Timespan        As Date
    Dim Milliseconds    As Double
    Dim Result          As Date
  
    ' Catch and return error in case of overflow
    ' when DateFromTimespan is called.
    On Error GoTo Err_MsecSerial
  
    ' Convert (invalid) numeric negative date values less than one day.
    EmendTime Date1
    
    If Millisecond = 0 Then
        ' Nothing to add. Return base date.
        Result = Date1
    Else
        ' Convert Date1 to a timespan (linear date value).
        Timespan = DateToTimespan(Date1)
        ' Convert the timespan to milliseconds and adjust with Millisecond.
        Milliseconds = CDbl(Timespan * MillisecondsPerDay) + CDec(Fix(Millisecond))
        ' Convert milliseconds to a timespan (linear date value).
        Timespan = CDate(Milliseconds / MillisecondsPerDay)
        ' Convert the timespan to a date value.
        Result = DateFromTimespan(Timespan)
    End If
        
    MsecSerial = Result
  
Exit_MsecSerial:
    Exit Function
    
Err_MsecSerial:
    Err.Raise Err.Number
    Resume Exit_MsecSerial
  
End Function

' Cleans Expression for the date and time value and, if possible,
' returns the millisecond value of Expression.
'
' Wraps ExtractMsec and acts as sister function for:
'   DateValueMsec and
'   TimeValueMsec
'
' 2016-09-17. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function MsecValueMsec( _
    ByVal Expression As String) _
    As Date
    
    Dim Result As Date
    
    Result = ExtractMsec(Expression)
    ' Validate Expression.
    TimeValue Expression
    
    MsecValueMsec = Result
  
End Function

' Returns the current local date and time including milliseconds.
'
' 2016-09-17. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function NowMsec() As Date
  
    Dim Result  As Date
    
    Result = Msec(DtMsecResult.dtDateTimeMillisecond)
    
    NowMsec = Result

End Function

' Rounds by 4/5 an integer count of milliseconds
' between -990 and 999 to a value between -1000 and 1000.
'
' Values outside +/-1000 will be limited to +/-1000.
'
' Parameter Digits determine the rounding:
'   1       => Rounding to hundreds
'   2       => Rounding to tens
'   <= 0    => Rounding to ones (no rounding)
'   >= 3    => Rounding to ones (no rounding)
'
' Examples:
'   RoundMillisecondMid(7)          ->    7
'   RoundMillisecondMid(57)         ->   57
'   RoundMillisecondMid(957)        ->  957
'   RoundMillisecondMid(7, 2)       ->   10
'   RoundMillisecondMid(57, 2)      ->   60
'   RoundMillisecondMid(957, 2)     ->  960
'   RoundMillisecondMid(7, 1)       ->    0
'   RoundMillisecondMid(57, 1)      ->  100
'   RoundMillisecondMid(957, 1)     -> 1000
'   RoundMillisecondMid(-57, 2)     ->  -60
'   RoundMillisecondMid(-957, 2)    -> -960
'
' 2019-11-04. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function RoundMillisecondMid( _
    ByVal Value As Integer, _
    Optional ByVal Digits As Integer) _
    As Integer
    
    Const MaxDigits As Integer = 3
    
    Dim Result      As Integer
    
    Select Case Value
        Case Is = 0
            ' Nothing to round.
        Case Is < -MillisecondsPerSecond
            ' Limit value.
            Result = -MillisecondsPerSecond
        Case Is > MillisecondsPerSecond
            ' Limit value.
            Result = MillisecondsPerSecond
        Case Else
            If Digits <= 0 Or Digits >= MaxDigits Then
                ' No rounding.
                Result = Value
            Else
                ' Round value.
                Digits = MaxDigits - Digits
                Result = Val(Format(Value / 10 ^ Digits, "0")) * 10 ^ Digits
            End If
    End Select
    
    RoundMillisecondMid = Result
    
End Function

' Rounds off Date1 to the second by removing a millisecond portion.
'
' 2019-11-04. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub RoundOffMilliseconds( _
    ByRef Date1 As Date)
    
    ConvDateToTimespan Date1
    Date1 = Fix(Date1 * CDec(SecondsPerDay)) / SecondsPerDay
    ConvTimespanToDate Date1

End Sub

' Returns the current local time including milliseconds.
'
' 2016-09-17. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function TimeMsec() As Date
  
    Dim Result  As Date
    
    Result = Msec(DtMsecResult.dtTimeMillisecond)
    
    TimeMsec = Result

End Function

' Returns the count of seconds from Midnight with millisecond resolution.
' Mimics Timer which, however, returns the value with a resolution of 10 milliseconds.
'
' 2016-09-17. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function TimerMsec() As Double

    Dim Result  As Double
  
    Result = Msec(DtMsecResult.dtTimeMillisecond) * SecondsPerDay

    TimerMsec = Result

End Function

' Returns the date/time value of the combined parameters for
' hour, minute, second and millisecond.
' Accepts decimal input for seconds within the range of Integer.
' The fraction of second is rounded to integer milliseconds.
'
' If input values for second or millisecond beyond those of data type
' Integer are expected, use MsecSerial() or DateAddMsec().
'
' 2016-09-17. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function TimeSerialMsec( _
    ByVal Hour As Integer, _
    ByVal Minute As Integer, _
    ByVal Second As Double, _
    Optional ByVal Millisecond As Integer) _
    As Date

    Dim Seconds         As Integer
    Dim Milliseconds    As Double
    Dim Result          As Date
  
    ' Raise error if integer part of second exceeds Integer datatype.
    Seconds = Fix(Second)
  
    ' Round decimal part of parameter Second by 4/5 to an integer
    ' count of milliseconds and add the value of parameter Millisecond.
    Milliseconds = Millisecond + Fix((CDec(Second) - Seconds) * MillisecondsPerSecond + Sgn(Seconds) * CDec(0.5))
    
    Result = MsecSerial(Milliseconds, TimeSerialDate(Hour, Minute, Seconds))

    TimeSerialMsec = Result
    
End Function

' Cleans Expression for milliseconds and a date part
' and, if possible, returns the time value of Expression.
'
' Note:
'   Will raise error 13 if the cleaned Expression
'   does not represent a valid time value.
'
' 2016-09-17. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function TimeValueMsec( _
    ByVal Expression As String) _
    As Date
  
    Dim Result As Date
  
    On Error GoTo Err_TimeValueMsec
  
    ' Strip a millisecond part from Expression.
    ExtractMsec Expression
    ' Convert Expression to a time value with no date part.
    Result = TimeValue(Expression)
    
    TimeValueMsec = Result

Exit_TimeValueMsec:
    Exit Function
    
Err_TimeValueMsec:
    Err.Raise Err.Number
    Resume Exit_TimeValueMsec
  
End Function


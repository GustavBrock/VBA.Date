Attribute VB_Name = "DateText"
Option Explicit
'
' DateText
' Version 1.3.5
'
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.Date
'
' Functions for formatting of date/time and parsing text for date/time values.
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

' Converts a "Swatch Internet Time" string to
' the count of .beats it holds.
'
' Internet Time has the format "@000".
'
' Reference:
'   https://www.swatch.com/en_us/internet-time/
'
' Examples:
'   "@000"  ->   0
'   "@459"  -> 459
'   "@8734" -> 873
'   "623"   -> 623
'   "7391"  -> 739
'
' 2018-01-25. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function CBeatInternetTime( _
    ByVal Expression As String) _
    As Integer
    
    ' Header to strip: @.
    Const HeaderAscii   As Integer = 64
    ' Maximum length of Internet Time value.
    Const Length        As Integer = 3
    
    Dim Start   As Integer
    Dim Result  As Integer
    
    If Expression <> "" Then
        If Asc(Expression) = HeaderAscii Then
            Start = 2
        Else
            Start = 1
        End If
        ' Extract and convert .beats.
        Result = Val(Mid(Expression, Start, Length))
    End If
    
    CBeatInternetTime = Result
    
End Function

' Converts a "Military Date Time Group" (DTG) formatted string
' to a date value.
'
' If IgnoreTimezone is True, the timezone identifier is
' ignored, and the date/time value returned as is.
' If IgnoreTimezone is False, the date/time value is
' converted to UTC.
'
' DTG must be formatted as "ddhhnnZmmmyy".
' If a non-parsable string is passed, an error is raised.
'
' Example:
'   071943ZFEB09 represents 2009-02-07 19:43:00
'
' 2016-01-21. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function CDateDtg( _
    ByVal Dtg As String, _
    Optional ByVal IgnoreTimezone As Boolean) _
    As Date

    Const DtgLength     As Integer = 12
    
    Dim Year            As Integer
    Dim Month           As Integer
    Dim Day             As Integer
    Dim Hour            As Integer
    Dim Minute          As Integer
    Dim Second          As Integer
    Dim Offset          As Integer
    
    Dim Result          As DateTime
    Dim ResultDate      As Date
        
    Dtg = Trim(Dtg)
    If Len(Dtg) = DtgLength Then
        Year = Val(Mid(Dtg, 11, 2))
        Month = MonthFromInvariant(Mid(Dtg, 8, 3))
        Day = Val(Mid(Dtg, 1, 2))
        Hour = Val(Mid(Dtg, 3, 2))
        Minute = Val(Mid(Dtg, 5, 2))
        Second = 0
        
        If IgnoreTimezone = False Then
            Offset = MilitaryTimezone(Mid(Dtg, 7, 1))
        End If
        ' Split the time and offset into its date and time parts.
        Result = SplitDateTime(DateAdd("h", Offset, TimeSerial(Hour, Minute, Second)))
        ' Sum the date parts.
        Result.Date = Result.Date + DateSerial(Year, Month, Day)
        ' Assemble the date and the time part.
        ResultDate = JoinDateTime(Result)
    Else
        Err.Raise DtError.dtTypeMismatch
        Exit Function
    End If
    
    CDateDtg = ResultDate

End Function

    
' Converts a "Swatch Internet Time" string to
' the corresponding time it holds.
'
' Internet Time has the format "@000".
'
' Timezone of the time is BMT, "Biel Meantime"
' which equals UTC+01:00.
'
' Reference:
'   https://www.swatch.com/en_us/internet-time/
'
' The result is by default rounded to the second.
' Optionally, by passing parameter RoundSeconds as False,
' deciseconds will be preserved: 4, 8, 2, 6, or 0.
'
' Examples:
'   RoundSeconds = True:
'   "@000"  -> 00:00:00
'   "@001"  -> 00:01:26
'   "@002"  -> 00:02:53
'   "@500"  -> 12:00:00
'   "@998"  -> 23:57:07
'   "@999"  -> 23:58:34
'
'   RoundSeconds = False:
'   "@000"  -> 00:00:00.000
'   "@001"  -> 00:01:26.400
'   "@002"  -> 00:02:52.800
'   "@998"  -> 23:57:07.200
'   "@999"  -> 23:58:33.800
'
' 2018-01-25. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function CDateInternetTime( _
    ByVal Expression As String, _
    Optional ByVal RoundSeconds As Boolean = True) _
    As Date
    
    Dim Beats   As Integer
    Dim Result  As Date
    
    Beats = CBeatInternetTime(Expression)
    Result = DateBeat(Beats, RoundSeconds)
    
    CDateInternetTime = Result
    
End Function

' Converts an ISO 8601 formatted date/time string to a date value.
'
' A timezone info is ignored.
' Optionally, a millisecond part can be ignored.
'
' Examples:
'   2029-02-17T19:43:08 +01.00  -> 2029-02-17 19:43:08
'   2029-02-17T19:43:08         -> 2029-02-17 19:43:08
'   ' IgnoreMilliseconds = False
'   2029-02-17T19:43:08.566     -> 2029-02-17 19:43:08.566
'   ' IgnoreMilliseconds = True
'   2029-02-17T19:43:08.566     -> 2029-02-17 19:43:08.000
'
' 2017-05-24. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function CDateIso8601( _
    ByVal Expression As String, _
    Optional ByVal IgnoreMilliseconds As Boolean) _
    As Date

    Const Iso8601Separator  As String = "T"
    Const NeutralSeparator  As String = " "

    ' Length of ISO 8601 date/time string like: 2029-02-17T19:43:08 [+00.00]
    Const Iso8601Length     As Integer = 19
    ' Length of ISO 8601 date/time string like: 2029-02-17T19:43:08.566
    Const Iso8601MsecLength As Integer = 23
    
    Dim Value       As String
    Dim Result      As Date
    
    Value = Replace(Expression, Iso8601Separator, NeutralSeparator)
    If InStr(Expression, MillisecondSeparator) <> Iso8601Length + 1 Then
        IgnoreMilliseconds = True
    End If
    
    If IgnoreMilliseconds = False Then
        Result = CDateMsec(Left(Value, Iso8601MsecLength))
    Else
        Result = CDate(Left(Value, Iso8601Length))
    End If
    
    CDateIso8601 = Result
    
End Function

' Converts a string expression for a sport result with
' 1/100 of seconds or 1/1000 of seconds to a date value
' including milliseconds.
'
' Example:
'   "3:12:23.48"    ->  03:12:23.480
'     "20:06.80"    ->  00:20:06.800
'        "19.56"    ->  00:00:19.560
'       "49.120"    ->  00:00:49.120
'       "23.328"    ->  00:00:23.328
'
' 2018-01-24. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function CDateRaceTime( _
    ByVal RaceTime As String) _
    As Date
    
    Dim Values      As Variant
    Dim Hour        As Integer
    Dim Minute      As Integer
    Dim Second      As Integer
    Dim Millisecond As Integer
    Dim Result      As Date
    
    Values = Split(RaceTime, MillisecondSeparator)
    Select Case UBound(Values)
        Case 0
            ' No split seconds.
        Case 1
            Millisecond = Val(Left(Values(1) & "00", 3))
        Case Else
            ' Invalid expression.
    End Select
    
    If UBound(Values) <= 1 Then
        ' Split time part.
        Values = Split(Values(0), TimeSeparator)
        Select Case UBound(Values)
            Case 0
                Second = Val(Values(0))
            Case 1
                Second = Val(Values(1))
                Minute = Val(Values(0))
            Case 2
                Second = Val(Values(2))
                Minute = Val(Values(1))
                Hour = Val(Values(0))
            Case Else
                ' Invalid expression.
                Millisecond = 0
        End Select
        Result = TimeSerialMsec(Hour, Minute, Second, Millisecond)
    End If
    
    CDateRaceTime = Result

End Function

' Parse a text expression for a possible date value.
' Will parse dd/mm or mm/dd as to the local settings.
' Returns zero date, 00:00:00, for invalid expressions like Null.
'
' Examples that can be parsed:
'   "03-08-2020"
'   "03.08.2020"
'   "1026/2006"
'   "O1/19/2007T09:00"
'   "02I21949"
'   "O6/13/1952"
'   "07/27:1956"
'   "07/7:1956"
'   "7/07:1956"
'   "7/7:1956"
'   "06/042/1952"
'
' 2019-09-18. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function CDateText( _
    ByVal Expression As String) _
    As Date

    ' Length of clean date string.
    Const DefaultLength As Long = 8
    ' Format of string expression for a date.
    Const DefaultFormat As String = "@@/@@/@@@@"
    
    Dim Text            As String
    Dim Value           As Date
    
    If IsNull(Expression) Or IsEmpty(Expression) Or IsObject(Expression) Then
        Exit Function
    Else
        Text = CStr(Expression)
    End If
    
    ' Replace date typos.
    Text = Replace(Replace(Replace(Text, "O", "0"), "l", "1"), "I", "1")
    ' Replace separator typos.
    Text = Replace(Replace(Replace(Replace(Text, "_", "/"), ".", "/"), "-", "/"), ":", "/")
    
    ' Attempt a convert after basic corrections.
    If IsDate(Text) Then
        Value = DateValue(Text)
    ElseIf Len(Text) = DefaultLength Then
        ' Convert an expression without separators.
        Text = Format(Text, DefaultFormat)
        If IsDate(Text) Then
            Value = DateValue(Text)
        End If
    Else
        ' Correct for missing leading zero.
        If Len(Text) - 2 < DefaultLength Then
            Text = "0" & Join(Split(Replace(Text, "/0", "/"), "/"), "/0")
        End If
        
        ' Remove date separator and spaces.
        Text = Replace(Replace(Text, "/", ""), " ", "")
        If Not IsNumeric(Text) Then
            ' Remove trailing text.
            Text = Format(Val(Text), String(DefaultLength, "0"))
        End If
        ' Remove month typos.
        If Len(Text) > DefaultLength Then
            Text = Left(Text, 4) & Right(Text, 4)
        End If
        ' Apply date format.
        Text = Format(Text, DefaultFormat)
        
        ' Convert to date if possible.
        If IsDate(Text) Then
            Value = DateValue(Text)
        End If
    End If

    CDateText = Value

End Function

' Converts a US formatted date/time string to a date value.
'
' Examples:
'   7/6/2016 7:00 PM    -> 2016-07-06 19:00:00
'   7/6 7:00 PM         -> 2018-07-06 19:00:00  ' Current year is 2018.
'   7/6/46 7:00 PM      -> 1946-07-06 19:00:00
'   8/9-1982 9:33       -> 1982-08-09 09:33:00
'   2/29 14:21:56       -> 2039-02-01 14:21:56  ' Month/year.
'   2/39 14:21:56       -> 1939-02-01 14:21:56  ' Month/year.
'   7/6/46 7            -> 1946-07-06 00:00:00  ' Cannot read time.
'   7:32                -> 1899-12-30 07:32:00  ' Time value only.
'   7:32 PM             -> 1899-12-30 19:32:00  ' Time value only.
'   7.32 PM             -> 1899-12-30 19:32:00  ' Time value only.
'   14:21:56            -> 1899-12-30 14:21:56  ' Time value only.
'
' 2018-03-31. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function CDateUs( _
    ByVal Expression As String) _
    As Date
    
    Const PartSeparator As String = " "
    Const DateSeparator As String = "/"
    Const DashSeparator As String = "-"
    Const MaxPartCount  As Integer = 2

    Dim Parts           As Variant
    Dim DateParts       As Variant
    
    Dim DatePart        As Date
    Dim TimePart        As Date
    Dim Result          As Date
    
    ' Split expression into maximum two parts.
    Parts = Split(Expression, PartSeparator, MaxPartCount)
    
    
    If IsDate(Parts(0)) Then
        ' A date or time part is found.
        ' Replace dashes with slashes.
        Parts(0) = Replace(Parts(0), DashSeparator, DateSeparator)
        If InStr(1, Parts(0), DateSeparator) > 1 Then
            ' A date part is found.
            DateParts = Split(Parts(0), DateSeparator)
            If UBound(DateParts) = 2 Then
                ' The date includes year.
                DatePart = DateSerial(DateParts(2), DateParts(0), DateParts(1))
            Else
                If IsDate(CStr(Year(Date)) & DateSeparator & Join(DateParts, DateSeparator)) Then
                    ' Use current year.
                    DatePart = DateSerial(Year(Date), DateParts(0), DateParts(1))
                Else
                    ' Expression contains month/year.
                    DatePart = CDate(Join(DateParts, DateSeparator))
                End If
            End If
            If UBound(Parts) = 1 Then
                If IsDate(Parts(1)) Then
                    ' A time part is found.
                    TimePart = CDate(Parts(1))
                End If
            End If
        Else
            ' A time part it must be.
            ' Concatenate an AM/PM part if present.
            TimePart = CDate(Join(Parts, PartSeparator))
        End If
    End If
    
    Result = DatePart + TimePart
        
    CDateUs = Result

End Function

' Escapes each character in Format and returns the result by reference.
' Optionally, only escapes digits and the millisecond separator.
'
' Examples:
'   All characters:
'       "milliseconds: 78"  -> "\m\i\l\l\i\s\e\c\o\n\d\s\:\ \7\8"
'   Digits only:
'       "nn:ss.078"         -> "nn:ss\.\0\7\8"
'
' 2016-12-11. Gustav Brock, Cactus Data ApS, CPH.
'
Private Sub EscapeFormat( _
    ByRef Format As String, _
    Optional ByVal DigitsOnly As Boolean)

    Dim Length      As Integer
    Dim Position    As Integer
    Dim Character   As String
    Dim Result      As String
    
    Length = Len(Format)

    For Position = 1 To Length
        Character = Mid(Format, Position, 1)
        If IsNumeric(Character) Or Character = MillisecondSeparator Or Not DigitsOnly Then
            Result = Result & EscapeCharacter
        End If
        Result = Result & Character
    Next
    
    Format = Result
    
End Sub

' Formats the output from AgeMonthsDays.
'
' 2020-10-05. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function FormatAgeYearsMonthsDays( _
    ByVal DateOfBirth As Date, _
    Optional ByVal AnotherDate As Variant) _
    As String

    Dim Years       As Integer
    Dim Months      As Integer
    Dim Days        As Integer
    Dim YearsLabel  As String
    Dim MonthsLabel As String
    Dim DaysLabel   As String
    Dim Result      As String
    
    Months = AgeMonthsDays(DateOfBirth, AnotherDate, Days)
    Years = Months \ MonthsPerYear
    Months = Months Mod MonthsPerYear
    
    YearsLabel = "year" & IIf(Years = 1, "", "s")
    MonthsLabel = "month" & IIf(Months = 1, "", "s")
    DaysLabel = "day" & IIf(Days = 1, "", "s")
    
    ' Concatenate the parts of the output.
    Result = CStr(Years) & " " & YearsLabel & ", " & CStr(Months) & " " & MonthsLabel & ", " & CStr(Days) & " " & DaysLabel
    
    FormatAgeYearsMonthsDays = Result
    
End Function

' Converts and formats a count of .beats
' to the "Swatch Internet Time".
'
' Internet Time has the format "@000".
'
' Reference:
'   https://www.swatch.com/en_us/internet-time/
'
' Examples:
'    28 -> "@028"
'   649 -> "@649"
'
' 2018-01-25. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function FormatBeatInternetTime( _
    ByVal Beats As Long) _
    As String
    
    Const Format    As String = EscapeCharacter & "@000"
    
    Dim Value       As Long
    Dim Result      As String
    
    Value = Abs(Beats) Mod BeatsPerDay
    Result = VBA.Format(Value, Format)
    
    FormatBeatInternetTime = Result
    
End Function

' Returns a date as a "Military Date Time Group" (DTG) formatted string.
' The format is: ddhhnnzmmmyy
' Optionally, a timezone offset can be passed, either in full hours
' a a value from -12 to 12 or as a DTG time code letter.
'
' Example:
'   2012-01-06 18:30 -01:00 -> "061830NJAN12"
'
' 2015-11-25. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function FormatDateDtg( _
    ByVal Date1 As Date, _
    Optional TimezoneOffset As Variant = "Z") _
    As String
    
    Const Identifiers   As String = "ABCDEFGHIKLMNOPGRSTUVWXYZ"
    
    Dim Identifier  As String
    Dim Dtg         As String
    
    Identifier = Trim(Left(Nz(TimezoneOffset, "Z"), 1))
    If InStr(Identifiers, Identifier) > 0 Then
        ' DTG time code letter identifier passed for timezone.
    Else
        ' Convert numeric timezone offset to DTG time code letter.
        Identifier = MilitaryTimeCodeLetter(TimezoneOffset)
    End If
    
    Dtg = _
        Format(Date1, "ddhhnn") & _
        Identifier & _
        UCase(MonthNameInvariant(Month(Date1), True)) & _
        Format(Date1, "yy")
    
    FormatDateDtg = Dtg

End Function

' Converts and formats the time part of a date value
' to the "Swatch Internet Time".
'
' Internet Time has the format "@000".
'
' Examples:
'   15:34:20  -> "@649"
'   01:14:32  -> "@052"
'
' Reference:
'   https://www.swatch.com/en_us/internet-time/
'
' 2018-01-25. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function FormatDateInternetTime( _
    ByVal Date1 As Date, _
    Optional ByVal RoundSeconds As Boolean = True) _
    As String
    
    Dim Beats       As Integer
    Dim Result      As String
    
    Beats = Beat(Date1, RoundSeconds)
    Result = FormatBeatInternetTime(Beats)
    
    FormatDateInternetTime = Result
    
End Function

' Returns, for a date value, a formatted string expression with milliseconds
' according to ISO-8601.
' Optionally, a T is used as separator between the date and time parts.
'
' Typical usage:
'
'   FormatDateIso8601(NowMsec)
'   ->  2017-03-27 13:48:32.017
'
'   FormatDateIso8601(NowMsec, True)
'   ->  2017-03-28T13:41:37.243
'
' 2017-04-28. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function FormatDateIso8601( _
    ByVal Expression As Variant, _
    Optional ByVal TSeparator As Boolean) _
    As String
  
    Const Iso8601Separator  As String = EscapeCharacter & "T"
    Const NeutralSeparator  As String = EscapeCharacter & " "
    
    Dim FormatDateParts     As Variant
    Dim FormatTimeParts     As Variant
    Dim Format              As String
    Dim Result              As String
    
    ' Compose the parts of the format string.
    FormatDateParts = Array( _
        IntervalSetting(DtInterval.dtYear), _
        String(2, IntervalSetting(DtInterval.dtMonth)), _
        String(2, IntervalSetting(DtInterval.dtDay)))
    FormatTimeParts = Array( _
        String(2, IntervalSetting(DtInterval.dtHour)), _
        String(2, IntervalSetting(DtInterval.dtMinute)), _
        String(2, IntervalSetting(DtInterval.dtSecond)) & _
        EscapeCharacter & MillisecondSeparator & _
        String(3, IntervalSetting(DtInterval.dtMillisecond, True)))
    
    If IsDateMsec(Expression) Then
        ' Assemble the format string.
        Format = _
            Join(FormatDateParts, IsoDateSeparator) & _
            IIf(TSeparator, Iso8601Separator, NeutralSeparator) & _
            Join(FormatTimeParts, TimeSeparator)
        ' Get the formatted date value.
        Result = FormatExt(Expression, Format)
    End If
    
    FormatDateIso8601 = Result
  
End Function

' Formats date, time, and milliseconds to any combination of intervals,
' native as well as extended.
' For options for intervals, see Enum DtInterval.
'
' Native intervals and non-date expressions are passed to and
' handled by Format.
'
' 2018-09-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function FormatExt( _
    ByVal Expression As Variant, _
    Optional ByVal Format As String, _
    Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbSunday, _
    Optional ByVal FirstWeekOfYear As VbFirstWeekOfYear = vbFirstJan1) _
    As String
    
    Const SymbolLengthMax   As Integer = 5
    Const SymbolLengthMin   As Integer = 1
    Const IgnoreCharacter   As String = "*"
    
    Dim Interval            As DtInterval
    Dim Symbol              As String
    Dim SymbolLength        As Integer
    Dim SymbolStart         As Integer
    Dim SymbolLead          As String
    Dim FormatSearch        As String
    Dim DateValue           As Date
    Dim TextValue           As String
    Dim DisplayMilliseconds As Boolean
    Dim SecondRoundUp       As Boolean
    Dim Result              As String
    
    If IsNamedFormat(Format) Then
        ' Call native Format.
        Result = VBA.Format(Expression, Format, FirstDayOfWeek, FirstWeekOfYear)
    ElseIf IsDateMsec(Expression) Then
    
        DateValue = CDateMsec(Expression)
        
        ' Replace the special format "ttttt" with its equivalent.
        Format = Replace(Format, "ttttt", "hh:nn:ss")
        
        ' Search for the longest symbols and down to the shortest.
        For SymbolLength = SymbolLengthMax To SymbolLengthMin Step -1
            ' Look up all symbols having the length of SymbolLength.
            For Interval = DtInterval.[_First] To DtInterval.[_Last]
                Symbol = IntervalSetting(Interval, True)
                If Len(Symbol) = SymbolLength Then
                    If IsIntervalSetting(Symbol, False) Then
                        ' Native symbol.
                        ' Will be handled by VBA.Format natively.
                    Else
                        ' Find and format all occurencies of this symbol in Format.
                        
                        Do
                            ' To leave escaped character untouched, replace these with a
                            ' character that will not match the first character of symbol.
                            FormatSearch = Replace(Format, EscapeCharacter & Left(Symbol, 1), EscapeCharacter & IgnoreCharacter)
                            ' Look up Symbol.
                            SymbolStart = InStr(1 + SymbolStart + Len(Symbol) - 1, FormatSearch, Symbol, vbTextCompare)
                            If SymbolStart > 0 Then
                                ' This Symbol was found.
                                ' Check the previous character - except if we are
                                ' positioned at the very start of Format (actually FormatSearch).
                                SymbolLead = Mid(FormatSearch, SymbolStart - Abs(SymbolStart > 1), 1)
                                If SymbolLead = EscapeCharacter Then
                                    ' This occurency of Symbol is a literal.
                                    ' Ignore, and search for other occurencies.
                                Else
                                    If IntervalValue(Symbol, True) = DtInterval.dtMillisecond Then
                                        ' Milliseconds require special handling.
                                        ' Format one occurrency of a millisecond part of Format and
                                        ' replace the complete Format.
                                        ' SecondRoundUp will return True if seconds of DateValue
                                        ' should be increased by one.
                                        Format = FormatMillisecond(DateValue, Format, SecondRoundUp)
                                        ' Escape the millisecond separator and all digits of Format to prevent
                                        ' zeroes to be replaced by numerical value of DateValue by Format.
                                        EscapeFormat Format, True
                                        ' Reset search entry to the start position of Format.
                                        SymbolStart = 1
                                        ' Set flag, that milliseconds should be rounded off to prevent that
                                        ' the millisecond value will cause a round up of displayed seconds.
                                        DisplayMilliseconds = True
                                    Else
                                        ' This is not a millisecond symbol.
                                        ' Replace the symbol in Format with the formatted value of Symbol.
                                        TextValue = CStr(DatePartExt(Symbol, DateValue))
                                        ' Escape each and every character of TextValue.
                                        EscapeFormat TextValue
                                        ' Split Format in head and tail and insert the formatted value.
                                        ' Join the parts for the next loop.
                                        Format = _
                                            Mid(Format, 1, SymbolStart - 1) & _
                                            TextValue & _
                                            Mid(Format, SymbolStart + Len(Symbol))
                                    End If
                                End If
                            End If
                        Loop Until SymbolStart = 0
                        
                    End If
                End If
            Next
        Next
        If DisplayMilliseconds = True Then
            ' Milliseconds are displayed. Prevent a round up of displayed seconds.
            RoundOffMilliseconds DateValue
        End If
        If SecondRoundUp = True Then
            ' Milliseconds were rounded up to 1000 ms, thus displayed as zero.
            ' Increase seconds by one.
            DateValue = DateAddExt(IntervalSetting(DtInterval.dtSecond), 1, DateValue)
        End If
        ' Extended symbols (if any) have been formatted.
        ' Call Format to format the native symbols.
        Result = VBA.Format(DateValue, Format, FirstDayOfWeek, FirstWeekOfYear)
    Else
        ' No date or time values to format.
        ' Call native Format.
        Result = VBA.Format(Expression, Format, FirstDayOfWeek, FirstWeekOfYear)
    End If
    
    FormatExt = Result
    
End Function

' Formats a time duration rounded to 1/100 second with trailing zeroes
' and with no leading hours and minutes if these are zero.
' This format is typical for sports results.
'
' Examples:
'   1:02:07.803 ->  1:02:07.80
'     02:07.803 ->     2:07.80
'        14.216 ->       14.22
'
' 2016-12-11. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function FormatRaceTime( _
    ByVal Expression As Variant) _
    As String
    
    Const FormatHour    As String = "h:n"
    Const FormatMinute  As String = "n:s"
    Const FormatBase    As String = "s.ff"
    
    Dim Duration    As Date
    Dim Format      As String
    Dim Result      As String
    
    If IsDateMsec(Expression) Then
        Duration = CDateMsec(Expression)
        If Hour(Duration) > 0 Then
            Format = FormatHour & FormatMinute
        ElseIf Minute(Duration) > 0 Then
            Format = FormatMinute
        End If
        Format = Format & FormatBase
        Result = FormatExt(Expression, Format)
    End If
    
    FormatRaceTime = Result
    
End Function

' Returns the sign of Expression, + or - for positive or negative
' values, or a space for zero.
' If ZeroPlus is True, + will be returned for values of zero.
'
' For a non-numeric value, a space is returned.
'
' Examples:
'   0.78    -> "+"
'   "-23.9" -> "-"
'   Null    -> " "
'   Date()  -> "+"
'   Time()  -> "+"
'   -Time() -> "-"
'   "Yes"   -> " "
'   0       -> " "  ' ZeroPlus = False
'   0       -> "+"  ' ZeroPlus = True
'
' 2019-12-09. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function FormatSign( _
    ByVal Expression As Variant, _
    Optional ByVal ZeroPlus As Boolean) _
    As String
    
    ' Possible return values.
    Const Signs As String = "- +"
    
    ' Always return exactly one character.
    Dim Sign    As String * 1
    Dim Index   As Integer
    
    If IsNumeric(Expression) Or IsDate(Expression) Then
        Index = Sgn(Expression)
        If Index = 0 And ZeroPlus = True Then
            Index = 1
        End If
    End If
    ' Pick the sign from the options.
    Sign = Mid(Signs, 2 + Index)
    
    FormatSign = Sign

End Function

' Obtain the system date format without API calls.
'
' 2021-01-26. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function FormatSystemDate() As String

    Const TestDate  As Date = #1/2/3333#
    
    Dim DateFormat  As String
    
    DateFormat = Replace(Replace(Replace(Replace(Replace(Format(TestDate), "3", "y"), "1", "m"), "2", "d"), "0m", "mm"), "0d", "dd")

    FormatSystemDate = DateFormat
    
End Function

' Obtain the system date separator without API calls.
'
' 2021-01-26. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function FormatSystemDateSeparator() As String

    Dim Separator   As String
    
    Separator = Format(Date, "/")

    FormatSystemDateSeparator = Separator
    
End Function

' Obtain the system time format without API calls.
'
' 2021-01-26. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function FormatSystemTime() As String

    Const TestTime  As Date = #1:02:03 AM#
    
    Dim TimeFormat  As String
    
    TimeFormat = Replace(Replace(Replace(Replace(Replace(Replace(Format(TestTime), "3", "s"), "2", "n"), "1", "h"), "0s", "ss"), "0n", "nn"), "0h", "hh")

    FormatSystemTime = TimeFormat
    
End Function

' Obtain the system time separator without API calls.
'
' 2021-01-26. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function FormatSystemTimeSeparator() As String

    Dim Separator   As String
    
    Separator = Format(Time, ":")

    FormatSystemTimeSeparator = Separator
    
End Function

' Returns, for a date value, a formatted string expression with
' year and weeknumber according to ISO-8601.
' Optionally, a W is used as separator between the year and week parts.
'
' Typical usage:
'
'   FormatWeekIso8601(Date)
'   ->  2017-23
'
'   FormatWeekIso8601(Date, True)
'   ->  2017W23
'
' 2017-04-28. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function FormatWeekIso8601( _
    ByVal Expression As Variant, _
    Optional ByVal WSeparator As Boolean) _
    As String
    
    Const Iso8601Separator  As String = "W"
    Const NeutralSeparator  As String = "-"
    
    Dim Result              As String
    
    Dim IsoYear As Integer
    Dim IsoWeek As Integer
    
    If IsDate(Expression) Then
        IsoWeek = Week(DateValue(Expression), IsoYear)
        Result = _
            VBA.Format(IsoYear, String(3, "0")) & _
            IIf(WSeparator, Iso8601Separator, NeutralSeparator) & _
            VBA.Format(IsoWeek, String(2, "0"))
    End If
    
    FormatWeekIso8601 = Result

End Function

' Returns True if parameter Format is a predefined named format.
' Also returns True if an empty Format (zero-length string) is passed.
' Returns False for any other string passed.
'
' 2018-09-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsNamedFormat( _
    ByVal Format As String) _
    As Boolean
    
    Const FullFormatItems   As String = ";General Date;Long Date;Medium Date;Short Date;Long Time;Medium Time;Short Time"
    
    Dim FullFormats         As Variant
    Dim Index               As Integer
    Dim Result              As Boolean
        
    FullFormats = Split(FullFormatItems, ";")
    
    For Index = LBound(FullFormats) To UBound(FullFormats)
        If StrComp(Format, FullFormats(Index), vbTextCompare) = 0 Then
            Result = True
            Exit For
        End If
    Next
    
    IsNamedFormat = Result

End Function

' Returns the "Military Time Code Letter" from and UTC timezone
' for use in the military DTG date/time format, Date Time Group.
' The timezone passed must be as the offset in full hours from UTC
' or as a string of the format [-]h or [-]h:nn.
'
' Accepted values and returned letters:
'
'   -12 -11 -10 -09 -08 -07 -06 -05 -04 -03 -02 -01   0
'     Y   X   W   V   U   T   S   R   Q   P   O   N   Z
'
'    01  02  03  04  05  06  07  08  09  10  11  12
'     A   B   C   D   E   F   G   H   I   K   L   M
'
' Returns the default "Z" for invalid TimezoneOffset values.
'
' 2015-11-25. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function MilitaryTimeCodeLetter( _
    TimezoneOffset As Variant) _
    As String

    Dim Identifier  As Integer
    Dim CodeLetter  As String
    
    ' Ignore empty values.
    If Trim(Nz(TimezoneOffset)) = "" Then Exit Function
    
    ' Extract hours from strings like "-08:00" or "11.5"
    ' while accepting numbers like -7 or 3.
    TimezoneOffset = Fix(Val(Nz(TimezoneOffset)))
       
    Select Case TimezoneOffset
        Case -12 To 12
            Select Case Sgn(TimezoneOffset)
                Case -1
                    Identifier = 13 - TimezoneOffset
                Case 0
                    Identifier = 26
                Case 1
                    ' Skip J.
                    Identifier = TimezoneOffset + TimezoneOffset \ 10
            End Select
        Case Else
            Identifier = 26
    End Select
    CodeLetter = Chr(64 + Identifier)
    
    MilitaryTimeCodeLetter = CodeLetter

End Function

' Returns the UTC timezone offset from a "Military Time Code Letter" for
' use when converting the military DTG date/time format, Date Time Group.

' The timezone offset returned will be in full hours from UTC.
'
' Accepted (not case sensitive) letters and returned values:
'
'     A   B   C   D   E   F   G   H   I   J   K   L   M
'    01  02  03  04  05  06  07  08  09  09  10  11  12
'
'     Y   X   W   V   U   T   S   R   Q   P   O   N   Z
'   -12 -11 -10 -09 -08 -07 -06 -05 -04 -03 -02 -01   0
'
' Returns 0 for invalid letters. J is accepted for letter I.
'
' 2015-11-25. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function MilitaryTimezone( _
    ByVal TimeCodeLetter As String) _
    As Integer

    Dim Letter      As Integer
    Dim TimeOffset  As Integer
    
    Letter = Asc(Trim(UCase(TimeCodeLetter))) - 64

    Select Case Letter
        Case 1 To 13
            ' A to M -> 1 to 12.
            TimeOffset = Letter - Letter \ 10
        Case 14 To 25
            ' N to Y -> -1 to -12
            TimeOffset = 13 - Letter
    End Select
    
    MilitaryTimezone = TimeOffset
    
End Function

' Returns the month number from the English month name.
' Abbreviated names are accepted.
' An ambigous abbreviation (i.e. "ju") will return the first match.
' Passing a non existing name or abbreviation will raise an error.
'
' For parsing localised month names, use function MonthValue.
'
' 2021-04-02. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function MonthFromInvariant( _
    ByVal MonthName As String) _
    As Integer

    Const FirstMonth    As Integer = MinMonthValue
    Const LastMonth     As Integer = MaxMonthValue
    
    Dim Month           As Integer
    
    MonthName = Trim(MonthName)
    If MonthName <> "" Then
        For Month = FirstMonth To LastMonth
            If InStr(1, MonthNameInvariant(Month, True), MonthName, vbTextCompare) = 1 Then
                Exit For
            End If
        Next
    End If
    If Month > LastMonth Then
        ' Month could not be found.
        Err.Raise DtError.dtTypeMismatch
        Exit Function
    End If
    
    MonthFromInvariant = Month

End Function

' Returns the English month name for the passed month number.
' Accepted numbers are 1 to 12. Other values will raise an error.
' If Abbreviate is True, the returned name is abbreviated.
'
' 2015-11-25. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function MonthNameInvariant( _
    ByVal Month As Long, _
    Optional ByVal Abbreviate As Boolean) _
    As String
    
    Const AbbreviatedLength As Integer = 3
    
    Dim MonthName( _
        MinMonthValue To _
        MaxMonthValue)      As String
    Dim Name                As String
    
    If Not IsMonth(Month) Then
        Err.Raise DtError.dtInvalidProcedureCallOrArgument
        Exit Function
    End If
    
    ' Non-localized (invariant) month names.
    MonthName(1) = "January"
    MonthName(2) = "February"
    MonthName(3) = "March"
    MonthName(4) = "April"
    MonthName(5) = "May"
    MonthName(6) = "June"
    MonthName(7) = "July"
    MonthName(8) = "August"
    MonthName(9) = "September"
    MonthName(10) = "October"
    MonthName(11) = "November"
    MonthName(12) = "December"
    
    If Abbreviate = True Then
        Name = Left(MonthName(Month), AbbreviatedLength)
    Else
        Name = MonthName(Month)
    End If
    
    MonthNameInvariant = Name

End Function

    
' Returns the month number from the localised month name.
' Abbreviated names are accepted.
' Passing a non existing name or abbreviation will raise an error.
'
' For parsing invariant (English) month names, use function
' MonthFromInvariant.
'
' 2021-04-02. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function MonthValue( _
    ByVal MonthName As String) _
    As Integer
    
    Const DayText   As String = "1 "
    
    Dim Value       As Integer
    
    Value = Month(DayText & MonthName)
    
    MonthValue = Value
    
End Function

' Parse a text expression for a possible date value.
' Will parse dd/mm or mm/dd as to the local settings.
'
' Examples that can be parsed:
'   "03-08-2020"
'   "03.08.2020"
'   "1026/2006"
'   "O1/19/2007T09:00"
'   "02I21949"
'   "O6/13/1952"
'   "07/27:1956"
'   "07/7:1956"
'   "7/07:1956"
'   "7/7:1956"
'   "06/042/1952"
'
' 2019-09-18. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function ParseDate( _
    ByVal Text As String) _
    As Variant

    ' Length of clean date string.
    Const DefaultLength As Long = 8
    ' Format of string expression for a date.
    Const DefaultFormat As String = "@@/@@/@@@@"
    
    Dim Value           As Variant
    
    ' Replace date typos.
    Text = Replace(Replace(Replace(Text, "O", "0"), "l", "1"), "I", "1")
    ' Replace separator typos.
    Text = Replace(Replace(Replace(Replace(Text, "_", "/"), ".", "/"), "-", "/"), ":", "/")
    
    ' Attempt a convert after basic corrections.
    If IsDate(Text) Then
        Value = DateValue(Text)
    ElseIf Len(Text) = DefaultLength Then
        ' Convert an expression without separators.
        Text = Format(Text, DefaultFormat)
        If IsDate(Text) Then
            Value = DateValue(Text)
        Else
            Value = Null
        End If
    Else
        ' Correct for missing leading zero.
        If Len(Text) - 2 < DefaultLength Then
            Text = "0" & Join(Split(Replace(Text, "/0", "/"), "/"), "/0")
        End If
        
        ' Remove date separator and spaces.
        Text = Replace(Replace(Text, "/", ""), " ", "")
        If Not IsNumeric(Text) Then
            ' Remove trailing text.
            Text = Format(Val(Text), String(DefaultLength, "0"))
        End If
        ' Remove month typos.
        If Len(Text) > DefaultLength Then
            Text = Left(Text, 4) & Right(Text, 4)
        End If
        ' Apply date format.
        Text = Format(Text, DefaultFormat)
        
        ' Convert to date if possible.
        If IsDate(Text) Then
            Value = DateValue(Text)
        Else
            Value = Null
        End If
    End If

    ParseDate = Value

End Function

' Returns the English weekday name for the passed weekday number.
' Mimics exactly WeekdayName() which, however, returns localised names.
' Accepted numbers are 1 to 7. Other values will raise an error, as
' will an invalid value for FirstDayOfWeek.
' If Abbreviate is True, the returned name is abbreviated.
'
' 2020-11-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function WeekdayNameInvariant( _
    ByVal Weekday As Long, _
    Optional ByVal Abbreviate As Boolean, _
    Optional ByVal FirstDayOfWeek As VbDayOfWeek = VbDayOfWeek.vbUseSystemDayOfWeek) _
    As String
    
    Const AbbreviatedLength As Integer = 2
    
    Dim WeekdayNames( _
        FirstWeekday To _
        LastWeekday)        As String
    Dim Name                As String
    
    If Not (IsWeekday(Weekday) And IsWeekday(FirstDayOfWeek)) Then
        Err.Raise DtError.dtInvalidProcedureCallOrArgument
        Exit Function
    End If
    
    ' Non-localized (invariant) weekday names.
    WeekdayNames(VBA.Weekday(1, FirstDayOfWeek)) = "Sunday"
    WeekdayNames(VBA.Weekday(2, FirstDayOfWeek)) = "Monday"
    WeekdayNames(VBA.Weekday(3, FirstDayOfWeek)) = "Tuesday"
    WeekdayNames(VBA.Weekday(4, FirstDayOfWeek)) = "Wednesday"
    WeekdayNames(VBA.Weekday(5, FirstDayOfWeek)) = "Thursday"
    WeekdayNames(VBA.Weekday(6, FirstDayOfWeek)) = "Friday"
    WeekdayNames(VBA.Weekday(7, FirstDayOfWeek)) = "Saturday"
    
    If Abbreviate = True Then
        Name = Left(WeekdayNames(Weekday), AbbreviatedLength)
    Else
        Name = WeekdayNames(Weekday)
    End If
    
    WeekdayNameInvariant = Name

End Function


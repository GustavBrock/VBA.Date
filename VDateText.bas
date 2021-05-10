Attribute VB_Name = "VDateText"
Option Explicit
'
' VDateText
' Version 1.3.4
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
'   DateText
'

' Converts a "Military Date Time Group" (DTG) formatted string
' to a date value.
'
' If IgnoreTimezone is True, the timezone identifier is
' ignored, and the date/time value returned as is.
' If IgnoreTimezone is False, the date/time value is
' converted to UTC.
'
' DTG must be formatted as "ddhhnnZmmmyy".
' If Dtg is not a valid formatted DTG string, Null is returned.
'
' Example:
'   071943ZFEB09 represents 2009-02-07 19:43:00
'
' 2015-11-21. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function CVDateDtg( _
    ByVal Dtg As Variant, _
    Optional ByVal IgnoreTimezone As Boolean) _
    As Variant
    
    Dim DateTime        As Variant
    
    DateTime = Null
    
    On Error Resume Next
    DateTime = CDateDtg(Dtg, IgnoreTimezone)
    On Error GoTo 0
    
    CVDateDtg = DateTime
    
End Function

' Converts an ISO 8601 formatted date/time string
' to a date Expression.
' A timezone info is ignored.
' Optionally, a millisecond part can be ignored.
' Returns Null if Expression in Null or invalid.
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
Public Function CVDateIso8601( _
    ByVal Expression As Variant, _
    Optional ByVal IgnoreMilliseconds As Boolean) _
    As Variant

    Dim DateTime        As Variant
    
    DateTime = Null
    
    On Error Resume Next
    DateTime = CDateIso8601(Expression, IgnoreMilliseconds)
    On Error GoTo 0
    
    CVDateIso8601 = DateTime

End Function

' Parse a text expression for a possible date value.
' Will parse dd/mm or mm/dd as to the local settings.
' Returns Null for invalid expressions like Null.
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
Public Function CVDateText( _
    ByVal Expression As Variant) _
    As Variant

    ' Length of clean date string.
    Const DefaultLength As Long = 8
    ' Format of string expression for a date.
    Const DefaultFormat As String = "@@/@@/@@@@"
    
    Dim Text            As String
    Dim Value           As Variant
    
    Value = Null
    
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

    CVDateText = Value

End Function

' Converts a US formatted date/time string to a date value.
' Returns Null if Expression in Null or invalid.
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
Public Function CVDateUs( _
    ByVal Expression As Variant) _
    As Variant

    Dim DateTime        As Variant
    
    If IsDate(Expression) Then
        DateTime = CDateUs(Expression)
    Else
        DateTime = Null
    End If
    
    CVDateUs = DateTime

End Function

' Returns a date as a "Military Date Time Group" (DTG) formatted string.
' Returns Null for a date value of Null.
' The format is: ddhhnnzmmmyy
'
' Example:
'   2012-01-06 18:30 -01:00 -> "061830NJAN12"
'
' 2015-11-25. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VFormatDateDtg( _
    ByVal Date1 As Variant, _
    Optional TimezoneOffset As Variant = "Z") _
    As Variant
    
    Dim Dtg         As Variant
    
    If IsDateExt(Date1) Then
        Dtg = FormatDateDtg(CDate(Date1), TimezoneOffset)
    Else
        Dtg = Null
    End If
    
    VFormatDateDtg = Dtg

End Function

' Returns the month number from the English month name.
' Abbreviated names are accepted.
' Returns Null if a non existing name or abbreviation is
' passed as MonthName.
'
' For parsing localised month names, use function VMonthValue.
'
' 2015-11-25. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VMonthFromInvariant( _
    ByVal MonthName As Variant) _
    As Variant

    Dim Month   As Variant
    
    Month = Null
    
    MonthName = Trim(Nz(MonthName))
    If MonthName <> "" Then
        ' Ignore the error resulting from a non existing name.
        On Error Resume Next
        Month = MonthFromInvariant(MonthName)
        On Error GoTo 0
    End If

    VMonthFromInvariant = Month

End Function

' Returns the English month name for the passed month number.
' Accepted numbers are 1 to 12. Other values will raise an error.
' If Abbreviate is True, the returned name is abbreviated.
'
' Returns Null for a month value of Null.
'
' 2015-11-25. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VMonthNameInvariant( _
    ByVal Month As Variant, _
    Optional ByVal Abbreviate As Boolean) _
    As Variant
    
    Dim Name                As Variant
    
    If IsNumeric(Month) Then
        Name = MonthNameInvariant(CLng(Month), Abbreviate)
    Else
        Name = Null
    End If
    
    VMonthNameInvariant = Name
    
End Function

' Returns the month number from the localised month name.
' Abbreviated names are accepted.
' Returns Null if a non existing name or abbreviation is
' passed as MonthName.
'
' For parsing invariant (English) month names, use function
' VMonthFromInvariant.
'
' 2021-04-02. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VMonthValue( _
    ByVal MonthName As Variant) _
    As Variant
    
    Dim Month   As Variant
    
    Month = Null
    
    MonthName = Trim(Nz(MonthName))
    If MonthName <> "" Then
        ' Ignore the error resulting from a non existing name.
        On Error Resume Next
        Month = MonthValue(MonthName)
        On Error GoTo 0
    End If
    
    VMonthValue = Month
    
End Function

' Returns the English weekday name for the passed weekday number.
' Accepted numbers are 1 to 7. Other values will raise an error.
' If Abbreviate is True, the returned name is abbreviated.
'
' Returns Null for a weekday value of Null or invalid values for
' FirstDayOfWeek.
'
' 2020-11-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VWeekdayNameInvariant( _
    ByVal Weekday As Variant, _
    Optional ByVal Abbreviate As Boolean, _
    Optional ByVal FirstDayOfWeek As VbDayOfWeek = VbDayOfWeek.vbUseSystemDayOfWeek) _
    As Variant
    
    Dim Name                As Variant
    
    If IsNumeric(Weekday) And IsWeekday(FirstDayOfWeek) Then
        Name = WeekdayNameInvariant(CLng(Weekday), Abbreviate, FirstDayOfWeek)
    Else
        Name = Null
    End If
    
    VWeekdayNameInvariant = Name
    
End Function


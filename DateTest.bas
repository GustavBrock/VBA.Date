Attribute VB_Name = "DateTest"
Option Explicit
'
' DateTest
' Version 1.2.4
'
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.Date
'
' Functions for testing and verification of functions related to date and time.
'
' License: MIT (http://opensource.org/licenses/mit-license.php)
'
' Required references:
'   None
'
' Required modules:
'   DateBank
'   DateBase
'   DateCalc
'   DateCore
'   DateFind
'   DateMsec
'

' Returns a formatted and justified string display of any date value
' including milliseconds and the numeric value of the date value.
'
' Example output:
'    100-01-01 00:00:00.000  -657434
'    100-01-01 00:00:00.001  -657434.000000012
'    100-01-01 23:59:59.999  -657434.999999988
'    999-12-30 00:00:00.000  -328718
'    999-12-30 00:00:00.001  -328718.000000012
'    999-12-30 23:59:59.999  -328718.999999988
'   1899-12-29 00:00:00.000       -1
'   1899-12-29 00:00:00.001       -1.00000001157407
'   1899-12-29 23:59:59.999       -1.99999998842593
'   1899-12-30 00:00:00.000        0
'   1899-12-30 00:00:00.001        1.15740740740741E-08
'   1899-12-30 23:59:59.999         .999999988425926
'   1899-12-31 00:00:00.000        1
'   1899-12-31 23:59:59.498        1.99999418981481
'   1999-01-03 23:59:58.000    36163.9999884259
'   1999-01-03 23:59:59.079    36163.9999893403
'   9999-12-30 23:59:59.009  2958464.99998853
'   9999-12-31 23:59:59.998  2958465.99999998
'   9999-12-31 23:59:59.999  2958465.99999999
'
' 2016-09-17. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DebugDate( _
    ByVal Date1 As Date) _
    As String
    
    ' Adjust tab width if needed. Default is 4.
    Const TabWidth      As Integer = 4
    
    Dim DatePart        As Date
    Dim TimePart        As Date
    Dim Milliseconds    As Integer
    Dim Value           As Double
    Dim ValueText       As String
    Dim Result          As String
    
    ' Create string representation of the numeric value of Date1.
    Value = CDbl(Date1)
    ValueText = Str(Value)
    ValueText = Replace(Replace(ValueText, "-.", "-0."), " .", "0.")
    
    ' Find count of milliseconds.
    Milliseconds = Millisecond(Date1)
    
    ' Round off milliseconds for correct display of seconds.
    RoundOffMilliseconds Date1
    
    ' Get the time part only, to obtain precise seconds.
    ' TimeValue() cannot be used as it is buggy for the date of 9999-12-31.
    DatePart = Fix(Date1)
    TimePart = Date1 - DatePart
    
    ' Build formatted, justified, and fixed length string presentation.
    Result = _
        Right(" " & CStr(Year(DatePart)), 4) & "-" & Right("0" & CStr(Month(DatePart)), 2) & "-" & Right("0" & CStr(Day(DatePart)), 2) & " " & _
        Right("0" & CStr(Hour(TimePart)), 2) & ":" & Right("0" & CStr(Minute(TimePart)), 2) & ":" & Right("0" & CStr(Second(TimePart)), 2) & "." & _
        Right("00" & CStr(Milliseconds), 3) & " " & _
        Space(9 - IIf(InStr(ValueText, ".") = 0, 1 + Len(ValueText), InStr(ValueText, "."))) & ValueText & _
        IIf(Len(ValueText) < 6, Space(4 * TabWidth), "")

    DebugDate = Result
    
End Function

' Lists the field values of all records in a recordset.
' Returns the count of records.
'
' 2017-04-11. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DebugRecordset( _
    ByRef Recordset As DAO.Recordset) _
    As Long

    Dim Field   As DAO.Field
    
    Dim Records As Long

    ' If no records, just exit.
    If Recordset Is Nothing Then Exit Function
    
    ' Records exist.
    If Recordset.RecordCount > 0 Then
        Recordset.MoveFirst
    End If
    While Not Recordset.EOF
        ' Print field values of this record.
        For Each Field In Recordset.Fields
            Debug.Print Field.Value, ;
        Next
        Recordset.MoveNext
        ' Print new line.
        Debug.Print
    Wend
    Records = Recordset.RecordCount
    
    Set Field = Nothing
    
    DebugRecordset = Records
    
End Function

' Append to table TestTimeMsec 1000 records of date BaseDate
' with a millisecond part from 0 to 999.
'
' If DateSelect is zero, BaseDate is used as base date.
' If DateSelect is -1 or 1, lower or upper limit of
' data type Date is used as base date.
'
'
' DAO example:
'
' To fill time table TestTimeMsec with minimum, maximum and zero
' date/time with millisecond values from 0 to 999:
'
'   Call FillTime(-1)
'   Call FillTime(1)
'   Call FillTime(0)
'
' Will insert 3000 records with 1ms resolution.
'
' To insert 1000 records for any other date with 1ms resolution:
'
'   Call FillTime(0, <requested date/time>)
'
' To insert 100 records for any other date with 10ms resolution:
'
'   Call FillTime(0, <requested date/time>, 10)
'
' To insert 10 records for any other date with 0.1s resolution:
'
'   Call FillTime(0, <requested date/time>, 100)
'
'
' 2016-09-11. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub FillTimeMsec( _
    Optional ByVal DateSelect As Integer, _
    Optional ByVal BaseDate As Date, _
    Optional ByVal MsecStep As Integer = 1)

    Const TableMsec   As String = "TestTimeMsec"
    
    Dim Database      As DAO.Database
    Dim Records       As DAO.Recordset
    
    Dim MsecCount     As Integer
    Dim Sql           As String
      
    If BaseDate = ZeroDateValue Then
        ' Use predefined date as base.
        Select Case Sgn(DateSelect)
            Case -1
                BaseDate = MinDateTimeValue
            Case 1
                BaseDate = MaxDateTimeValue
        End Select
    Else
        ' Use specified date/time as base.
    End If
    Sql = "Select * From " & TableMsec
    
    If MsecStep < 1 Then
        MsecStep = 1
    End If
    
    Set Database = CurrentDb
    Set Records = Database.OpenRecordset(Sql)
    With Records
        For MsecCount = MinMillisecondCount To MaxMillisecondCount Step MsecStep
            .AddNew
                .Fields(0).Value = MsecSerial(MsecCount, BaseDate)
            .Update
        Next
        .Close
    End With
    
    Set Records = Nothing
    Set Database = Nothing
  
End Sub

' Print out all dates of a full calendar year starting at
' the specified start date for a financial/fiscal year.
'
' 2020-12-30. Gustav Brock. Cactus Data ApS, CPH.
'
Public Sub TestDateCalendar()

    Dim StartDate   As Date
    Dim ThisDate    As Date
    Dim Index       As Integer
    Dim Offset      As Integer
    
    ' Non-trivial example start date. Adjust as needed.
    StartDate = DateSerial(2020, 9, 20)
    
    ' Set start date of the financial year.
    DateFinancialStart Month(StartDate), Day(StartDate)
    
    ' Print header line.
    Debug.Print "Index", "Financial Date", ;
    For Offset = 1 To 12
        Debug.Print "Month" & Str(Offset), ;
    Next
    Debug.Print
    
    ' Print, for each month, from latest earlier ultimo date, a month plus the subsequent primo date(s).
    For Index = -1 To 31
        ThisDate = DateAdd("d", Index, StartDate)
        Debug.Print Index, ThisDate, , ;
        For Offset = 1 To 12
            Debug.Print DateCalendar(DateAdd("m", Offset - 1, ThisDate)), ;
        Next
        Debug.Print
    Next

End Sub

' Print out all dates of a full financial year starting at
' the specified start date.
'
' 2020-12-30. Gustav Brock. Cactus Data ApS, CPH.
'
Public Sub TestDateFinancial()

    Dim StartDate   As Date
    Dim ThisDate    As Date
    Dim Index       As Integer
    Dim Offset      As Integer
    
    ' Non-trivial example start date. Adjust as needed.
    StartDate = DateSerial(2020, 9, 10)
    
    ' Set start date of the financial year.
    DateFinancialStart Month(StartDate), Day(StartDate)
    
    ' Print header line.
    Debug.Print "Index", "Calendar Date", ;
    For Offset = 1 To 12
        Debug.Print "Month" & Str(Offset), ;
    Next
    Debug.Print
    
    ' Print, for each month, from latest earlier ultimo date, a month plus the subsequent primo date(s).
    For Index = -1 To 31
        ThisDate = DateAdd("d", Index, StartDate)
        Debug.Print Index, ThisDate, , ;
        For Offset = 1 To 12
            Debug.Print DateFinancial(DateAdd("m", Offset - 1, ThisDate)), ;
        Next
        Debug.Print
    Next

End Sub

' Print the five days and the semimonths' numbers around a semimonth shift
' for a sequence of 20 semimonths.
'
' 2019-12-28. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub TestSemimonths(ByVal Semimonths As Integer)

    Const StartDate As Date = #12/1/2018#
    
    Dim Date1           As Date
    Dim Date2           As Date
    Dim Semimonth1      As Integer
    Dim Semimonth2      As Integer

    Dim Period          As Integer
    Dim Offset          As Integer
    Dim MonthPart       As Integer
    Dim Match           As Boolean
    Dim Year            As Integer
    Dim Month           As Integer
    Dim Day             As Integer
    
    ' Print header line.
    Debug.Print "Period", "MonthPart", "Offset", "Date1", "Date2", "Semimonth1", "Semimonth2", "Test OK"
    
    For Period = 0 To 20
        For Offset = -2 To 2
            ' Define start date.
            MonthPart = ((Semimonth(Date1) + Semimonths) Mod SemimonthsPerMonth + SemimonthsPerMonth) Mod SemimonthsPerMonth
            Year = VBA.Year(StartDate)
            Month = VBA.Month(StartDate) + Period \ SemimonthsPerMonth
            Day = (Period Mod SemimonthsPerMonth) * DaysPerSemimonth + VBA.Day(StartDate) + Offset
            Date1 = DateSerial(Year, Month, Day)
            ' Add semimonths to start date to calculate Date2.
            Date2 = DateAddExt(IntervalSetting(dtSemimonth, True), Semimonths, Date1)
            ' Get semimonth of the start date and of the calculated date.
            Semimonth1 = Semimonth(Date1)
            Semimonth2 = Semimonth(Date2)
            ' Check result.
            Match = (Semimonth1 + Semimonths + SemimonthsPerYear) Mod SemimonthsPerMonth = Semimonth2 Mod SemimonthsPerMonth
            ' Report values and result.
            Debug.Print Period, MonthPart, Offset, Format(Date1, "yyyy-mm-dd"), Format(Date2, "yyyy-mm-dd"), Semimonth1, Semimonth2, Match
        Next
    Next

End Sub

' Print the five days and the tertiamonths' numbers around a tertiamonth shift
' for a sequence of 20 tertiamonths.
'
' 2019-12-28. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub TestTertiamonths(ByVal Tertiamonths As Integer)

    Const StartDate As Date = #12/1/2018#
    
    Dim Date1           As Date
    Dim Date2           As Date
    Dim Tertiamonth1    As Integer
    Dim Tertiamonth2    As Integer

    Dim Period          As Integer
    Dim Offset          As Integer
    Dim MonthPart       As Integer
    Dim Match           As Boolean
    Dim Year            As Integer
    Dim Month           As Integer
    Dim Day             As Integer
    
    ' Print header line.
    Debug.Print "Period", "MonthPart", "Offset", "Date1", "Date2", "Tertiamonth1", "Tertiamonth2", "Test OK"
    
    For Period = 0 To 20
        For Offset = -2 To 2
            ' Define start date.
            MonthPart = ((Tertiamonth(Date1) + Tertiamonths) Mod TertiamonthsPerMonth + TertiamonthsPerMonth) Mod TertiamonthsPerMonth
            Year = VBA.Year(StartDate)
            Month = VBA.Month(StartDate) + Period \ TertiamonthsPerMonth
            Day = (Period Mod TertiamonthsPerMonth) * DaysPerTertiamonth + VBA.Day(StartDate) + Offset
            Date1 = DateSerial(Year, Month, Day)
            ' Add Tertiamonths to start date to calculate Date2.
            Date2 = DateAddExt(IntervalSetting(dtTertiamonth, True), Tertiamonths, Date1)
            ' Get tertiamonth of the start date and of the calculated date.
            Tertiamonth1 = Tertiamonth(Date1)
            Tertiamonth2 = Tertiamonth(Date2)
            ' Check result.
            Match = (Tertiamonth1 + Tertiamonths + TertiamonthsPerYear) Mod TertiamonthsPerMonth = Tertiamonth2 Mod TertiamonthsPerMonth
            ' Report values and result.
            Debug.Print Period, MonthPart, Offset, Format(Date1, "yyyy-mm-dd"), Format(Date2, "yyyy-mm-dd"), Tertiamonth1, Tertiamonth2, Match
        Next
    Next

End Sub

' Lists the fortnights around New Year of 2032.
'
' 2016-01-11. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function TheFortnights()

    Dim Date1       As Date
    Dim Date2       As Date
    Dim Number      As Integer
    Dim Loops       As Integer
    
    Loops = 10
    
    ' Forward sequence.
    Date1 = #12/1/2032#
    For Number = 0 To Loops
        Date2 = DateAdd("ww", Number, Date1)
        Debug.Print _
            Date1, _
            Date2, _
            Fortnight(Date1), _
            Fortnight(Date2), _
            Weeks(Date1, Date2), _
            Fortnights(Date1, Date2), _
            DateAddExt("vv", Fortnights(Date1, Date2), Date1)
    Next
    Debug.Print "-"
    
    ' Reverse sequence.
    Date1 = #2/11/2033#
    For Number = 0 To Loops
        Date2 = DateAdd("ww", -Number, Date1)
        Debug.Print _
            Date1, _
            Date2, _
            Fortnight(Date1), _
            Fortnight(Date2), _
            Weeks(Date1, Date2), _
            Fortnights(Date1, Date2), _
            DateAddExt("vv", Fortnights(Date1, Date2), Date1)
    Next
    
End Function

' Counts errors when displaying the seconds of the last day, 9999-12-31, of Date.
'
' TimeValue - and thus Second, DatePart, and Format - will all fail by one second
' for 43136 of the 86400 seconds of the day of 9999-12-31.
' The first part of the function demonstrates this.
'
' To prevent a faulty display, extract seconds from the time part only.
' The second part of the function demonstrates this.
'
' No other day of the entire range of Date will expose this bug.
'
' 2016-01-11. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function TheLastSeconds()

    Dim Hour        As Integer
    Dim Minute      As Integer
    Dim Second      As Integer
    Dim TestDate    As Date
    Dim TestTime    As Date
    
    Dim Total       As Long
    Dim Fail        As Long
    
    Debug.Print "Test with errors:"
    For Hour = 0 To 23
        For Minute = 0 To 59
            For Second = 0 To 59
                Total = Total + 1
                TestTime = TimeSerial(Hour, Minute, Second)
                ' Build date and time value.
                TestDate = #12/31/9999# + TestTime
                ' Get seconds of date and time.
                If VBA.Second(TestDate) <> Second Then
                    Fail = Fail + 1
                End If
            Next
        Next
        Debug.Print Right(Str(Hour), 2), Fail
    Next
    Debug.Print "Total: "; Total & "   Failed: "; Fail
    
    Total = 0
    Fail = 0
    Debug.Print
    
    Debug.Print "Test without errors:"
    For Hour = 0 To 23
        For Minute = 0 To 59
            For Second = 0 To 59
                Total = Total + 1
                TestTime = TimeSerial(Hour, Minute, Second)
                ' Build date and time value.
                TestDate = #12/31/9999# + TestTime
                ' Get seconds of time part only.
                If VBA.Second(TestDate - Fix(TestDate)) <> Second Then
                    Fail = Fail + 1
                End If
            Next
        Next
        Debug.Print Right(Str(Hour), 2), Fail
    Next
    Debug.Print "Total: "; Total & "   Failed: "; Fail

End Function

' Verify function DateWeekdayInWeek by printing 7 x 7 lines like this:
'
'   ThisDate 01-05-2017 ThisWeekday 2 StartWeekday 1 StartOffset -1 StartDate 30-04-2017 EndDate 06-05-2017 TestWeekday 3 TestOffset 2 TestDate 02-05-2017 = Result: 02-05-2017
'
' 2017-05-03. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub TheWeekdaysOfWeek()

    Dim Date1           As Date
    Dim Offset          As Integer
    
    ' Select date to test.
    Date1 = Date
    ' Select offset from test date.
    Offset = -3
    
       
    Dim StartWeekday    As VbDayOfWeek
    Dim TestWeekday     As VbDayOfWeek
    Dim ThisWeekday     As VbDayOfWeek
    
    Dim StartOffset     As Integer
    Dim TestOffset      As Integer
    Dim StartDate       As Date
    Dim TestDate        As Date
    
    Date1 = DateAdd("d", Offset, Date1)
    ThisWeekday = Weekday(Date1)
    
    For StartWeekday = vbSunday To vbSaturday
        For TestWeekday = vbSunday To vbSaturday
            StartOffset = (StartWeekday - ThisWeekday - DaysPerWeek) Mod DaysPerWeek
            StartDate = DateAdd("d", StartOffset, Date1)
            TestOffset = (TestWeekday - StartWeekday + DaysPerWeek) Mod DaysPerWeek
            TestDate = DateAdd("d", TestOffset, StartDate)
            Debug.Print _
                "ThisDate " & Date1 & " " & _
                "ThisWeekday " & ThisWeekday & " " & _
                "StartWeekday " & StartWeekday & " " & _
                "StartOffset " & Str(StartOffset) & " " & _
                "StartDate " & StartDate & " " & _
                "EndDate " & DateAdd("d", 6, StartDate) & " " & _
                "TestWeekday " & TestWeekday & " " & _
                "TestOffset " & TestOffset & " " & _
                "TestDate " & TestDate & " = " & _
                "Result: " & DateWeekdayInWeek(Date1, TestWeekday, StartWeekday)
        Next
        Debug.Print
    Next

End Sub

' Verify for a full year the values returned by function YearFraction for
' either a common or a leap year.
'
' 2020-12-28. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub TheYearFractionDays(ByVal LeapYear As Boolean)

    Const DaysOfCommonYear  As Integer = 365
    
    Dim ThisDate            As Date
    Dim FirstDate           As Date
    Dim ThisDay             As Integer
    Dim LastDay             As Integer
    Dim SomeDay             As Integer
    Dim ThisMonth           As Integer

    LastDay = DaysOfCommonYear + Abs(LeapYear)
    
    ' List all days in the year.
    For ThisDay = 1 To LastDay
        ThisDate = DateSerial(2001 - Abs(LeapYear), 1, ThisDay)
        If ThisMonth <> Month(ThisDate) Then
            ThisMonth = Month(ThisDate)
            Debug.Print
            Debug.Print Right(Str(ThisMonth), 2), ;
        End If
        Debug.Print Right("  " & Str(Day(ThisDate)), 4);
    Next
    Debug.Print
    
    ' List the day count from Jan. 1. of the year to Jan. 1. of the next year
    ' using the rounded output of TotalYears.
    ' This must (and will) match the value of ThisDay.
    FirstDate = DateSerial(2001 - Abs(LeapYear), 1, 1)
    For ThisDay = 1 To LastDay + 1
        ThisDate = DateSerial(2001 - Abs(LeapYear), 1, ThisDay)
        If ThisMonth <> Month(ThisDate) Then
            ThisMonth = Month(ThisDate)
            Debug.Print
            Debug.Print Right(Str(ThisMonth), 2), ;
        End If
        SomeDay = CInt(TotalYears(FirstDate, ThisDate) * LastDay)
        Debug.Print Right("  " & Str(SomeDay), 4);
        If ThisDay <> (SomeDay + 1) Then
            ' Indicate a no-match.
            ' Will not happen.
            Debug.Print "!";
        End If
    Next
    Debug.Print
    Debug.Print
    ThisMonth = 0
    
    ' List the day count from Jan. 1. of the year to Jan. 1. of the next year
    ' using the rounded output of YearFraction.
    ' This must (and will) match the value of ThisDay.
    FirstDate = DateSerial(2001 - Abs(LeapYear), 1, 1)
    For ThisDay = 1 To LastDay
        ThisDate = DateSerial(2001 - Abs(LeapYear), 1, ThisDay)
        If ThisMonth <> Month(ThisDate) Then
            ThisMonth = Month(ThisDate)
            Debug.Print
            Debug.Print Right(Str(ThisMonth), 2), ;
        End If
        SomeDay = CInt(YearFraction(ThisDate) * LastDay)
        Debug.Print Right("  " & Str(SomeDay), 4);
        If ThisDay <> (SomeDay + 0) Then
            ' Indicate a no-match.
            ' Will not happen.
            Debug.Print "!";
        End If
    Next
    Debug.Print
    
End Sub

' List the average count of days per year for one to four years
' around a leapling's day of birth.
'
' Output will be:
'
'   2000-02-29
'   -1
'   2000-02-28    2001-02-28     1  366
'   2000-02-28    2002-02-28     2  365.5
'   2000-02-28    2003-02-28     3  365.33
'   2000-02-28    2004-02-28     4  365.25
'    0
'   2000-02-29    2001-02-28     1  365       2001-03-01     366
'   2000-02-29    2002-02-28     2  365       2002-03-01     365.5
'   2000-02-29    2003-02-28     3  365       2003-03-01     365.33
'   2000-02-29    2004-02-29     4  365.25    2004-02-29     365.25
'    1
'   2000-03-01    2001-03-01     1  365
'   2000-03-01    2002-03-01     2  365
'   2000-03-01    2003-03-01     3  365
'   2000-03-01    2004-03-01     4  365.25
'
' 2017-05-03. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub TheYearsAverageDayCount()

    Const Date1 As Date = #2/29/2000#
    Const Iso   As String = "yyyy-mm-dd"
    
    Dim Date2   As Date
    Dim Date3   As Date
    Dim Years   As Integer
    Dim Offset  As Integer
    
    Debug.Print Format(Date1, Iso)
    For Offset = -1 To 1
        Debug.Print Offset
        For Years = 1 To 4
            Date2 = DateAdd("d", Offset, Date1)
            Date3 = DateSerial(Year(Date1) + Years, 3, 1)
            Debug.Print Format(Date2, Iso), Format(DateAdd("yyyy", Years, Date2), Iso), Str(Years) & " " & Str(Round(DateDiff("d", Date2, DateAdd("yyyy", Years, Date2)) / Years, 2));
            If Offset = 0 Then
                If Day(DateAdd("yyyy", Years, Date2)) <> 29 Then
                    Debug.Print , Format(Date3, Iso), Str(Round(DateDiff("d", Date2, Date3) / Years, 2))
                Else
                    Debug.Print , Format(DateAdd("yyyy", Years, Date2), Iso), Str(Round(DateDiff("d", Date2, DateAdd("yyyy", Years, Date2)) / Years, 2))
                End If
            Else
                Debug.Print
            End If
        Next
    Next
    
End Sub

' Return the current version of VBA.
'
' 2020-02-27. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VbeVersion() As String

    Const Separator As String = "."
    Const Major     As Integer = 0
    Const Minor     As Integer = 1
    
    Dim Parts       As Variant
    Dim Version     As String
    
    Parts = Split(VBE.Version, Separator)
    Version = Parts(Major) & Separator & CStr(Val(Parts(Minor)))

    VbeVersion = Version

End Function


Attribute VB_Name = "DateSets"
Option Explicit
'
' DateSets
' Version 1.1.8
'
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.Date
'
' Functions to generate sequences of date/time values as arrays or recordsets.
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

' Returns a sorted DAO read-only recordset with all dates within
' the specified interval from Date1 inclusive to Date2 exclusive.
' The dates can be any date within the range of Date.
'
' If Date2 is equal to Date1, no dates are found and an
' empty recordset is returned.
'
' If Date2 is earlier than to Date1, the dates are returned in
' reverse sequence.
' Note that the sequence will always contain the start date but
' not the end date. Thus, reversing the sequence will return a
' recordset that differs by one record.
'
' Note that, initially, DatesPeriod.RecordCount will not report
' the actual record count but a count of 1 only.
' MoveLast must be called to obtain the true record count.
'
' 2015-12-15. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DatesPeriod( _
    ByVal Interval As String, _
    ByVal Date1 As Date, _
    ByVal Date2 As Date) _
    As DAO.Recordset
    
    ' Sequence step.
    Const Number    As Double = 1
    
    Dim Records     As DAO.Recordset
    Dim Count       As Double
    
    ' Will fail for invalid parameters.
    Count = DateDiff(Interval, Date1, Date2)
    Set Records = DatesSequence(Interval, Sgn(Count) * Number, Date1, Abs(Count))
    
    Set DatesPeriod = Records

End Function

    
' Returns a sorted array of all dates within the specified
' interval from Date1 inclusive to Date2.
'
' If Date2Excluded is True, Date2 is excluded and the returned
' array will not be Dim'ed if Date1 and Date2 are equal.
' If Date2Excluded is False, Date2 is included and the returned
' array will always have at least one item.
'
' The dates can be any date within the range of Date.
'
' If Date2 is earlier than to Date1, the dates are
' returned in reverse sequence.

' Will create up to about 62 mio. items in 75 seconds
' with a memory consumption of 512 MB.
'
' 2016-04-12. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DatesPeriodArray( _
    ByVal Interval As String, _
    ByVal Date1 As Date, _
    ByVal Date2 As Date, _
    Optional Date2Excluded As Boolean = True) _
    As Date()

    Dim Number  As Long
    Dim Days    As Long
    Dim Dates() As Date
    
    Days = DateDiff(Interval, Date1, Date2)
    If Days <> 0 Or Date2Excluded = False Then
        Days = Days - Abs(Date2Excluded)
        ReDim Dates(0 To Days)
        
        For Number = 0 To Days
            Dates(Number) = DateAdd(Interval, Number, Date1)
        Next
    End If
    
    DatesPeriodArray = Dates
  
End Function

' Returns a DAO read-only recordset with random date/time values
' spaced Interval * Number within the range from Date1 to Date2.
'
' The count of values and the interval must be specified.
' The values for Number and Interval must be valid for DateAdd().
' If invalid parameters are passed, an error will be raised.
' The maximum count of values are empirically limited to 30 mio. as
' larger values may not fit within the maximum database size of 2GB.
'
' A step value (Number) can be specified. The default value is 1.
' If Number is set to, say, 15 and Interval is "s", values of the range
' returned will be spaced by multipla of 15 seconds.
'
' The recordset contains two fields:
'   Id      (Long)
'   Date    (DateTime)
' Id will be from 0 to Count - 1.
'
' Execution time, examples:
'   Count:    1000  .04s
'   Count: 1000000   11s
'
' 2015-12-18. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DatesRandom( _
    ByVal Interval As String, _
    ByVal Number As Double, _
    ByVal Date1 As Date, _
    ByVal Date2 As Date, _
    ByVal Count As Double) _
    As DAO.Recordset
    
    ' Query that generate a date/time sequence from a start date/time.
    Const QueryName     As String = "DatesRandom"
    
    ' Maximum recordcount that can be generated within the
    ' 2 GB limit for the filesize of Access.
    ' Empiric value is about 30.9 mio. records.
    ' Thus, being a bit cautious:
    Const RecordLimit   As Double = 30 * 10 ^ 6
        
    Dim Query           As DAO.QueryDef
    Dim Records         As DAO.Recordset
    
    Dim Test            As Double
    
    ' Raise error if the recordset generated is likely to
    ' exceed the database size of about 2 GB or if
    ' Number * Count will exceed the range of the query
    ' causing the record count to be smaller than Count.
    If Count > RecordLimit Then
        Err.Raise DtError.dtOverflow
        Exit Function
    ElseIf Count <= 0 Then
        ' Return Nothing.
    Else
        ' Validate input parameters.
        ' Will fail for invalid parameters and combinations
        ' that would cause the query to fail.
        Test = DateDiff(Interval, Date1, Date2, vbUseSystemDayOfWeek, vbUseSystem)
        
        ' No errors. Create recordset.
        Set Query = CurrentDb.QueryDefs(QueryName)
        Query.Parameters("Interval").Value = Interval
        Query.Parameters("Number").Value = Number
        Query.Parameters("Date1").Value = Date1
        Query.Parameters("Date2").Value = Date2
        Query.Parameters("Count").Value = Count
        Set Records = Query.OpenRecordset(dbOpenSnapshot, dbReadOnly)
        
        Query.Close
        Set Query = Nothing
    End If
    
    Set DatesRandom = Records
    
End Function

' Returns a sorted DAO read-only recordset with a range of date/time
' values from a start date/time.
'
' The count of values and the interval between these must be specified.
' The values for Number and Interval must be valid for DateAdd() or DateAddExt.
' If invalid parameters are passed, an error will be raised.
' The maximum count of values are empirically limited to 30 mio. as
' larger values may not fit within the maximum database size of 2GB.
'
' A step value (Number) can be specified. The default value is 1.
' If Number is set to, say, 15 and Interval is "s", values of the range
' returned will be spaced by 15 seconds.
'
' Number * Count is limited to the value of MaxAddNumber, 2^31 - 1, or
' slightly above 2 * 10^9.
'
' The recordset contains two fields:
'   Id      (Long)
'   Date    (DateTime)
' Id will be from 0 to Count - 1.
'
' Date will hold from DateStart and the next Count - 1 values. Thus:
'   Interval = "h"
'   Number = 1
'   DateStart = 2020-03-01
'   Count = 24
' will return 24 values from:
'   Id =  0, Date = 2020-03-01 00:00
' to:
'   Id = 23, Date = 2020-03-01 23:00
'
' Execution time, examples:
'   Number: 100 Count:  100000  1.8s
'   Number: 100 Count:    1000  .04s
'   Number:  10 Count: 1000000   17s
'   Number:  10 Count:  400000  4.8s
'   Number:  10 Count:  200000  2.8s
'   Number:  10 Count:  100000  1.8s
'   Number:   1 Count: 1000000   17s
'   Number:   1 Count:  400000  4.8s
'   Number:   1 Count:  200000  2.8s
'   Number:   1 Count:  100000  1.8s
'   Number:   1 Count:   10000  .20s
'   Number:   1 Count:    1000  .04s
'   All seconds of one day      1.0s
'   All hours of one month      .03s
'   All dates of one year       .02s
'   All months of one century   .03s
'   All dates of range of Date   43s
'
' 2017-10-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DatesSequence( _
    ByVal Interval As String, _
    ByVal Number As Double, _
    ByVal DateStart As Date, _
    ByVal Count As Double) _
    As DAO.Recordset
    
    ' QueryName that generates a date/time sequence from a start date/time.
    Const QueryNative   As String = "DatesSequence"
    ' QueryName that generates an extended date/time sequence from a start date/time.
    Const QueryExtended As String = "DatesSequenceExt"
    
    ' Maximum recordcount that can be generated within the
    ' 2 GB limit for the filesize of Access.
    ' Empiric value is about 30.9 mio. records.
    ' Thus, being a bit cautious:
    Const RecordLimit   As Double = 30 * 10 ^ 6
        
    Dim Query           As DAO.QueryDef
    Dim Records         As DAO.Recordset
    
    Dim Test            As Date
    Dim QueryName       As String
    
    ' Raise error if the recordset generated is likely to
    ' exceed the database size of about 2 GB or if
    ' Number * Count will exceed the range of the QueryName
    ' causing the record count to be smaller than Count.
    If Count > RecordLimit Then
        Err.Raise DtError.dtOverflow
        Exit Function
    ElseIf Number * Count > MaxAddNumber Then
        Err.Raise DtError.dtOverflow
        Exit Function
    ElseIf Count <= 0 Then
        ' Return Nothing.
    Else
        ' Validate input parameters.
        ' Will fail for invalid parameters and combinations
        ' that would cause the QueryName to fail.
        If IsIntervalSetting(Interval, False) Then
            Test = DateAdd(Interval, Number, DateStart)
            QueryName = QueryNative
        ElseIf IsIntervalSetting(Interval, True) Then
            If Interval = DtInterval.dtDecimalSecond Then
                ' Convert to milliseconds as parameter Number of
                ' QueryName DatesSequenceExt rounds decimals off.
                Interval = DtInterval.dtMillisecond
                Number = Number * MillisecondsPerSecond
            End If
            Test = DateAddExt(Interval, Number, DateStart)
            QueryName = QueryExtended
        End If
        
        If QueryName <> "" Then
            ' No errors. Create recordset.
            Set Query = CurrentDb.QueryDefs(QueryName)
            Query.Parameters("Interval").Value = Interval
            Query.Parameters("Number").Value = Number
            Query.Parameters("Date").Value = DateStart
            Query.Parameters("Count").Value = Count
            Set Records = Query.OpenRecordset(dbOpenSnapshot, dbReadOnly)
            
            Query.Close
            Set Query = Nothing
        Else
            ' Return Nothing.
        End If
    End If
    
    Set DatesSequence = Records
    
End Function

' Returns a sorted DAO read-only recordset with all dates of
' the specified year.
' The year can be any valid year, including year 9999.
' If no year or an invalid year is specified, the dates of the
' current year are returned.
'
' Even though the source query itself will not fail, an error
' is raised if an invalid year is specified.
'
' 2015-12-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DatesYear( _
    Optional ByVal Year As Integer) _
    As DAO.Recordset
    
    ' Query that generate dates for any valid year.
    Const Query As String = "DatesYear"
    
    Dim qd      As DAO.QueryDef
    Dim rs      As DAO.Recordset
    
    If Year = 0 Then
        ' Use year of current date.
        Year = VBA.Year(Date)
    ElseIf Not IsYear(Year) Then
        ' Don't accept years outside the range of Date.
        Err.Raise DtError.dtOverflow
        Exit Function
    End If
        
    Set qd = CurrentDb.QueryDefs(Query)
    qd.Parameters(0) = Year
    ' Open snapshot with all dates of the year.
    Set rs = qd.OpenRecordset(dbOpenSnapshot, dbReadOnly)
    
    qd.Close
    Set qd = Nothing
    
    ' Return recordset
    Set DatesYear = rs
    
End Function

' Returns a sorted array of all dates within the specified year.
'
' The year can be any valid year, including year 9999.
' If no year or an invalid year is specified, the dates of the
' current year are returned.
'
' 2016-04-12. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DatesYearArray( _
    Optional ByVal Year As Integer) _
    As Date()
    
    Dim Dates   As Variant
    
    If Year = 0 Then
        ' Use year of current date.
        Year = VBA.Year(Date)
    ElseIf Not IsYear(Year) Then
        ' Don't accept years outside the range of Date.
        Err.Raise DtError.dtOverflow
        Exit Function
    End If
    
    Dates = DatesPeriodArray("d", _
        DateSerial(Year, MinMonthValue, MinDayValue), _
        DateSerial(Year, MaxMonthValue, MaxDayValue), _
        False)

    DatesYearArray = Dates
    
End Function


Attribute VB_Name = "DateWork"
Option Explicit
'
' DateWork
' Version 1.2.5
'
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.Date
'
' Functions for calculations on workdays.
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
'
' Required additionally:
'   Table of holidays
'

' Common constants.

    ' Workdays per week.
    Public Const WorkDaysPerWeek    As Long = 5
    ' Average count of holidays per week maximum.
    ' For normal holiday sequences, use 1.
    ' For taking company week-long private holidays into account, use 4 or even 5.
    Public Const HolidaysPerWeek    As Long = 1

' Adds Number of full workdays to Date1 and returns the found date.
' Number can be positive, zero, or negative.
' Optionally, if WorkOnHolidays is True, holidays are counted as workdays.
'
' For excessive parameters that would return dates outside the range
' of Date, either 100-01-01 or 9999-12-31 is returned.
'
' Will add 500 workdays in about 0.01 second.
'
' Requires table Holiday with list of holidays.
'
' 2021-12-09. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function DateAddWorkdays( _
    ByVal Number As Long, _
    ByVal Date1 As Date, _
    Optional ByVal WorkOnHolidays As Boolean) _
    As Date
    
    Const Interval      As String = "d"
    
    Dim Holidays()      As Date

    Dim Days            As Long
    Dim DayDiff         As Long
    Dim MaxDayDiff      As Long
    Dim Sign            As Long
    Dim Date2           As Date
    Dim NextDate        As Date
    Dim DateLimit       As Date
    Dim HolidayId       As Long

    Sign = Sgn(Number)
    NextDate = Date1
    
    If Sign <> 0 Then
        If WorkOnHolidays = True Then
            ' Holidays are workdays.
        Else
            ' Retrieve array with holidays between Date1 and Date1 + MaxDayDiff.
            ' Calculate the maximum calendar days per workweek.
            If (WorkDaysPerWeek - HolidaysPerWeek) > 1 Then
                MaxDayDiff = Number * DaysPerWeek / (WorkDaysPerWeek - HolidaysPerWeek)
            Else
                MaxDayDiff = Number * DaysPerWeek
            End If
            ' Add one week to cover cases where a week contains multiple holidays.
            MaxDayDiff = MaxDayDiff + Sgn(MaxDayDiff) * DaysPerWeek
            
            If Sign > 0 Then
                If DateDiff(Interval, Date1, MaxDateValue) < MaxDayDiff Then
                    MaxDayDiff = DateDiff(Interval, Date1, MaxDateValue)
                End If
            Else
                If DateDiff(Interval, Date1, MinDateValue) > MaxDayDiff Then
                    MaxDayDiff = DateDiff(Interval, Date1, MinDateValue)
                End If
            End If
            Date2 = DateAdd(Interval, MaxDayDiff, Date1)
            ' Retrive array with holidays.
            Holidays = DatesHoliday(Date1, Date2)
        End If
        Do Until Days = Number
            If Sign = 1 Then
                DateLimit = MaxDateValue
            Else
                DateLimit = MinDateValue
            End If
            If DateDiff(Interval, DateAdd(Interval, DayDiff, Date1), DateLimit) = 0 Then
                ' Limit of date range has been reached.
                Exit Do
            End If
            
            DayDiff = DayDiff + Sign
            NextDate = DateAdd(Interval, DayDiff, Date1)
            Select Case Weekday(NextDate)
                Case vbSaturday, vbSunday
                    ' Skip weekend.
                Case Else
                    ' Check for holidays to skip.
                    ' Ignore error when using LBound and UBound on an unassigned array.
                    On Error Resume Next
                    For HolidayId = LBound(Holidays) To UBound(Holidays)
                        If Err.Number > 0 Then
                            ' No holidays between Date1 and Date2.
                        ElseIf DateDiff(Interval, NextDate, Holidays(HolidayId)) = 0 Then
                            ' This NextDate hits a holiday.
                            ' Subtract one day before adding one after the loop.
                            Days = Days - Sign
                            Exit For
                        End If
                    Next
                    On Error GoTo 0
                    Days = Days + Sign
            End Select
        Loop
    End If
    
    DateAddWorkdays = NextDate

End Function

' Returns the count of full workdays between Date1 and Date2.
' The date difference can be positive, zero, or negative.
' Optionally, if WorkOnHolidays is True, holidays are regarded as workdays.
'
' Note that if one date is in a weekend and the other is not, the reverse
' count will differ by one, because the first date never is included in the count:
'
'   Mo  Tu  We  Th  Fr  Sa  Su      Su  Sa  Fr  Th  We  Tu  Mo
'    0   1   2   3   4   4   4       0   0  -1  -2  -3  -4  -5
'
'   Su  Mo  Tu  We  Th  Fr  Sa      Sa  Fr  Th  We  Tu  Mo  Su
'    0   1   2   3   4   5   5       0  -1  -2  -3  -4  -5  -5
'
'   Sa  Su  Mo  Tu  We  Th  Fr      Fr  Th  We  Tu  Mo  Su  Sa
'    0   0   1   2   3   4   5       0  -1  -2  -3  -4  -4  -4
'
'   Fr  Sa  Su  Mo  Tu  We  Th      Th  We  Tu  Mo  Su  Sa  Fr
'    0   0   0   1   2   3   4       0  -1  -2  -3  -3  -3  -4
'
' Execution time for finding working days of three years is about 4 ms.
'
' Requires table Holiday with list of holidays.
'
' 2015-12-19. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function DateDiffWorkdays( _
    ByVal Date1 As Date, _
    ByVal Date2 As Date, _
    Optional ByVal WorkOnHolidays As Boolean) _
    As Long
    
    Dim Holidays()      As Date
    
    Dim Diff            As Long
    Dim Sign            As Long
    Dim NextHoliday     As Long
    Dim LastHoliday     As Long
    
    Sign = Sgn(DateDiff("d", Date1, Date2))
    If Sign <> 0 Then
        If WorkOnHolidays = True Then
            ' Holidays are workdays.
        Else
            ' Retrieve array with holidays between Date1 and Date2.
            Holidays = DatesHoliday(Date1, Date2, False) 'CBool(Sign < 0))
            ' Ignore error when using LBound and UBound on an unassigned array.
            On Error Resume Next
            NextHoliday = LBound(Holidays)
            LastHoliday = UBound(Holidays)
            ' If Err.Number > 0 there are no holidays between Date1 and Date2.
            If Err.Number > 0 Then
                WorkOnHolidays = True
            End If
            On Error GoTo 0
        End If
        
        ' Loop to sum up workdays.
        Do Until DateDiff("d", Date1, Date2) = 0
            Select Case Weekday(Date1)
                Case vbSaturday, vbSunday
                    ' Skip weekend.
                Case Else
                    If WorkOnHolidays = False Then
                        ' Check for holidays to skip.
                        If NextHoliday <= LastHoliday Then
                            ' First, check if NextHoliday hasn't been advanced.
                            If NextHoliday < LastHoliday Then
                                If Sgn(DateDiff("d", Date1, Holidays(NextHoliday))) = -Sign Then
                                    ' Weekend hasn't advanced NextHoliday.
                                    NextHoliday = NextHoliday + 1
                                End If
                            End If
                            ' Then, check if Date1 has reached a holiday.
                            If DateDiff("d", Date1, Holidays(NextHoliday)) = 0 Then
                                ' This Date1 hits a holiday.
                                ' Subtract one day to neutralize the one
                                ' being added at the end of the loop.
                                Diff = Diff - Sign
                                ' Adjust to the next holiday to check.
                                NextHoliday = NextHoliday + 1
                            End If
                        End If
                    End If
                    Diff = Diff + Sign
            End Select
            ' Advance Date1.
            Date1 = DateAdd("d", Sign, Date1)
        Loop
    End If
    
    DateDiffWorkdays = Diff

End Function

' Adds one full workday to Date1 and returns the found date.
' Optionally, if WorkOnHolidays is True, holidays are counted as workdays.
'
' Requires table Holiday with list of holidays.
'
' 2015-12-19. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function DateNextWorkday( _
    ByVal Date1 As Date, _
    Optional ByVal WorkOnHolidays As Boolean) _
    As Date
    
    Const Number        As Long = 1
    
    Dim ResultDate      As Date
    
    ResultDate = DateAddWorkdays(Number, Date1, WorkOnHolidays)
    
    DateNextWorkday = ResultDate

End Function

' Subtracts one full workday to Date1 and returns the found date.
' Optionally, if WorkOnHolidays is True, holidays are counted as workdays.
'
' Requires table Holiday with list of holidays.
'
' 2015-12-19. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function DatePreviousWorkday( _
    ByVal Date1 As Date, _
    Optional ByVal WorkOnHolidays As Boolean) _
    As Date
    
    Const Number        As Long = -1
    
    Dim ResultDate      As Date
    
    ResultDate = DateAddWorkdays(Number, Date1, WorkOnHolidays)
    
    DatePreviousWorkday = ResultDate

End Function

' Returns the holidays between Date1 and Date2.
' The holidays are returned as an array with the
' dates ordered ascending, optionally descending.
'
' The array is declared static to speed up
' repeated calls with identical date parameters.
'
' Requires table Holiday with list of holidays.
'
' 2021-12-09. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DatesHoliday( _
    ByVal Date1 As Date, _
    ByVal Date2 As Date, _
    Optional ByVal OrderDesc As Boolean) _
    As Date()
    
    ' Constants for the arrays.
    Const DimRecordCount    As Long = 2
    Const DimFieldOne       As Long = 0
    
    Static Date1Last        As Date
    Static Date2Last        As Date
    Static OrderLast        As Boolean
    Static DayRows          As Variant
    Static Days             As Long
    
    Dim Records             As DAO.Recordset
    
    ' Cannot be declared Static.
    Dim Holidays()          As Date
    
    If DateDiff("d", Date1, Date1Last) <> 0 Or _
        DateDiff("d", Date2, Date2Last) <> 0 Or _
        OrderDesc <> OrderLast Then
        
        ' Retrieve new range of holidays.
        Set Records = RecordsHoliday(Date1, Date2, OrderDesc)
        
        ' Save the current set of date parameters.
        Date1Last = Date1
        Date2Last = Date2
        OrderLast = OrderDesc
        
        Days = Records.RecordCount
        If Days > 0 Then
            ' Get the full record count.
            Records.MoveLast
            Days = Records.RecordCount
            ' As repeated calls may happen, do a movefirst.
            Records.MoveFirst
            DayRows = Records.GetRows(Days)
            ' Records is now positioned at the last record.
        End If
        Records.Close
    End If
    
    If Days = 0 Then
        ' Leave Holidays() as an unassigned array.
        Erase Holidays
    Else
        ' Fill array to return.
        ReDim Holidays(Days - 1)
        For Days = LBound(DayRows, DimRecordCount) To UBound(DayRows, DimRecordCount)
            Holidays(Days) = DayRows(DimFieldOne, Days)
        Next
    End If
        
    Set Records = Nothing
    
    DatesHoliday = Holidays()
    
End Function

' Returns the first workday of the month of Date1.
' Optionally, if WorkOnHolidays is True, holidays are regarded as workdays.
'
' Requires table Holiday with list of holidays.
'
' 2022-07-09. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateWorkdayMonthFirst( _
    ByVal Date1 As Date, _
    Optional ByVal WorkOnHolidays As Boolean) _
    As Date
    
    Dim ResultDate  As Date

    ResultDate = DateNextWorkday(DatePreviousMonthUltimo(Date1), WorkOnHolidays)
    
    DateWorkdayMonthFirst = ResultDate

End Function

' Returns the last workday of the month of Date1.
' Optionally, if WorkOnHolidays is True, holidays are regarded as workdays.
'
' Requires table Holiday with list of holidays.
'
' 2022-07-09. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateWorkdayMonthLast( _
    ByVal Date1 As Date, _
    Optional ByVal WorkOnHolidays As Boolean) _
    As Date
    
    Dim ResultDate  As Date

    ResultDate = DatePreviousWorkday(DateNextMonthPrimo(Date1), WorkOnHolidays)
    
    DateWorkdayMonthLast = ResultDate

End Function

' Returns the count of holidays between two dates.
'
' Requires table Holiday with list of holidays.
'
' 2021-12-09. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function HolidayCount( _
    ByVal Date1 As Date, _
    ByVal Date2 As Date) _
    As Long
  
    Dim Records     As DAO.Recordset
  
    Dim Holidays    As Long

    Set Records = RecordsHoliday(Date1, Date2)
    If Records.RecordCount > 0 Then
        Records.MoveLast
        Holidays = Records.RecordCount
    End If
    Records.Close
    
    Set Records = Nothing
 
    HolidayCount = Holidays
 
End Function

' Returns True if the passed date is a holiday as recorded in the Holiday table.
'
' Requires table Holiday with list of holidays.
'
' 2021-12-11. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsDateHoliday( _
    ByVal Date1 As Date) _
    As Boolean
    
    Dim IsHoliday   As Boolean
    
    IsHoliday = CBool(HolidayCount(Date1, Date1))
    
    IsDateHoliday = IsHoliday

End Function

' Returns True if the passed date is not a holiday as recorded in the Holiday table
' or a weekend day ("off day") as specified by parameter WeekendType.
'
' Default check is for the days of a long (Western) weekend, Saturday and Sunday.
' Requires table Holiday with list of holidays.
'
' 2021-12-11. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsDateWorkday( _
    ByVal Date1 As Date, _
    Optional ByVal WeekendType As DtWeekendType = DtWeekendType.dtLongWeekend) _
    As Boolean
    
    Dim IsWorkday   As Boolean
    
    If Not IsDateWeekend(Date1, WeekendType) Then
        IsWorkday = Not IsDateHoliday(Date1)
    End If

    IsDateWorkday = IsWorkday

End Function

' Returns the holidays between Date1 and Date2 as a recordset.
' The holidays are returned as a recordset with the
' dates ordered ascending, optionally descending.
'
' Requires table Holiday with list of holidays.
'
' 2021-12-11. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function RecordsHoliday( _
    ByVal Date1 As Date, _
    ByVal Date2 As Date, _
    Optional ByVal ReverseOrder As Boolean) _
    As DAO.Recordset
        
    ' The table that holds the holidays.
    Const Table         As String = "Holiday"
    ' The field of the table that holds the dates of the holidays.
    Const Field         As String = "Date"
    
    Dim Records         As DAO.Recordset
    
    Dim FirstDate       As Date
    Dim LastDate        As Date
    Dim ReverseValues   As Boolean
    Dim Sql             As String
    Dim SqlFirstDate    As String
    Dim SqlLastDate     As String
    Dim Order           As String
    
    ReverseValues = (DateDiff("d", Date1, Date2) < 0)
    If ReverseValues Then
        FirstDate = Date2
        LastDate = Date1
    Else
        FirstDate = Date1
        LastDate = Date2
    End If
    
    SqlFirstDate = Format(FirstDate, "\#yyyy\/mm\/dd\#")
    SqlLastDate = Format(LastDate, "\#yyyy\/mm\/dd\#")
    ReverseOrder = ReverseOrder Xor ReverseValues
    Order = IIf(ReverseOrder, "Desc", "Asc")

    Sql = "Select " & Field & " From " & Table & " " & _
        "Where " & Field & " Between " & SqlFirstDate & " And " & SqlLastDate & " " & _
        "Order By 1 " & Order

    Set Records = CurrentDb.OpenRecordset(Sql, dbOpenSnapshot)
        
    Set RecordsHoliday = Records
    
End Function

' Find the workday number counting from the first Monday of the month.
' Return Null for non-working days.
'
' 2023-11-01. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function WorkDayMonth( _
    ByVal Date1 As Date) _
    As Variant
  
    Dim DateFirst       As Date
    Dim DayNumber       As Variant
    
    If Weekday(Date1, vbMonday) > 5 Then
        DayNumber = Null
    Else
        DateFirst = DateWeekdayInMonthFirst(Date1, vbMonday)
        If DateFirst > Date1 Then
            ' Date1 belongs to the previous month.
            DateFirst = DateWeekdayInMonthFirst(DateAdd("m", -1, Date1), vbMonday)
        End If
        DayNumber = 1 + DateDiffWorkdays(DateFirst, Date1)
    End If
  
    WorkDayMonth = DayNumber
  
End Function

' Returns the count of workdays of the month of Date1.
' Optionally, if WorkOnHolidays is True, holidays are regarded as workdays.
'
' Requires table Holiday with list of holidays.
'
' 2019-02-03. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function WorkdaysInMonth( _
    ByVal Date1 As Date, _
    Optional ByVal WorkOnHolidays As Boolean) _
    As Long
    
    Dim DateFirst   As Date
    Dim DateLast    As Date
    Dim Workdays    As Long
    
    DateFirst = DateSerial(VBA.Year(Date1), VBA.Month(Date1), MinDayValue)
    DateLast = DateAdd(IntervalSetting(DtInterval.dtMonth), 1, DateFirst)
    
    Workdays = DateDiffWorkdays(DateFirst, DateLast, WorkOnHolidays)
    
    WorkdaysInMonth = Workdays
    
End Function

' Returns the count of workdays of the months between the
' month of Date1 and the month of Date2.
' Optionally, if WorkOnHolidays is True, holidays are regarded as workdays.
'
' Requires table Holiday with list of holidays.
'
' 2019-02-03. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function WorkdaysInMonths( _
    ByVal Date1 As Date, _
    ByVal Date2 As Date, _
    Optional ByVal WorkOnHolidays As Boolean) _
    As Long
    
    Dim Months      As Long
    Dim DateFirst   As Date
    Dim DateLast    As Date
    Dim Workdays    As Long
    
    Months = DateDiff(IntervalSetting(DtInterval.dtMonth), Date1, Date2)
    DateFirst = DateSerial(VBA.Year(Date1), VBA.Month(Date1), MinDayValue)
    DateLast = DateAdd(IntervalSetting(DtInterval.dtMonth), Months + 1, DateFirst)
    
    Workdays = DateDiffWorkdays(DateFirst, DateLast, WorkOnHolidays)
    
    WorkdaysInMonths = Workdays
    
End Function


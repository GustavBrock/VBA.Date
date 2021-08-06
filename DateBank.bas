Attribute VB_Name = "DateBank"
Option Explicit
'
' DateBank
' Version 1.3.0
'
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.Date
'
' Functions for handling financial years.
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


' A fiscal year is designated as the calendar year in which it ends.
' For example, if the fiscal year runs from June 1, 2022, to May 31, 2023,
' it could be designated:
'
'   financial year 2023
'   fiscal year 2023
'   FY2023
'
' Initially, call function DateFinancialStart or DateFinancialEnd to define
' the financial/fiscal year.


' Constants.

    ' Default start day and month of the financial/fiscal year applied a
    ' neutral year for storing the values as a Date value.
    Private Const DefaultStart  As Date = #1/1/2000#
    
' Statics.

    ' Local static variable to hold the selected start day and month of the financial/fiscal year.
    ' These are defined and read by the functions:
    '
    '   DateFinancialStart
    '   DateFinancialEnd
    '
    Private FinancialStart      As Date
'

' Returns a financial/fiscal (pseudo) date as its closest equivalent date of the calendar year.
' If argument FinancialDate contains a time part, this will be preserved.
'
' Example:
'   Function TestDateCalendar will print out all calendar dates from a full financial year.
'
' 2020-12-30. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function DateCalendar( _
    ByVal FinancialDate As Date) _
    As Date
    
    Dim StartMonth      As Integer
    Dim StartDay        As Integer
    Dim CalendarDate    As Date
    Dim MonthOffset     As Integer
    
    ' Obtain start of financial year.
    StartMonth = Month(DateFinancialStart)
    StartDay = Day(DateFinancialStart)
    
    ' Offset the financial date to a month having 31 days, January.
    MonthOffset = MinMonthValue - Month(FinancialDate)
    FinancialDate = DateAdd("m", MonthOffset, FinancialDate)
    
    ' Align calendar date to the first day.
    CalendarDate = DateAdd("d", MinDayValue - StartDay, FinancialDate)
    ' Align calendar date to the first month.
    CalendarDate = DateAdd("m", MinMonthValue - StartMonth - MonthOffset, CalendarDate)
    
    DateCalendar = CalendarDate

End Function

' Returns a calendar date as its equivalent pseudo date of the financial/fiscal year.
' Thus, the financial date for the calendar date of the start of the financial year
' is returned as January 1st.
' If argument CalendarDate contains a time part, this will be preserved.
'
' This is mostly useful for creating search criteria, or to obtain the quarter, the
' month, or the year of the financial/fiscal year for a calendar date.
'
' Example:
'   Function TestDateFinancial will print out all dates of a full financial year.
'
' 2021-05-09. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function DateFinancial( _
    ByVal CalendarDate As Date) _
    As Date
    
    Dim StartMonth      As Integer
    Dim StartDay        As Integer
    Dim FinancialDate   As Date
    Dim MonthOffset     As Integer
    
    ' Obtain start of financial year.
    StartMonth = Month(DateFinancialStart)
    StartDay = Day(DateFinancialStart)
    If StartMonth = MinMonthValue And StartDay = MinDayValue Then
        ' The financial year is the calendar year.
        FinancialDate = CalendarDate
    Else
        ' Offset the calendar date to a month having 31 days, January.
        MonthOffset = MinMonthValue - Month(CalendarDate)
        CalendarDate = DateAdd("m", MonthOffset, CalendarDate)
        
        ' Align financial date to the start day.
        FinancialDate = DateAdd("d", MaxDayValue - StartDay + MinDayValue, CalendarDate)
        ' Align financial date to the start month.
        FinancialDate = DateAdd("m", MaxMonthValue - StartMonth - MonthOffset, FinancialDate)
    End If
    
    DateFinancial = FinancialDate
      
End Function

' Gets or sets - based on the end day and month - the end day and month of the
' financial/fiscal year as a date value applied a neutral year.
'
' The end month can be any month.
' The end day can be any day larger than 1. However, if end day is larger than 28,
' which is the highest day value valid for any month, the start day is set to 1.
'
' Default value is December 31st.
'
' Examples:
'   ' Set financial year.
'   EndDate = DateFinancialEnd(9, 30)
'   ' EndDate -> 2000-09-30
'
'   ' Get financial year.
'   StartDate = DateFinancialStart
'   ' StartDate -> 2000-10-01
'   EndDate = DateFinancialEnd
'   ' EndDate -> 2000-09-30
'
' 2021-05-08. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function DateFinancialEnd( _
    Optional ByVal EndMonth As Integer, _
    Optional ByVal EndDay As Integer) _
    As Date
    
    Dim StartMonth  As Integer
    Dim StartDay    As Integer
    Dim EndDate     As Date
    
    ' Validate input.
    If IsMonth(EndMonth) And IsDay(EndDay) Then
        ' Set FinancialStart based on the end date.
        If EndDay < MaxDayAllMonthsValue - 1 Then
            ' Set start date of the financial year as the day after the end day.
            StartMonth = EndMonth
            StartDay = EndDay + 1
        Else
            ' Set start date of the financial year as primo next month.
            StartMonth = EndMonth Mod MonthsPerYear + 1
            StartDay = 1
        End If
        DateFinancialStart StartMonth, StartDay
    End If
        
    ' Return the end date of the financial year.
    EndDate = DateAdd("d", -1, DateFinancialStart)
    
    DateFinancialEnd = EndDate
        
End Function

' Gets or sets the start day and month of the financial/fiscal year as a
' date value applied a neutral year.
'
' The start month can be any month.
' The start day can be any day less than or equal 28, which is the
' highest day value valid for any month.
'
' Default value is January 1st.
'
' Examples:
'   ' Set financial year.
'   StartDate = DateFinancialStart(10, 1)
'   ' StartDate -> 2000-10-01
'
'   ' Get financial year.
'   StartDate = DateFinancialStart
'   ' StartDate -> 2000-10-01
'   EndDate = DateFinancialEnd
'   ' EndDate -> 2000-09-30
'
' 2021-05-08. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function DateFinancialStart( _
    Optional ByVal StartMonth As Integer, _
    Optional ByVal StartDay As Integer) _
    As Date
    
    ' Validate input.
    If IsMonth(StartMonth) And IsDayAllMonths(StartDay) Then
        FinancialStart = DateSerial(Year(DefaultStart), StartMonth, StartDay)
    End If
    If FinancialStart = #12:00:00 AM# Then
        FinancialStart = DefaultStart
    End If
    
    DateFinancialStart = FinancialStart

End Function

' Returns the primo calendar date of the specified financial/fiscal year.
' Returns the primo calender date of the current financial year, if no
' financial year is specified.
'
' 2021-05-09. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function DateFinancialYearPrimo( _
    Optional ByVal FinancialYear As Integer) _
    As Date
    
    Dim Month       As Integer
    Dim Day         As Integer
    Dim Years       As Integer
    Dim Primo       As Date
    
    Month = VBA.Month(DateFinancialStart())
    Day = VBA.Day(DateFinancialStart())
    
    If IsYear(FinancialYear) Then
        Years = FinancialYear - VBA.Year(FinancialStart)
        If Month = MinMonthValue And Day = MinDayValue Then
            ' The financial year is the calendar year.
        Else
            Years = Years - 1
        End If
    Else
        Years = VBA.Year(DateCalendar(Date)) - VBA.Year(FinancialStart)
    End If
    
    Primo = DateAdd("yyyy", Years, FinancialStart)

    DateFinancialYearPrimo = Primo
    
End Function

' Returns the ultimo calendar date of the specified financial/fiscal year.
' Returns the ultimo calender date of the current financial year, if no
' financial year is specified.
'
' 2021-05-09. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function DateFinancialYearUltimo( _
    Optional ByVal FinancialYear As Integer) _
    As Date
    
    Dim Ultimo      As Date
        
    Ultimo = DateAdd("d", -1, DateAdd("yyyy", 1, DateFinancialYearPrimo(FinancialYear)))
    
    DateFinancialYearUltimo = Ultimo

End Function

' Returns the financial/fiscal month of the passed calendar date.
'
' 2019-10-11. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function MonthFinancial( _
    ByVal CalendarDate As Date) _
    As Integer
    
    Dim Value       As Integer
    
    Value = Month(DateFinancial(CalendarDate))
    
    MonthFinancial = Value

End Function

' Returns the financial/fiscal quarter of the passed calendar date.
'
' 2019-10-11. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function QuarterFinancial( _
    ByVal CalendarDate As Date) _
    As Integer
    
    Dim Value       As Integer
    
    Value = Quarter(DateFinancial(CalendarDate))
    
    QuarterFinancial = Value

End Function

' Returns the financial/fiscal year of the passed calendar date.
'
' 2019-10-11. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function YearFinancial( _
    ByVal CalendarDate As Date) _
    As Integer
    
    Dim Value       As Integer
    
    Value = Year(DateFinancial(CalendarDate))
    
    YearFinancial = Value

End Function


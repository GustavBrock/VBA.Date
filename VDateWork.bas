Attribute VB_Name = "VDateWork"
Option Explicit
'
' DateWork
' Version 1.2.3
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
'   DateCore
'   DateFind
'   DateMsec
'   DateWork
'
' Required additionally:
'   Table of holidays
'

' Adds Number of full workdays to Date1 and returns the found date.
' Number can be positive, zero, or negative.
' Optionally, if WorkOnHolidays is True, holidays are counted as workdays.
' Returns Null if any parameter is invalid.
'
' For excessive parameters that would return dates outside the range
' of Date, either 100-01-01 or 9999-12-31 is returned.
'
' Will add 500 workdays in about 0.01 second.
'
' Requires table Holiday with list of holidays.
'
' 2015-12-19. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function VDateAddWorkdays( _
    ByVal Number As Variant, _
    ByVal Date1 As Variant, _
    Optional ByVal WorkOnHolidays As Boolean) _
    As Variant
    
    Dim ResultDate      As Variant
    
    ResultDate = Null
    
    If IsDateExt(Date1) Then
        If IsNumeric(Number) Then
            On Error Resume Next
            ResultDate = DateAddWorkdays(CDbl(Number), CDate(Date1), WorkOnHolidays)
            On Error GoTo 0
        End If
    End If
    
    VDateAddWorkdays = ResultDate
    
End Function

' Returns the count of full workdays between Date1 and Date2.
' The date difference can be positive, zero, or negative.
' Optionally, if WorkOnHolidays is True, holidays are regarded as workdays.
' Returns Null if any parameter is invalid.
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
Public Function VDateDiffWorkdays( _
    ByVal Date1 As Variant, _
    ByVal Date2 As Variant, _
    Optional ByVal WorkOnHolidays As Boolean) _
    As Variant

    Dim Result          As Variant
    
    If IsDateExt(Date1) And IsDateExt(Date2) Then
        Result = DateDiffWorkdays(CDate(Date1), CDate(Date2), WorkOnHolidays)
    Else
        Result = Null
    End If
    
    VDateDiffWorkdays = Result

End Function

' Adds one full workday to Date1 and returns the found date.
' Optionally, if WorkOnHolidays is True, holidays are counted as workdays.
' Returns Null if any parameter is invalid.
'
' Requires table Holiday with list of holidays.
'
' 2015-12-19. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function VDateNextWorkday( _
    ByVal Date1 As Variant, _
    Optional ByVal WorkOnHolidays As Boolean) _
    As Variant
    
    Dim ResultDate      As Variant
    
    If IsDateExt(Date1) Then
        ResultDate = DateNextWorkday(CDate(Date1), WorkOnHolidays)
    Else
        ResultDate = Null
    End If
    
    VDateNextWorkday = ResultDate

End Function

' Subtracts one full workday to Date1 and returns the found date.
' Optionally, if WorkOnHolidays is True, holidays are counted as workdays.
' Returns Null if any parameter is invalid.
'
' Requires table Holiday with list of holidays.
'
' 2015-12-19. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function VDatePreviousWorkday( _
    ByVal Date1 As Variant, _
    Optional ByVal WorkOnHolidays As Boolean) _
    As Variant
    
    Dim ResultDate      As Variant
    
    If IsDateExt(Date1) Then
        ResultDate = DatePreviousWorkday(CDate(Date1), WorkOnHolidays)
    Else
        ResultDate = Null
    End If
    
    VDatePreviousWorkday = ResultDate

End Function

' Returns True if the passed date is a holiday as recorded in the Holiday table.
' Returns Null if any parameter is invalid.
'
' Requires table Holiday with list of holidays.
'
' 2021-12-11. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VIsDateHoliday( _
    ByVal Date1 As Date) _
    As Boolean
    
    Dim Result      As Variant
    
    If IsDateExt(Date1) Then
        Result = IsDateHoliday(CDate(Date1))
    Else
        Result = Null
    End If
    
    VIsDateHoliday = Result

End Function

' Returns True if the passed date is a holiday as recorded in the Holiday table or
' a weekend day ("off day") as specified by parameter WeekendType.
' Returns Null if any parameter is invalid.
'
' Default check is for the days of a long (Western) weekend, Saturday and Sunday.
' Requires table Holiday with list of holidays.
'
' 2021-12-11. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VIsDateWorkday( _
    ByVal Date1 As Date, _
    Optional ByVal WeekendType As DtWeekendType = DtWeekendType.dtLongWeekend) _
    As Boolean
    
    Dim Result      As Variant
    
    If IsDateExt(Date1) Then
        Result = IsDateWorkday(CDate(Date1))
    Else
        Result = Null
    End If

    VIsDateWorkday = Result

End Function


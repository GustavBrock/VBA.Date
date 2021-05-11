# VBA.Date

## A vast collection of functions for all sorts of handling Date and Time in *Microsoft Access* and *Microsoft Excel*

*(c) Gustav Brock, Cactus Data ApS, CPH*

![Help](https://raw.githubusercontent.com/GustavBrock/VBA.Date/master/images/EE%20Title.png)

## Introduction
This is a comprehensive collection of more than 500 functions for dealing with Date and Time in every imaginable way.

The aim is, that every function meets the highest standards, thus no quick-n-dirty methods will be found.

Should you find a bug or a missing feature, please report these under Discussions.

### Goals
For any function, where relevant, the goals to meet have been:

1. Shall handle any date value within the entire range of data type Date with millisecond precision.
2. Shall perform error handling like native functions of VBA, meaning a relevant error code is returned (no messageboxes).
3. Shall not require third-party tools or libraries.
4. Shall use consistent naming convention for function names and variables without Hungarian notation, weird names and abbreviations.
5. Shall make extensive use of constants and enums to avoid smart numbers and unreadable code.
6. Shall reuse other functions to avoid duplicated code.
7. Shall have in-line documentation for all non-trivial code blocks.

### Organisation
The majority (about 300) of the functions are kept in these ten modules:

- DateBank.bas
- DateBase.bas
- DateCall.bas
- DateCore.bas
- DateFind.bas
- DateMsec.bas
- DateSets.bas
- DateSpan.bas
- DateText.bas
- DateWork.bas

This is mainly to keep related functions together, though a stringent separation is not possible. For example, all functions directly related to milliseconds are held in module DateMsec, but functions aware of milliseconds can be found in other modules as well, for example in the modules _DateCore_ and _DateSpan_.

### Functions for queries
The remaining functions are siblings to the main functions intended for usage in queries and in VBA where arguments and variables can be _Null_. These are all prefixed with a **V** and held in modules also having a **V** prefix, i.e. _VDateCore_.
A special module is _VDateTime_ which holds **V** siblings to all the native date functions of VBA.

These function will _never fail_, only return _Null_ in case one or more arguments are _Null_ or invalid. This simplifies a lot of the cases where you otherwise would have to prevent errors by using _Nz_ or filtering for _Null_ values.

### Intervals
In addition to the native date and time intervals - year, month, hour, etc. - all functions (where relevant) can handle a set of predefined custom intervals. These are held in enum DtInterval:

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

As for weeks, a full set of functions for dealing with weeknumbers according to the **ISO 8601** standard is included, for example the function _Week_:

	' Returns the ISO 8601 week of a date.
	' The related ISO year is returned by ref.
	'
	' 2016-01-06. Gustav Brock, Cactus Data ApS, CPH.
	'
	Public Function Week( _
	    ByVal Date1 As Date, _
	    Optional ByRef IsoYear As Integer) _
	    As Integer
	
	    Dim Month       As Integer
	    Dim Interval    As String
	    Dim Result      As Integer
	    
	    Interval = IntervalSetting(dtWeek)
	    
	    Month = VBA.Month(Date1)
	    ' Initially, set the ISO year to the calendar year.
	    IsoYear = VBA.Year(Date1)
	    
	    Result = DatePart(Interval, Date1, vbMonday, vbFirstFourDays)
	    If Result = MaxWeekValue Then
	        If DatePart(Interval, DateAdd(Interval, 1, Date1), vbMonday, vbFirstFourDays) = MinWeekValue Then
	            ' OK. The next week is the first week of the following year.
	        Else
	            ' This is really the first week of the next ISO year.
	            ' Correct for DatePart bug.
	            Result = MinWeekValue
	        End If
	    End If
	        
	    ' Adjust year where week number belongs to next or previous year.
	    If Month = MinMonthValue Then
	        If Result >= MaxWeekValue - 1 Then
	            ' This is an early date of January belonging to the last week of the previous ISO year.
	            IsoYear = IsoYear - 1
	        End If
	    ElseIf Month = MaxMonthValue Then
	        If Result = MinWeekValue Then
	            ' This is a late date of December belonging to the first week of the next ISO year.
	            IsoYear = IsoYear + 1
	        End If
	    End If
	    
	    ' IsoYear is returned by reference.
	    Week = Result
	        
	End Function


### Datasets - sets of dates
One module, _DateSets_, is devoted generating recordsets of all sorts of date values and intervals. An example is function *DatesSequence*:

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

## Documentation

Top level documentation generated by [MZ-Tools](https://www.mztools.com/) is included for [Microsoft Access and Excel](https://htmlpreview.github.io?https://github.com/GustavBrock/VBA.Date/blob/master/documentation/Date.htm).

Detailed documentation is in-line. 

<hr>

*If you wish to support my work or need extended support or advice, feel free to:*

<p>

[<img src="https://raw.githubusercontent.com/GustavBrock/VBA.Date/master/images/BuyMeACoffee.png">](https://www.buymeacoffee.com/gustav/)

Attribute VB_Name = "Date9597"
Option Explicit
'
' Date9597
' Version 1.0.0
'
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.Date
'
' Functions to replace the native VBA.DateTime functions in
' Access 95 and 97.
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

' Direct replacement for DateAdd for Access 95/97 only.
'
' In Access 95/97, DateAdd() is buggy as it will return invalid
' date values between the numeric data values -1 and 0.
' With Access 2000 and newer, DateAdd() can be used as is.
'
' 2006-02-15. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateAdd9x( _
    ByVal Interval As String, _
    ByVal Number As Double, _
    ByVal Date1 As Date) _
    As Date

    ' Version major of Access 97.
    Const VersionMajorMax   As Byte = 8
  
    ' Store current version of Access.
    Static VersionMajor     As Byte
  
    Dim Result              As Date
    Dim Factor              As Long
    Dim Milliseconds        As Double
  
    If VersionMajor = 0 Then
        ' Read and store the current version of Access.
        VersionMajor = Val(SysCmd(acSysCmdAccessVer))
    End If
  
    If VersionMajor > VersionMajorMax Then
        ' Use DateAdd() as is.
        Result = DateAdd(Interval, Number, Date1)
    Else
        Select Case Interval
            Case "h"
                Factor = HoursPerDay
            Case "n"
                Factor = MinutesPerDay
            Case "s"
                Factor = SecondsPerDay
        End Select
        If Factor > 0 Then
            Milliseconds = MillisecondsPerDay * Number / Factor
            Result = MsecSerial(Milliseconds, Date1)
        Else
            Result = DateAdd(Interval, Number, Date1)
        End If
    End If
    
    DateAdd9x = Result
  
End Function


Attribute VB_Name = "Seconds_"
'
'###############################################################################
'# Visual Basic 6 source file "Seconds.bas"
'#
'# Timing with microsecond accuracy (more or less).
'#
'# John Trenholme - started 4 Jan 2004.
'###############################################################################

' Usage:
'
'  Dim t As Double
'  t = seconds()
'  '... operation(s) to be timed
'  t = seconds() - t  ' t now has seconds taken by operation(s)

Option Explicit

Private Const Version_c As String = "2007-01-23"

Private Declare Function QueryPerformanceFrequency _
  Lib "kernel32" (f As Currency) As Boolean
Private Declare Function QueryPerformanceCounter _
  Lib "kernel32" (p As Currency) As Boolean

'===============================================================================
' Return number of seconds since first call to this routine. A return of -86400
' indicates an error.
' The granularity will be around a microsecond.
Public Function seconds() As Double
Attribute seconds.VB_Description = "Return number of seconds since first call to this routine. Good to a few microseconds. A return of -86400 indicates an error."
Static s_base As Currency  ' initializes to 0
Static s_freq As Currency  ' initializes to 0
Const c_Default As Double = -86400#  ' 1 day in seconds, negated
If s_freq = 0@ Then  ' routine not initialized, or unable to read frequency
  QueryPerformanceFrequency s_freq  ' try to read frequency
  ' if frequency is good, try to read base time (else it stays at 0)
  If s_freq <> 0@ Then QueryPerformanceCounter s_base
End If
' if we have a good base time, then we must have a good frequency also
If s_base <> 0@ Then
  Dim time As Currency
  QueryPerformanceCounter time
  If time <> 0@ Then
    seconds = (time - s_base) / s_freq
  Else
    seconds = c_Default
  End If
Else  ' something is wrong - return error value
  seconds = c_Default
End If
End Function

'===============================================================================
' Return a string with a number of successive-call intervals in microseconds.
Public Function secondsIntervals() As String
Dim str As String
str = "seconds() intervals: "
Dim t1 As Double
t1 = seconds()  ' warm up
Dim t2 As Double
t2 = seconds()  ' warm up some more
Dim j As Integer
For j = 1 To 8  ' allow space for up to 8 * 99.999 microsecs in 80 characters
  t1 = seconds()
  t2 = seconds()
  str = str & Format$(1000000# * (t2 - t1), "0.000") & " "
Next j
' the Chr$(181) value below is Greek mu
secondsIntervals = str & Chr$(181) & "s"
End Function

'===============================================================================
' Return an approximation to the uncertainty when seconds() is used, in seconds.
Public Function secondsJitter() As Double
Attribute secondsJitter.VB_Description = "Return an approximation to the uncertainty in the ""seconds()"" routine, in seconds."
Dim t1 As Double
Dim t2 As Double
Const c_N As Long = 15&  ' number of samples - must be > 2, & preferably odd
Dim t(1 To c_N) As Double
Dim j As Long
' prime the pump (get code in cache, etc.)
t1 = seconds()
' collect samples
For j = 1 To c_N
  t2 = seconds()
  Do
    t1 = t2
    t2 = seconds()
  Loop While t1 = t2  ' wait for clock to tick
  t(j) = t2 - t1
Next j
' sort the sample values (insertion sort)
Dim k As Long
For j = 2 To c_N
  t1 = t(j)
  k = j - 1&
  Do While t(k) > t1
    t(k + 1&) = t(k)
    k = k - 1&
    If k < 1& Then Exit Do
  Loop
  t(k + 1&) = t1
Next j
' return median value, to avoid unusually low or high values
secondsJitter = t(1& + c_N \ 2&)
End Function

'===============================================================================
' Return a string with the date of the latest revision to this file
Public Function secondsVersion() As String
secondsVersion = Version_c
End Function

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


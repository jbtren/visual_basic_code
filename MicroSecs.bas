Attribute VB_Name = "MicroSecs_"
'
'###############################################################################
'# Visual Basic 6 source file "MicroSecs.bas"
'#
'# Microsecond timing (more or less).
'#
'# John Trenholme - started 4 Jan 2004. This version 17 Mar 2004.
'###############################################################################

Option Explicit

Private Declare Function QueryPerformanceFrequency _
  Lib "kernel32" (f As Currency) As Boolean
Private Declare Function QueryPerformanceCounter _
  Lib "kernel32" (p As Currency) As Boolean

'===============================================================================
' Returns time since first call to this routine, in microseconds.
Public Function microSecs() As Currency
Attribute microSecs.VB_Description = "Returns number of microseconds since first call to this routine, as Currency for many digits of result."
Static s_base As Currency  ' initializes to 0
Static s_freq As Currency  ' initializes to 0
If s_freq = 0@ Then  ' routine not initialized, or unable to read frequency
  QueryPerformanceFrequency s_freq
  ' if frequency is good, try to read base time (else it stays at 0)
  If s_freq <> 0@ Then QueryPerformanceCounter s_base
End If
' if we have a good base time, then we must have a good frequency also
If s_base <> 0@ Then
  Dim time As Currency
  QueryPerformanceCounter time
  microSecs = Int(1000000@ * (time - s_base) / s_freq)
Else  ' something is wrong - return error value
  microSecs = -1@
End If
End Function

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


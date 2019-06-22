Attribute VB_Name = "TimeStampMod"
'###############################################################################
'#
'# Visual Basic module "TimeStampMod.bas" ' avoids name conflct with function
'# Time-stamp with accurate centiseconds by John Trenholme - started 2014-07-18
'#
'###############################################################################

Option Base 0          ' array base value, when not specified - the default
Option Compare Binary  ' string comparison based on Asc(char) - the default
Option Explicit        ' forces explicit variable declaration - changes default

Private Const Version_c As String = "2019-06-21"

'===============================================================================
Public Function TimeStamp() As String
' Standard file-name-compatible date-time stamp routine for VBA
' You get back a string something like "2016-03-26_15-22-41.37"
' Use with (e.g.) "OutputFile_" & TimeStamp() & ".txt"
' Warning - you only get even seconds on the Macintosh
Dim nw As Double, hrs As Single, min As Single, sec As Single
Do: nw = Now(): sec = Timer(): Loop Until Int(nw) = Int(Now()) ' same day
hrs = Int(sec / 3600!): sec = sec - 3600! * hrs: min = Int(sec / 60!)
sec = Round(sec - 60! * min, 2)
' change the following if you want a different date-part or time-part separator
' note that the 'normal' time separator (":") is not legal in a file name
Const DS As String = "-", TS As String = "-"  ' date-part & time-part separators
' force numeric values (except 4-digit year & ms) to always have 2 digits
TimeStamp = CStr(Year(nw)) & DS & Format$(Month(nw), "00") & DS & _
  Format$(Day(nw), "00") & "_" & Format$(hrs, "00") & TS & _
  Format$(min, "00") & TS & Format$(sec, "00.00")  ' cs is about best we can do
End Function

'===============================================================================
Public Sub TimerInc()
Dim tmr1 As Double, tmr2 As Double, k As Long
Debug.Print "=== Timer() steps in milliseconds ==="
For k = 1 To 7
  Do: tmr1 = Timer(): tmr2 = Timer(): Loop Until tmr1 < tmr2
  Debug.Print 1000# * (tmr2 - tmr1); " ";
Next k
Debug.Print
End Sub

'===============================================================================
Private Sub ClearImmediate()
' Clear the Immediate window from the VBA editor
Application.SendKeys "^g^a~" ' leaves blank line at top - BackSpace removes it
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

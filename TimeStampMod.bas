Attribute VB_Name = "TimeStampMod"
'###############################################################################
'#
'# Visual Basic module "TimeStamp.bas"
'# Time-stamp with accurate centiseconds by John Trenholme - started 2014-07-18
'#
'###############################################################################

Option Base 0          ' array base value, when not specified - the default
Option Compare Binary  ' string comparison based on Asc(char) - the default
Option Explicit        ' forces explicit variable declaration - changes default

Private Const Version_c As String = "2019-05-24"

'===============================================================================
Public Function TimeStamp() As String
' Standard file-name-compatible date-time stamp routine for VBA
' You get back a string something like "2016-03-26_15-22-41.37"
' Use with (e.g.) "OutputFile_" & TimeStamp() & ".txt"
' Warning - you only get even seconds on the Macintosh
Dim nw, hrs As Single, min As Single, sec As Single
nw = Now(): sec = Timer(): hrs = Int(sec / 3600!)
sec = sec - 3600! * hrs: min = Int(sec / 60!): sec = Round(sec - 60! * min, 2)
' change the following if you want a different date-part or time-part separator
' note that the 'normal' time separator (":") is not legal in a file name
Const DS As String = "-", TS As String = "-"  ' date-part & time-part separators
' force numeric values (except 4-digit year & ms) to always have 2 digits
TimeStamp = CStr(Year(nw)) & DS & Format$(Month(nw), "00") & DS & _
  Format$(Day(nw), "00") & "_" & Format$(hrs, "00") & TS & _
  Format$(min, "00") & TS & Format$(sec, "00.00")  ' cs is about best we can do
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

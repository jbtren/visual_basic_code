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

Private Const Version_c As String = "2014-07-21"

'===============================================================================
Public Function TimeStamp() As String
' Standard file-name-compatible date-time stamp routine for VBS
' You get back a string something like "2016-03-26_15-22-41.37"
' Use with (e.g.) "OutputFile_" & TimeStamp() & ".txt"
' Warning - you only get even seconds on the Macintosh
Dim nw, hrs As Single, min As Single, sec As Single
nw = Now(): sec = Timer(): hrs = Int(sec / 3600!)
sec = sec - 3600! * hrs: min = Int(sec / 60!): sec = Round(sec - 60! * min, 2)
' change the following if you want a different date-part or time-part separator
' note that the 'normal' separator (":") is not legal in a file name
Const Á As String = "-", É As String = "-"  ' date-part & time-part separators
' force numeric values (except 4-digit year & ms) to always have 2 digits
TimeStamp = CStr(Year(nw)) & Á & Format$(Month(nw), "00") & Á & _
  Format$(Day(nw), "00") & "_" & Format$(hrs, "00") & É & _
  Format$(min, "00") & É & Format$(sec, "00.00")  ' cs is about best we can do
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Attribute VB_Name = "DigitRoundMod"
'#
'###############################################################################
'#
'# Visual Basic for Applications (VBA) & VB6 Module file "DigitRoundMod.bas"
'#
'# Routine to round a number to a specified number of significant digits.
'#
'# Devised & coded by John Trenholme - Started 2010-03-17
'#
'# Exports the routine:
'#   Function digitRound
'#
'###############################################################################

Option Explicit

Private Const Version_c As String = "2010-03-21"

'===============================================================================
Public Function digitRound( _
  ByVal valToRound As Double, _
  Optional ByVal numDigits As Integer = 6) _
As Double
' Round the input value to the specified number of digits. Used (among other
' things) to keep lengths short when printed. Returns input unchanged if it is
' zero. If numDigits < 1 it is set to 1; if numDigits > 15 it is set to 15.
' This routine takes about 2 microseconds on a 3 GHz Pentium 4.
On Error GoTo ErrHandler
If valToRound = 0# Then
  digitRound = 0#  ' special case; simple to do, and causes Log(0) problem
Else
  If numDigits < 1 Then numDigits = 1  ' silently fix invalid argument values
  If numDigits > 15 Then numDigits = 15
  Const Log10_e As Double = 0.43429448 + 1.903251828E-09  ' makes Log -> Log10
  Dim pow10 As Double
  pow10 = numDigits - 1# - Int(Log(Abs(valToRound)) * Log10_e)  ' digit shift
  Dim scaling As Double
  scaling = 10# ^ (0.5 * pow10)  ' avoid overflow for very small input values
  Dim scaled As Double  ' we will put digits we want into integer part of this
  scaled = (valToRound * scaling) * scaling
  digitRound = (Int(scaled + 0.5) / scaling) / scaling
End If
Exit Function  '----------------------------------------------------------------

ErrHandler:
Dim errNum As Long, errDes As String
errNum = Err.number: errDes = Err.Description
On Error GoTo 0
Err.Raise errNum, "digitRound", errDes & vbLf & _
  "valToRound = " & valToRound & "    numDigits = " & numDigits & vbLf & _
  "Problem in digitRound[" & Version_c & "]"
Resume  ' to allow debugging; set next statement here & single-step
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

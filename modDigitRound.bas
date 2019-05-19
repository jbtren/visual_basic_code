Attribute VB_Name = "modDigitRound"
Attribute VB_Description = "Module holding function to round its input value to a specified number of digits."
'
'###############################################################################
'#
'# Visual Basic 6 file "modDigitRound.bas"
'#
'# Devised and coded by John Trenholme - Initial version 14 Aug 2003
'#
'###############################################################################

Option Base 0
Option Compare Binary
Option Explicit

Private Const c_fileName As String = "modDigitRound"  ' ID for this file
Private Const c_version As String = "2003-08-18"  ' version (date) of this file

'*******************************************************************************
Public Function digitRound(ByVal valToRound As Double, _
                           ByVal nDigits As Integer) _
As Double
Attribute digitRound.VB_Description = "Round the input value to the specified number of digits. Used (among other things) to keep lengths short when printed. Returns input unchanged if it is zero, or if nDigits < 1 or nDigits > 15."
' ------------------------------------------------------------------------------
' Round the input value to the specified number of digits. Used (among other
' things) to keep lengths short when printed. Returns input unchanged if it is
' zero, or if nDigits < 1 or nDigits > 15.
' ------------------------------------------------------------------------------
If (nDigits > 0) And (valToRound <> 0#) Then  ' input makes sense - proceed
  If nDigits > 15 Then nDigits = 15  ' Double can't print more than 15 in VB
  digitRoundDbl = Format$(valToRound, "0." & String$(nDigits - 1, "0") & "E-0")
Else  ' nDigits zero or negative, or input value = 0.0, so silently do nothing
  digitRoundDbl = valToRound
End If
End Function

'-------------------------------- end of file ----------------------------------


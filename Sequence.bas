Attribute VB_Name = "Sequence_"
'
'###############################################################################
'# Visual Basic 6 source file "Sequence.bas"
'#
'# Initial version 19 Mar 2004 by John Trenholme. This version 19 Mar 2004.
'###############################################################################

Option Explicit

'===============================================================================
' Supply a unique sequence number on each call (1, 2, 3, ...).
' Note: this will become -2147483648 after 2147483647 calls, and roll over after
' 4294967296 calls. But why are you getting that many unique numbers anyway?
Public Function sequence() As Long
Static s_value As Long
If s_value < 2147483647 Then s_value = s_value + 1& Else s_value = -2147483648#
sequence = s_value
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


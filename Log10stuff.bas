Attribute VB_Name = "Log10stuff"
'
'###############################################################################
'#
'# Visual Basic 6 source file "Log10stuff.bas"
'#
'# John Trenholme - initial version 21 Aug 2003
'#
'###############################################################################

Option Base 0
Option Compare Binary
Option Explicit
Option Private Module  ' no effect in Visual Basic; globals project-only in VBA

Public Const Log10stuffVersion As String = "2009-10-29"

' this value is accurate to the last bit
' so Log10(1En) gives an exact integer for -308 <= n <= 308
Public Const Log10e As Double = 0.43429448 + 1.903251828E-09

'===============================================================================
Public Function Log10(ByVal x As Double) As Double
Static calls_s As Double  ' number of times this routine has been called
calls_s = calls_s + 1#    ' stops adding at 9,007,199,254,740,992 calls
On Error GoTo ErrorHandler

Log10 = Log(x) * Log10e
Exit Function '*********** routine has just one exit (for debugging); this is it

ErrorHandler:  '----------------------------------------------------------------
Dim errNum As Long, errDsc As String  ' Err object Property holders
errNum = Err.Number
errDsc = Err.Description
On Error GoTo 0  ' avoid recursion
Const ID As String = "Log10stuff.Log10"
errDsc = "Error in " & ID & " call " & calls_s & vbLf & _
  "argument = " & x & vbLf & errDsc
Err.Raise errNum, ID, errDsc ' send error on up the call chain
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


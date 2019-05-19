Attribute VB_Name = "IntegrateArray_"
'
'###############################################################################
'# Visual Basic source file "IntegrateArray.bas"
'#
'# Initial version 2010-02-04 by John Trenholme. This version 2010-10-27.
'###############################################################################

Option Explicit

'===============================================================================
Public Function integrateArray(ByRef arg() As Double) As Double
' Returns an approximation to the integral of the data in the input array,
' which contains F(X) values at equally-spaced values of X. See Numerical
' Recipes section 4.1 (formula 4.1.14) for the method used. The integral is
' from the X of the first F(X) point to the X of the last, assuming the point
' spacing is unity. If the spacing is "h" multiply the result by that quantity.
' This formula integrates functions up to cubic exactly. It requires at least
' seven points to make sense.
Const ID_c As String = "integrateArray"  ' name of this routine
Static calls_s As Double  ' number of times this routine has been called
calls_s = calls_s + 1#    ' stops adding at 9,007,199,254,740,992 calls

Dim j1 As Long, j2 As Long
j1 = LBound(arg)
j2 = UBound(arg)
If j2 - j1 < 6& Then
  Err.Raise 5&, ID_c, "Too few points. Need 7 or more but got " & j2 - j1 + 1& & vbLf & _
    "Problem in " & ID_c & " call " & calls_s
End If
Const A1 As Double = 3# / 8#  ' see the Numerical Recipes section
Const A2 As Double = 7# / 6#
Const A3 As Double = 23# / 24#
Dim ret As Double
ret = A1 * (arg(j1) + arg(j2)) + A2 * (arg(j1 + 1&) + arg(j2 - 1&)) + _
  A3 * (arg(j1 + 2&) + arg(j2 - 2&))
Dim j As Long
For j = j1 + 3& To j2 - 3&
  ret = ret + arg(j)
Next j
integrateArray = ret
End Function  '-----------------------------------------------------------------

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


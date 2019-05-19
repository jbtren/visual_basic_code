Attribute VB_Name = "FudgeFunctions"
Attribute VB_Description = "An assortment of useful functions for smooth or simple approximation to discontinuous or difficult functions. Devised and coded by John Trenholme."
'
'###############################################################################
'#
'# Visual Basic 6 and VBA source file "FudgeFunctions.bas"
'#
'# Designed and coded by John Trenholme
'# Begun 5 Jan 2006
'#
'###############################################################################

Option Base 0
Option Compare Binary
Option Explicit

Public Const FudgeFunctionsVersion As String = "2006-01-25"

Private Const c_Pi As Double = 3.1415926 + 5.35897932E-08

'===============================================================================
Function smoothInt(ByVal x As Double, _
                   Optional ByVal width As Double = 0.001) _
As Double
Attribute smoothInt.VB_Description = "Smoothed, continuous approximation to the Int function, which equals zero at X = 0, and steps upward by unity at x = N+1/2 (N integer). The ""width"" parameter sets the X distance over which the smoothed unit jump extends."
' Smoothed, continuous approximation to the Int function, which equals zero at
' X = 0, and steps upward by unity at X = N + 1/2 (N integer). The "width"
' parameter sets the X distance over which the smoothed unit jump extends.

' impose sanity on width of transition
Const c_del As Double = 0.000000000001
If width < c_del Then width = c_del Else If width > 1# Then width = 1#

' We get in trouble when Atn() and CLng() both try to jump at the same point,
' since errors in Pi et. al. will make the jumps at slightly different points.
' Therefore we make two versions with the jump points offset by 1/2, and switch
' between them along 45 degree diagonals through the bottom and top of the jumps
' at X = N + 1/2 (N integer). Thus we never cross a jump point.
Dim dx As Double
dx = x - 0.5 - CLng(x - 0.5)  ' distance from point where CLng() jumps
' the dx limit below gives equal distance to the two forms from X = 0 & X = 1/2
' it is a fast 2%-error approx. to the exact value Atn(Sqr(width)) / Pi
If Abs(dx) > 0.31831 * Sqr(width) * (1# - 0.229268 * width) Then
  ' this approximation form is centered at X = N (N integer)
  smoothInt = Atn(width * Tan(x * c_Pi)) / c_Pi + CLng(x)
Else
  ' this approximation form is centered at X = N + 1/2 (N integer)
  smoothInt = 0.5 + Atn(Tan(c_Pi * (x - 0.5)) / width) / c_Pi + CLng(x - 0.5)
End If
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~


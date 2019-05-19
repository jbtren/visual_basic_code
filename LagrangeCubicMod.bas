Attribute VB_Name = "LagrangeCubicMod"
Option Explicit

'===============================================================================
Public Function LagrangeCubic( _
  ByVal atX As Double, _
  ByRef Xarray As Range, _
  ByRef Yarray As Range, _
  Optional ByVal allowExtrap As Boolean = False) _
As Double
' Performs Lagrange cubic interpolation at the location 'atX', using ranges of
' known X and Y values to get the coefficients of the cubic polynomial.
' The Xarray range must be in strictly ascending order and contain at least four
' elements. Requests for values in the end intervals (or beyond) use the cubic
' through the end four points. For Lagrange cubics, values match at nodes, but
' slopes have jump changes in general. Of course, cubics are done exactly.

' Set the optional argument True to allow extrapolation beyond the end values.

' note: if errors are raised within a Function, you get a #VALUE! error in Excel
If (Xarray.Count < 4&) Or (Xarray.Count <> Yarray.Count) Then Err.Raise 5&
With Application
  If Not allowExtrap Then
    If (atX < .Min(Xarray)) Or (atX > .Max(Xarray)) Then Err.Raise 5&
  End If

  Dim ndx As Long  ' index of X value at or below atX in Xarray
  If atX <= Xarray.Item(2&) Then  ' Match would fail below first element
    ndx = 2&  ' use lowest useful value
  ElseIf atX >= Xarray.Item(Xarray.Count - 2&) Then
    ndx = Xarray.Count - 2&  ' use highest useful value
  Else
    ndx = .Match(atX, Xarray, 1)  ' find place in arrays
    ' keep values used in cubic calculation within array bounds
    If ndx > Xarray.Count - 2& Then ndx = Xarray.Count - 2&
  End If
End With

' make up the Lagrange interpolant, doing the whole calculation each time
Dim j As Long, k As Long
Dim ret As Double, term As Double, xj As Double
For j = ndx - 1& To ndx + 2&
  term = 1#
  xj = Xarray.Item(j)
  For k = ndx - 1& To ndx + 2&
    If j <> k Then  ' get next factor in Lagrange formula term
      term = term * (atX - Xarray.Item(k)) / (xj - Xarray.Item(k))
    End If
  Next k
  ret = ret + term * Yarray.Item(j)  ' term complete; add to result
Next j
LagrangeCubic = ret
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

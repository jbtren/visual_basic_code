Attribute VB_Name = "BrentMin"
Attribute VB_Description = "Minimization of a univariate function. Based on the method of Brent (see 'fmin' in NetLib). Coded by John Trenholme."
'
'###############################################################################
'
' Brent-method univariate function minimization routines.
'
' Visual Basic (VB6 or VBA) module "BrentMin.bas"
'
' Coded by John Trenholme
'
'###############################################################################

Option Base 0
Option Compare Binary
Option Explicit
' Option Private Module  ' no effect in VB6; visible-this-Project-only in VBA

Public Const BrentMinVersion As String = "2005-12-13"

'*******************************************************************************
Public Function BrentMin( _
  ByVal ax As Double, _
  ByVal bx As Double, _
  Optional ByVal functionIndex As Long = 0&, _
  Optional ByVal xRelErr As Double = 0#, _
  Optional ByVal nMax As Long = 100&) As Double

Dim a As Double
Dim b As Double
Dim c As Double
Dim d As Double
Dim e As Double
Dim fu As Double
Dim fv As Double
Dim fw As Double
Dim fx As Double
Dim n As Integer
Dim p As Double
Dim parab As Boolean
Dim q As Double
Dim r As Double
Dim tol As Double
Dim tol1 As Double
Dim tol2 As Double
Dim u As Double
Dim v As Double
Dim w As Double
Dim x As Double
Dim xm As Double
  
a = ax
b = bx
If (b < a) Then
  c = a
  a = b
  b = c
End If
c = (3# - Sqr(5#)) / 2#
x = a + c * (b - a)
fx = BrentMinFunc(x, functionIndex)
n = 1&
e = 0#
v = x
fv = fx
w = x
fw = fx
If (b <= a) Then nMax = 0&
Do While n < nMax
  xm = a + 0.5 * (b - a)
  tol2 = tol * (Abs(a) + Abs(b))
  tol1 = 0.5 * tol2
  If Abs(x - xm) <= (tol2 - 0.5 * (b - a)) Then
    nMax = 0&
  Else
    If tol1 < Abs(e) Then
      r = (x - w) * (fx - fv)
      q = (x - v) * (fx - fw)
      p = (x - v) * q - (x - w) * r
      q = 2# * (q - r)
      If (0# <= q) Then p = -p Else q = -q
      parab = Abs(p) < Abs(0.5 * q * e) And p > q * (a - x) And p < q * (b - x)
    Else
      parab = False
    End If
    If parab Then
      e = d
      d = p / q
      u = x + d
      If u - a < tol2 Then If x < xm Then d = tol1 Else d = -tol1
      If b - u < tol2 Then If x < xm Then d = tol1 Else d = -tol1
    Else
      If xm <= x Then e = a - x Else e = b - x
      d = c * e
    End If
    If Abs(d) < tol1 Then If d >= 0# Then d = tol1 Else d = -tol1
    u = x + d
    fu = BrentMinFunc(u, functionIndex)
    n = n + 1&
    If fu = fx And fu = fw Then
      nMax = 0&
    Else
      If fu <= fx Then
        If x <= u Then a = x Else b = x
        v = w
        fv = fw
        w = x
        fw = fx
        x = u
        fx = fu
      Else
        If u < x Then a = u Else b = u
        If fu <= fw Or w = x Then
          v = w
          fv = fw
          w = u
          fw = fu
        ElseIf fu <= fv Or v = x Or v = w Then
          v = u
          fv = fu
        End If
      End If
    End If
  End If
Loop
BrentMin = x
End Function

'*******************************************************************************
Public Function BrentMinFunc( _
  ByVal x As Double, _
  Optional ByVal functionIndex As Long) As Double
' Put code here for the function(s) you want to minimize.
Select Case functionIndex
  Case 0&
    BrentMinFunc = 2# - x * (x - 2#)
  Case Else
    BrentMinFunc = 0#
End Select
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

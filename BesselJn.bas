Attribute VB_Name = "BesselJn"
'
'###############################################################################
'#
'# Visual Basic 6 source file "BesselJn.bas"
'#
'# Bessel function of arbitrary order.
'#
'###############################################################################

Option Explicit

Function jSubNu#(x#, gnu#)  '===================================================
'-- From Atlas of Functions, section 53:8 - relative error < .0001 if Jn <> 0
'-- WARNING!  Gives garbage in transition region for J > 65
'-- Call with negative x ONLY if gnu is integer
'-- This QuickBASIC 4.5 version copyright 15 May 1992 by John Trenholme
Const Pi# = 3.14159265358979, TwoPi# = 2# * Pi, HalfPi# = 0.5 * Pi
Dim absGnu#, ak#, akMHf#, angl#, coef#, gn#, gn2#, sum#, t#, term#, x2#, xx#
Dim living%
xx = Abs(x)    '-- localize x; silently enforce x >= 0.0 for non-integer gnu
'-- avoid problems for x near zero
If xx < 0.000000000001 Then xx = 0.000000000001
x2 = xx * xx                        '-- common subexpression
absGnu = Abs(gnu)
'-- use power series expansion below transition point
If xx <= absGnu + 0.5 * absGnu ^ (1# / 3#) + 9.5 * Exp(-0.2 * absGnu) - 4! Then
  gn = gnu                            '-- localize gnu, since we change it
  term = 0.5
  '-- test if gnu is a negative integer; if so do special Gamma function
  If gn + Int(Abs(gn)) = 0# Then     '-- note zero enters but does nothing
    gn = -gn
    term = 0.5 * (-1#) ^ gn
  End If
  sum = term                         '-- sum up the series
  ak = 0#
  Do
    ak = ak + 1#
    sum = sum + term
    term = -term * x2 / (4# * ak * (ak + gn))
    sum = sum + term
  Loop While (20000000# * Abs(term) > Abs(sum)) Or (gn <= -ak)
  If gn <> 0# Then                    '-- need to do Gamma calculation
    Do                                 '-- find (x/2)^gnu / Gamma(1+gnu)
      gn = gn + 1#
      sum = 2# * sum * gn / xx
    Loop While gn <= 3#
    gn2 = gn * gn
    t = (1# + (1# / (1.5 * gn2) - 1#) / (3.5 * gn2)) / (30# * gn * gn2)
    t = (t - 1#) / (12# * gn) + gn * (1# + Log(0.5 * xx / gn))
    sum = sum * Exp(t) / Sqr(TwoPi * gn)
  End If
Else                             '-- above transition; use sum-of-cosines form
  akMHf = 0.5                         '-- (k - 1/2) as a float
  coef = 1# / Sqr(HalfPi * xx)
  angl = HalfPi * (gnu + 0.5) - xx
  sum = 0#
  gn2 = gnu * gnu                     '-- common subexpression
  living% = True
  Do
    sum = sum + coef * Cos(angl)
    coef = coef * (akMHf * akMHf - gn2) / (xx * (2# * akMHf + 1#))
    angl = angl + HalfPi
    akMHf = akMHf + 1#
    '-- if asymptotic series is diverging, bail out
    If akMHf * akMHf >= xx * (2# * akMHf + 1#) + gn2 Then living% = False
    If akMHf >= absGnu Then           '-- make sure terms are oscillating
      If 20000000# * Abs(coef) < Abs(sum) Then living% = False
    End If
  Loop While living%
End If
'-- fixup for negative argument
If x < 0# Then
  If gnu = Int(gnu) Then
    sum = (-1#) ^ gnu * sum  ' gnu is an integer; it's OK
  Else
    sum = 0#  ' can't do this case, so silently return zero
  End If
End If
jSubNu = sum
End Function

'--------------------------------- end of file ---------------------------------


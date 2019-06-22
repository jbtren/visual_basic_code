Attribute VB_Name = "Elliptics"
'
'        8888888888 888 888 d8b          888    d8b
'        888        888 888 Y8P          888    Y8P
'        888        888 888              888
'        8888888    888 888 888 88888b.  888888 888  .d8888b .d8888b
'        888        888 888 888 888 "88b 888    888 d88P"    88K
'        888        888 888 888 888  888 888    888 888      "Y8888b.
'        888        888 888 888 888 d88P Y88b.  888 Y88b.         X88
'        8888888888 888 888 888 88888P"   "Y888 888  "Y8888P  88888P'
'                               888
'                               888
'                               888
'  by
'         __       __           ______                 __         __.
'    __  / /___   / /   ___    /_  __/____ ___  ___   / /  ___   / /__ _  ___.
'   / /_/ // _ \ / _ \ / _ \    / /  / __// -_)/ _ \ / _ \/ _ \ / //  ' \/ -_)
'   \____/ \___//_//_//_//_/   /_/  /_/   \__//_//_//_//_/\___//_//_/_/_/\__/
'
'###############################################################################
'#
'# Visual Basic for Applications (VBA) Module file "Elliptics.bas"
'#
'# Routines for calculation of elliptic integrals and Jacobi elliptic functions.
'# Based on code by Cody (complete elliptics) and Fukushima (elliptic functions).
'#
'# Otherwise devised and coded by John Trenholme - started 2014-06-04
'#
'# Exports the routines:
'#   Function completeElliptic1
'#   Function completeElliptic1_comp
'#   Function completeElliptic1_max
'#   Sub ellipticsUnitTest
'#   Function EllipticsVersion
'#   Function jacobiCn
'#   Function jacobiCycleOvr4
'#   Sub jacobiCnDnSn
'#   Function jacobiDn
'#   Function jacobiSn
'#
'###############################################################################

Option Base 0          ' array base value, when not specified - the default
Option Compare Binary  ' string comparison based on Asc(char) - the default
Option Explicit        ' forces explicit variable declaration - changes default
'Option Private Module  ' No visibility outside this VBA Project

Private Const Version_c As String = "2014-07-23"
Private Const File_c As String = "Elliptics.bas[" & Version_c & "]"

Private Const Invalid_c As Long = 5&  ' "Invalid procedure call or argument"
' the smallest number that changes 1.0 when subtracted from it = 2 ^ (-53)
Private Const MachineEpsilonM_c As Double = 1.11022302462516E-16
' the smallest number that changes 1.0 when added to it = 2 ^ (-52)
Private Const MachineEpsilonP_c As Double = 2.22044604925031E-16
' the very smallest possible Double - note it is unnormalized (1-bit precision)
Private Const TinyTiny_c As Double = 4.94065645841247E-324
Private Const Silly_c As Double = -1.23456E+302  ' a really unlikely return value

Private cn_m As Double  ' cache of previously-calculated values
Private dn_m As Double
Private sn_m As Double

Private ƒ¤ As Integer  ' file unit for unit test file output

'########################### Exported Routines #################################

'===============================================================================
Function EllipticsVersion(Optional ByVal trigger As Variant) As String
' Date of the latest revision to this code, as a string with format "yyyy-mm-dd"
EllipticsVersion = Version_c
End Function

'===============================================================================
Function completeElliptic1(ByVal parameter As Double) As Double
' Return complete elliptic integral of the first kind for the argument domain
' "-infinity" <= parameter <= 1 (actually -1.79769313486E308 <= parameter <= 1).
'
'  K(m) = Integral of 1 / Sqr((1 - t^2)*(1 - m*t^2)) over 0 <= t <= 1
'
' Note that this is 1/4 of the real period of Cn & Sn, and 1/2 the period of Dn,
' but you should use jacobiCycleOvr4() if you want to find the the period of any
' Jacobi elliptic function, since it works for parameter values that are > 1.
' Warning - you should use "completeElliptic1_comp" instead if at all possible.
' The largest value of the parameter, below unity, is (1 - MachineEpsilonM_c)
' which gives 19.7546946459584, however the results are very granular when the
' parameter is that close to unity (for example, the next smaller Double, which
' is (1 - MachineEpsilonP_c) leads to 19.4081210556785) and the user should be
' cautious. Note that parameter = 1 returns completeElliptic1_max() = 700, not
' the mathematically correct value of infinity; see "completeElliptic1_comp" for
' details. The absolute error is 5E-15 or less, and the relative error is 3E-15
' or less, both ruled by roundoff error in the polynomial evaluation; the
' relative error deteriorates steadily for m > 0.999 (but why are you there?)
' The parameter here is the one used in the tables in AMS55 and denoted by "m".
' One often sees "k" used instead, where m = k^2. Take care. Notations vary.
completeElliptic1 = completeElliptic1_comp(1# - parameter)
End Function

'===============================================================================
Function completeElliptic1_comp(ByVal oneMinusParam As Double) As Double
' Return the complete elliptic integral of the first kind for the arguments of
' 0 <= oneMinusParam <= "infinity" (actually 0 <= arg <= 1.79769313486E308).
' Note that this is 1/4 of the real period of Cn & Sn, and 1/2 the period of Dn,
' but you should use jacobiCycleOvr4() if you want to find the the period of any
' Jacobi elliptic function, since it works for oneMinusParam values that are < 0.
' Mathematically, K(m) = K_comp(1 - m), but use of floating point spoils this.
' The argument is the complementary parameter (1 - parameter) because the
' function has a singularity at parameter = 1, and floating point granularity
' leads to jumpy behavior near there. This is reduced when the origin is put at
' the singularity, as it is here. We silently avoid an infinite return from zero
' argument by instead returning a value that is larger than possible using IEEE
' 754 arithmetric, yet small enough that it will not cause exponential overflow.
' Thus completeElliptic1_max() = 700, not infinity, is completeElliptic1_comp(0).
' The absolute error is 5E-15 or less, and the relative error is 3E-15 or less,
' both ruled by roundoff error in the polynomial evaluation; the relative error
' is good to at least m1 < 1E-10 (and probably 1E-300, but why are you there?).
' The parameter here is the one used in the tables in AMS55 and denoted by "m".
' One often sees "k" used instead, where m = k^2. Take care. Notations vary.
Const ID_ As String = File_c & " : completeElliptic1_comp"
Dim m1 As Double  ' this is the complementary parameter m1 = 1 - m
m1 = oneMinusParam  ' shorthand for 1 - m
If m1 < 0# Then  ' out of our domain (but you can do elliptic functions here)
  ' result is complex, with real part equal to an incomplete elliptic integral
  ' see AMS55 17.4.15 for the relation
  Err.Raise Invalid_c, ID_, _
    "Domain ERROR: expected m < 1 & 0 < m1 but got" & vbLf & _
    "m = " & CStr(1# - m1) & "   m1 = " & CStr(m1) & vbLf & _
    "Problem in " & ID_
End If

' The smallest Double is TinyTiny_c = 4.94065645841247E-324 (it's unnormalized),
' which leads to a return value of completeElliptic1_comp = 373.606330321811.
' Function values near this IEEE 754 lower limit are very granular; beware!
' Input arguments within 1/2 of TinyTiny_c are set to TinyTiny_c by VBA.
' Values below TinyTiny_c / 2 are set exactly to zero, where function = infinity;
' we instead return an integer value that is larger than possible in IEEE 754.
' Note that Excel sets function inputs below 2.225E-308 to zero, causing a jump
' there and the return of the 'impossible' value.
If m1 = 0# Then ' m1 was too close to 0 to be represented as a Double
  completeElliptic1_comp = completeElliptic1_max()  ' = 700
  Exit Function  ' no need to cache the value; evaluation is "instantaneous"
End If

' we return the previous value if the argument has not changed, to save time
Static init_s As Boolean, argWas_s As Double, valWas_s As Double
If Not init_s Then  ' on cold start, or after reset, values are invalid
  argWas_s = 1#  ' pretend the previous value was m1 = 1 (m = 0)
  Const PiOvr2_c As Double = 1.5707963 + 2.67948965E-08  ' need sum for accuracy
  valWas_s = PiOvr2_c  ' function = Pi / 2 at m = 0, m1 = 1
  init_s = True  ' bypass this once a valid "previous" value is available
End If
If m1 = argWas_s Then  ' this is a repeat of the previous call
  completeElliptic1_comp = valWas_s  ' use the cached result
  Exit Function  '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
End If

' argument has changed, so actually do the calculation
argWas_s = m1  ' save arg value in case next call has same argument
Dim p As Double, q As Double, r As Double
' allow for negative m, which is m1 > 1, by using inverse transformation
' see AMS55 17.4.17, which implies CEI1_comp(m1) = CEI1_comp(1 / m1) / Sqr(m1)
r = 1#           ' the "usual" case - no scaling needed
If m1 > 1# Then  ' m1 > 1 so use reciprocal-m1 relation
  m1 = 1# / m1   ' now 0 < m1 < 1, so we can use Cody's result...
  r = Sqr(m1)    ' ...but we have to scale the result
End If

' we use Cody's result, which is based on Hasting's form P - Q * log(m1), which
' in turn is a modified Legendre form shown by Airey in his 1935 paper:
' J. R. Airey, "Toroidal functions and complete elliptic integrals,"
' Phil. Mag., s. 7, v. 19, Jan. 1935, pp. 177-188
' see W. J. Cody, "Chebyshev Approximations for the Complete Elliptic Integrals
' K and E", Mathematics of Computation 19, 1965, pp. 105-112, Table II, n = 9
' note that Cody's result has equal-absolute-error ripple, but should have had
' equal-relative-error ripple, since the function value has large variation
' that's OK here because the error is dominated by IEEE 754 roundoff noise

' polynomial coefficients, to maximum possible accuracy
' if not written as sums, you get only 15-digit accuracy reading from file (yes!)
' to match the asymptotic behavior exactly, we need p0 = ln(4) & q0 = 1 / 2
' note: the choice of the p0 constant here depends on the use of the function
' if it will be called by VBA, use the "correct" value of ln(4)
' if it will be called from Excel, use the other value to equalize + and - error
' it is not clear why Excel offsets the result by about 5E-15, but it does
' note that correct behavior at m = 1 requires that sum of Pj's = Pi / 2
Const p0 As Double = 1.3862943611 + 1.989062E-11  ' ln(4) exactly, for VBA
'Const p0 As Double = 1.3862943611 + 1.98857E-11  ' constant for Excel use
Const p1 As Double = 0.096573590301 + 7.425285E-13
Const p2 As Double = 0.030885173001 + 8.997099E-13
Const p3 As Double = 0.014942029142 + 2.820783E-13
Const p4 As Double = 0.0089266462945 + 5.64662E-14
Const p5 As Double = 0.0075193867218 + 8.38102E-15
Const p6 As Double = 0.01058995362 + 9.893585E-13
Const p7 As Double = 0.01079599049 + 5.916349E-13
Const p8 As Double = 0.003968470902 + 9.897819E-14
Const p9 As Double = 0.00030072519903 + 6.864838E-15

Const q0 As Double = 0.5
Const q1 As Double = 0.12499999999 + 7.640658E-12
Const q2 As Double = 0.070312495459 + 5.466082E-13
Const q3 As Double = 0.048827155048 + 1.180099E-13
Const q4 As Double = 0.037335546682 + 2.860296E-13
Const q5 As Double = 0.029503729348 + 6.88713E-13
Const q6 As Double = 0.020690240005 + 1.008404E-13
Const q7 As Double = 0.0092811603829 + 6.860419E-14
Const q8 As Double = 0.0017216147097 + 9.865212E-14
Const q9 As Double = 0.000066631752464 + 6.073151E-16

' uncomment this block to check the coefficient values by hand
'Debug.Print p0, p1, p2, p3  ' uncomment block to check coefficients & sum
'Debug.Print p4, p5, p6, p7
'Debug.Print p8, p9
'' the sum should be exactly Pi / 2, so after subtraction we want zero
'Debug.Print p0 + p1 + p2 + p3 + p4 + p5 + p6 + p7 + p8 + p9 - 2# * Atn(1#)
'Debug.Print q0, q1, q2, q3
'Debug.Print q4, q5, q6, q7
'Debug.Print q8, q9

' split evaluation into two parts to avoid "Expression Too Complex" error
p = p3 + (p4 + (p5 + (p6 + (p7 + (p8 + p9 * m1) * m1) * m1) * m1) * m1) * m1
p = p0 + (p1 + (p2 + p * m1) * m1) * m1
q = q3 + (q4 + (q5 + (q6 + (q7 + (q8 + q9 * m1) * m1) * m1) * m1) * m1) * m1
q = q0 + (q1 + (q2 + q * m1) * m1) * m1
valWas_s = r * (p - q * Log(m1))  ' save func value in case same arg next call
completeElliptic1_comp = valWas_s
End Function

'===============================================================================
Function completeElliptic1_max() As Double
' The maximum value that will be returned from the completeElliptic1 functions,
' when parameter = 1 and oneMinusParam = 0, in lieu of the true value of
' infinity. We return a large number, but not so large that it will cause Exp to
' overflow or underflow when used as an argument. It is an integer to catch the
' user's attention.
completeElliptic1_max = 700#
End Function

'===============================================================================
Function jacobiCn(ByVal x As Double, ByVal parameter As Double) As Double
' Compute and save all three Jacobi elliptic functions, then return Cn(x|m)
' the input argument is the parameter (modulus^2 = parameter), as used in AMS55
jacobiCnDnSn x, parameter, cn_m, dn_m, sn_m
jacobiCn = cn_m
End Function

'===============================================================================
Function jacobiCycleOvr4(ByVal parameter As Double) As Double
' Return one quarter of the full period (or cycle) of the Jacobi elliptic
' functions Cn(x|m) and Sn(x|m), which is one half the period of Sn(x|m). This
' function is supplied so you can get the correct answer when parameter > 1,
' where the complete elliptic integral is complex but the elliptic function
' period is well defined. When parameter < 1, this is the same as
' completeElliptic1 except for near-roundoff arguments.
Dim m1 As Double  ' force granularity of 1.11E-16 for all param's
m1 = 1# - parameter  ' complementary parameter = 1 - parameter = 1 - m = m1
If m1 = 0# Then  ' special case, in lieu of infinity
  jacobiCycleOvr4 = completeElliptic1_max()
ElseIf m1 > 0# Then  ' the "normal" case
  jacobiCycleOvr4 = completeElliptic1_comp(m1)
Else  ' m > 1 & m1 < 0 so use relation in AMS55 17.4.17 (note F = 0)
  Dim temp As Double
  temp = 1# / (1# - m1)  ' just divide once - no cancellation since m1 < 0
  jacobiCycleOvr4 = completeElliptic1_comp(-m1 * temp) * Sqr(temp)
End If
End Function

'===============================================================================
Sub jacobiCnDnSn(ByVal u As Double, ByVal parameter As Double, _
  ByRef cn As Double, ByRef dn As Double, ByRef sn As Double)
' Return all three Jacobi elliptic functions by altering ByRef arguments, taking
' the parameter (modulus^2 = parameter) as the second argument.
' The parameter here is the one used in the tables in AMS55 and denoted by "m".
' One often sees "k" used instead, where m = k^2. Take care. Notations vary.
'
' The code is taken from T. Fukushima, "Precise & Fast Computation of Jacobian
' Elliptic Functions by Conditional Duplication," Numer. Math. (2013) 123,
' pp. 585-605, plus the range reduction and folding logic in his paper "Fast
' Computation of Jacobian Elliptic Functions & Incomplete Elliptic Integrals
' for Constant Values of Elliptic Parameter & Elliptic Characteristic," Celest
' Mech Dyn Astr (2009) 105 pp. 245-260, with simplifications & improvements
' by John Trenholme

' Note that 1/4 of the full period of Cn(u|m) and Sn(u|m), and 1/2 the period of
' Dn(u|m), is given by the function jacobiCycleOvr4(), which works even for
' values of parameter > 1.

' WARNING - this code would fail with divide-by-zero for parameter values very
' close to unity - you can not, and should not, use values closer below 1 than
' 1.11E-16, or closer above 1 than 2.22E-16. Exactly 1 is OK; it gives sech
' and tanh functions of "infinite" period.
Const ID_ As String = File_c & " : jacobiCnDnSn"

Dim m1 As Double
m1 = 1# - parameter  ' complementary parameter = 1 - parameter = 1 - m = m1
Dim m As Double      ' parameter value, but with 1 or 2 E-16 granularity near 0
m = 1# - m1          ' might lose digits; if 4.94E-324 < m1 < 1.11E-16, m = 1

' test for repeated arguments; if found, return cached values
Static init_s As Boolean, uWas_s As Double, m1Was_s As Double
If Not init_s Then  ' cold start or reset; values invalid; set good values
  uWas_s = 0#   ' the values below are correct for any parameter value
  m1Was_s = 1#  ' parameter = 0, so pretending we used trig functions
  cn_m = 1#
  dn_m = 1#
  sn_m = 0#
  init_s = True  ' bypass this, once valid "previous" values are available
End If
If (u = uWas_s) And (m1 = m1Was_s) Then  ' this is a repeat of the last call
  cn = cn_m
  dn = dn_m
  sn = sn_m
  Exit Sub  '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
End If
' save function arguments for the next call
uWas_s = u
m1Was_s = m1

If m1 = 1# Then  ' unlikely special case: trigonometric functions
  cn_m = Cos(u)  ' period is 2 * Pi
  dn_m = 1#
  sn_m = Sin(u)
  cn = cn_m
  dn = dn_m
  sn = sn_m
  Exit Sub  '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
ElseIf m1 = 0# Then  ' unlikely special case: hyperbolic functions
  ' infinite "period", but we fail at +/- 709 (and, what are you doing there?)
  If Abs(u) > 709.782712893384 Then   ' Exp(u) will overflow or underflow
    Err.Raise Invalid_c, ID_, _
      "Invalid function argument (Exp overflow)" & vbLf & _
      "Need |u| <= 709.782712893384 but u = " & u & vbLf & _
      "Problem in " & ID_
  End If
  Dim temp As Double  ' Cn is sech(u), which does not suffer from roundoff
  temp = Exp(u)
  cn_m = 2# / (temp + 1# / temp)
  dn_m = cn_m  ' Dn = Cn = sech
  Dim au As Double  ' Sn is tanh(u), which has roundoff noise near the origin
  au = Abs(u)
  If au < 0.503 Then  ' roundoff would spoil mathematical identity
    Dim u2 As Double
    u2 = u * u
    ' worst relative error 3E-16 - see the Maple file "TanhApprox.mws"
    sn_m = u - u * (19.0563004085872 + (0.9231700917023 + _
      0.0003630096197 * u2) * u2) * u2 / _
      (57.1689012257678 + (25.637070765062 + u2) * u2)
  ElseIf au < 20# Then  ' mathematical identity with Exp(2 * u) is accurate
    temp = Exp(u + u)  ' won't overflow, with |u| < 20
    sn_m = (temp - 1#) / (temp + 1#)
  Else  ' |u| >= 20, so result is +1 or -1 to IEEE 754 accuracy
    sn_m = Sgn(u)
  End If
  cn = cn_m
  dn = dn_m
  sn = sn_m
  Exit Sub  '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
End If

Dim wasNeg As Boolean, wasBig As Boolean  ' flags to be used at end of routine
wasNeg = False
wasBig = False
' test if m1 > 1 & so m < 0; if so, force parameter to 0 < m < 1
If m1 > 1# Then  ' m < 0 & m1 > 1 so use relations in AMS55 16.10
  wasNeg = True
  Dim divNeg As Double
  divNeg = Sqr(m1)
  u = u * divNeg
  m1 = 1# / m1  ' this is the same as m = -m / (1 - m)
  m = 1# - m1
' test if m1 < 0 & so m > 1, if so, force parameter to 0 < m < 1
ElseIf m1 < 0# Then  ' m > 1 & m1 < 0 so use relations in AMS55 16.11
  wasBig = True
  Dim divBig As Double
  divBig = Sqr(m)
  u = u * divBig
  m1 = -m1 / m  ' this is the same as m = 1 / m
  m = 1# - m1
End If
'Debug.Assert (m1 > 0#) And (m1 < 1#)

' now 0 < m1 < 1, so we can find the quarter period & related things
Dim k1 As Double, k2 As Double, k3 As Double, k4 As Double
k1 = completeElliptic1_comp(m1)  ' must be Pi / 2 <= k1 <= 19.7546946459584
k2 = k1 + k1  ' this is the full period of Dn
k3 = k1 + k2
k4 = k2 + k2  ' this is the full period of Cn & Sn

' now reduce the argument u to the standard period domain 0 <= u <= k4
' note that this will lose digits if we are a long way away
If (u < 0#) Or (u > k4) Then u = u - k4 * Int(u / k4)  ' slow on Intel chips
'Debug.Assert (u >= 0#) And (u <= k4)

' fold the four domain sections into 0 <= u <= k1 by signs & reflections
Dim cSgn As Double, sSgn As Double
cSgn = 1#
sSgn = 1#
If u > k3 Then  ' zone 3: reflect around end at u = k4 & flip Sn
  u = k4 - u
  sSgn = -1#
ElseIf u > k2 Then  ' zone 2:  translate to zone 0 & flip both Cn & Sn
  u = u - k2
  cSgn = -1#
  sSgn = -1#
ElseIf u > k1 Then  ' zone 1: reflect around center at u = k2 & flip Cn
  u = k2 - u
  cSgn = -1#
End If  ' base zone 0: do nothing
'Debug.Assert (u >= 0#) And (u <= k1)

' reduce the active domain to 0 <= u <= k1 / 2 by reflection in k1 & more work
' see addition relations in AMS55 16.17 for u = -u, v = k1, or Fukushima 2009
Dim foldK1 As Boolean
foldK1 = False
If u + u > k1 Then
  foldK1 = True
  u = k1 - u
End If
'Debug.Assert (u >= 0#) And (u <= 0.5 * k1)

Dim uT As Double  ' once u is below uT, accuracy = double-precision roundoff
uT = 0.005217 - m1 * 0.002143  ' see Fukushima's 2013 paper, appendix A.3
Dim n As Long
Dim u0 As Double  ' reduced argument u0 = u / (2 ^ N) < uT
u0 = u + u  ' back up so first test is u0 = u
Const nMax_ As Long = 20&  ' this should be safe even for m1 = 0, u = 2800
For n = 0& To nMax_  ' we should end up with N = 8 to 12 in most cases
  u0 = 0.5 * u0  ' make u0 smaller and smaller, for half-argument  iterations
  If u0 < uT Then Exit For  ' u0 is small enough, so remember N for later
Next n
If u0 >= uT Then  ' we did not get u0 below uT - something is horribly wrong
  ' worst N is when m1 = 1, u = k4, where N = Ln2(k4 / (0.005217 - 0.002143))
  ' this is 20 when k4 = 3223, but 4 * completeElliptic1_max() = 2800
  cn_m = Silly_c  ' just in case the user foolishly trys to proceed
  dn_m = Silly_c
  sn_m = Silly_c
  cn = cn_m
  dn = dn_m
  sn = sn_m
  m1Was_s = Silly_c  ' to stop any reading of the cache
  Err.Raise Invalid_c, ID_, _
    "Domain ERROR: input argument u = " & u & vbLf & _
    "Complementary parameter oneMinusParam = " & m1 & vbLf & _
    "Reduced u by factor " & u0 / u & " to " & u0 & vbLf & _
    "but needed u0 < uT = " & uT & vbLf & _
    "Problem in " & ID_
End If

Dim v As Double  ' u0 ^ 2
v = u0 * u0
' constant factors in the two-variable polynomial below
Const p0 As Double = 1# / 24#, p1 As Double = 1# / 6#
Const q0 As Double = 1# / 720#, q1 As Double = 11# / 180#, _
  q2 As Double = 1# / 45#
Dim b As Double  ' b is the complement of the Cn function, b = 1 - Cn
' initial value of b is a polynomial in u0 ^ 2 & m chosen to speed convergence
' see Fukushima 2013 Appendix A.8 & A.12 for the gory details
' rewritten from Fukushima's form to avoid "Expression Too Complex" error
b = (0.5 - (p0 + p1 * m - (q0 + (q1 + q2 * m) * m) * v) * v) * v
Dim a As Double  ' b will be found from b / a = numerator / denominator
a = 1#  ' initial denominator
Dim uA As Double  ' rough point where would we lose precision by subraction
uA = 1.76269 + m1 * 1.16357  ' see Fukushima's 2013 paper, appendix A.9
Dim j As Long  ' iteration counter
Dim switched As Boolean  ' flag for switching variables
switched = False
Dim y As Double, z As Double, my As Double, mc2 As Double, m2 As Double
If u < uA Then  ' there will not be loss of precision by subtraction
  ' this is the "fast" inner loop, where no test is needed
  For j = 1& To n  ' we iterate a double-angle formula, avoiding division
    y = b * (a + a - b)
    z = a * a
    my = m * y
    ' the double-argument formula is b(J+1) = 2*y(J)*(1-m*y(j))/(1-m*Y(J)^2)
    b = (y + y) * (z - my)
    a = z * z - my * y
  Next j
Else  ' there could be loss of precision by subtraction; monitor possibility
  ' this is the "slow" inner loop, with a test for impending precision loss
  For j = 1& To n
    y = b * (a + a - b)
    z = a * a
    my = m * y
    If z < my + my Then  ' precision loss impending; switch variables
      switched = True
      Exit For  ' loop will be completed at "If switched" statement below
    End If
    b = (y + y) * (z - my)
    a = z * z - my * y
  Next j
End If
Dim c As Double, d As Double, s As Double
If switched Then  ' finish up the inner loop with the switched variable
  Dim x As Double, w As Double, xz As Double
  c = a - b  ' switch from b = 1 - c to c as the iterated quantity
  mc2 = m1 + m1
  m2 = m + m
  For j = j To n  ' note we start at j, not 1
    x = c * c
    z = a * a
    w = m * x * x - m1 * z * z
    xz = x * z
    c = mc2 * xz + w
    a = m2 * xz - w
  Next j
  If a = 0# Then  ' this is not supposed to happen, but...
    Err.Raise Invalid_c, ID_, _
      "Domain ERROR: input argument u = " & u & vbLf & _
      "Complementary parameter oneMinusParam = " & m1 & vbLf & _
      "Iteration denominator = 0 - can't proceed [1]" & vbLf & _
      "Problem in " & ID_
  End If
  c = c / a
  x = c * c
  s = Sqr(1# - x)
  d = Sqr(m1 + m * x)
Else  ' didn't switch, or didn't even test, so just finish up
  If a = 0# Then  ' this is not supposed to happen, but...
    Err.Raise Invalid_c, ID_, _
      "Domain ERROR: input argument u = " & u & vbLf & _
      "Complementary parameter oneMinusParam = " & m1 & vbLf & _
      "Iteration denominator = 0 - can't proceed [2]" & vbLf & _
      "Problem in " & ID_
  End If
  b = b / a
  y = b * (2# - b)
  c = 1# - b
  s = Sqr(y)
  d = Sqr(1# - m * y)
End If

' at this point, we have values for 0 <= u <= k1 / 2
Dim rootM1 As Double, dInv As Double, cTemp As Double
If foldK1 Then  ' results that were above k1 / 2 need to be fixed up
  dInv = 1# / d
  rootM1 = Sqr(m1)  ' the price is a square root
  d = rootM1 * dInv
  cTemp = c
  c = d * s
  s = cTemp * dInv
End If

' now have vaues for 0 <= u <= k1; apply the signs for the quadrant we are in
c = c * cSgn
s = s * sSgn

' at this point, we have values for 0 <= u <= k4; fix up for "funny" parameters
If wasNeg Then  ' m < 0 & m1 > 1 so use relations in AMS55 16.10
  d = 1# / d  ' Dn(u|m) is now always >= 1
  c = c * d
  s = s * d / divNeg
ElseIf wasBig Then  ' m > 1 & m1 < 0 so use relations in AMS55 16.11
  cTemp = c       ' Cn(u|m) and Dn(u|m) swap roles
  c = d
  d = cTemp
  s = s / divBig  ' Sn(u|m) is reduced in magnitude (|max| < 1)
End If

' set the output values, and the cached values
cn_m = c
dn_m = d
sn_m = s
cn = c
dn = d
sn = s
End Sub

'===============================================================================
Function jacobiDn(ByVal x As Double, ByVal parameter As Double) As Double
' Compute and save all three Jacobi elliptic functions, then return Dn(x|m)
' the input argument is the parameter (modulus^2 = parameter), as used in AMS55
jacobiCnDnSn x, parameter, cn_m, dn_m, sn_m
jacobiDn = dn_m
End Function

'===============================================================================
Function jacobiSn(ByVal x As Double, ByVal parameter As Double) As Double
' Compute and save all three Jacobi elliptic functions, then return Sn(x, m)
' the input argument is the parameter (modulus^2 = parameter), as used in AMS55
jacobiCnDnSn x, parameter, cn_m, dn_m, sn_m
jacobiSn = sn_m
End Function

'########################### Unit Tests ########################################

#If True Then  ' "False" to hide unit tests, "True" to expose

'===============================================================================
Sub ellipticsUnitTest()
' Carry out some high-accuracy spot checks of elliptic integrals & functions
' results are sent to Immediate window; copy from there and paste into a file
Dim pathx As String
pathx = Environ$("UserProfile") & "\Desktop\"  ' put output on desktop
Dim fileName As String
fileName = pathx & "EllipticsUnitTest_" & TimeStamp() & ".txt"
ƒ¤ = FreeFile  ' get a free file unit (note: ƒ¤ is module global)
Open fileName For Output Access Write Lock Write As #ƒ¤

teeOut "===== Unit Tests of Elliptic Function & Integral Routines ====="
teeOut "File " & File_c & "   Now " & TimeStamp()
teeOut
Dim j As Long, param As Double, ce As Double
Dim vals() As Variant  ' Array has to be variants
' "exact" Maple 18-digit values for comparison with completeElliptic1()
vals = Array(1.5335928197134 + 5.688E-14, _
  1.5707963267948 + 9.662E-14, 1.6124413487202 + 1.94E-14, _
  1.6596235986105 + 2.8001E-14, 1.7138894481787 + 9.106E-14, _
  1.7775193714912 + 5.332E-14, 1.8540746773013 + 7.192E-14, _
  1.949567749806 + 2.588E-14, 2.0753631352924 + 6.914E-14, _
  2.2572053268208 + 5.3655E-14, 2.5780921133481 + 7.319E-14, _
  3.6956373629898 + 7.468E-14, 4.8411325605502 + 9.703E-14, _
  5.9915893405069 + 9.64E-14, 7.1427724505817 + 7.819E-14, _
  8.2940514636154 + 3.999E-14, 9.4453423977326 + 1.682E-14, _
  10.596634757087 + 6.603E-13, 11.74792728228 + 7.802E-14)
teeOut "-- Tests of complete elliptic integral of the first kind"
teeOut "parameter", "this code", , "Maple 60-digit result", "relative error"
' note that we are actually using the complementary parameter here
Dim parComp As Double
For j = -1& To 17&  ' step over param = -0.1, 0, 0.1, 0.2, ... but...
  ' ...after param = 0.9, get closer and closer to 1 by inverse powers of 10
  If j < 10& Then parComp = 1# - j * 0.1 Else parComp = 10# ^ (8# - j)
  ce = completeElliptic1_comp(parComp)  ' completeElliptic1() gets noisy at 0.999
  teeOut 1# - parComp, ce, vals(j + 1&), CSng(ce / vals(j + 1&) - 1#)
Next j
' test the cycle length of the Jacobi elliptics
Const Je_ As Double = 1.311028777146 + 5.9905E-14  ' from Maple
Dim u As Double
u = 2#
ce = jacobiCycleOvr4(u)
teeOut "Jacobi-function quarter cycle length test - parameter = " & u
teeOut u, ce, Je_, CSng(ce / Je_ - 1#)
'GoTo Done_L  ' uncomment to just test the complete elliptic integral
teeOut

teeOut "-- Tests of the Jacobi elliptic functions"
teeOut "", "this code", , "Maple result", , "relative error"
teeOut "Test at the 16 'folded' points of one base cycle " & _
  "(tests quadrant signs & K/2 folding)"
Dim arg As Double, per As Double
arg = 0.5
param = 0.643856219147755
per = 4# * completeElliptic1(param)
teeOut "argument = " & arg & " (1/16 of period)  parameter = " & param & _
  "  period = " & per
' "exact" values from Maple JacobiPQ(arg, Sqrt(param)); see EllipticStuff.mw
Dim cn As Double, dn As Double, sn As Double
Dim cnEx As Double, dnEx As Double, snEx As Double
cn = jacobiCn(arg, param): cnEx = 0.88358957167486 + 9.103E-15
dn = jacobiDn(arg, param): dnEx = 0.92672649084087 + 3.698E-15
sn = jacobiSn(arg, param): snEx = 0.46826217958257 + 2.477E-15
teeOut "cn", cn, cnEx, CSng(cn / cnEx - 1#)
teeOut "dn", dn, dnEx, CSng(dn / dnEx - 1#)
teeOut "sn", sn, snEx, CSng(sn / snEx - 1#)
arg = 1.5
teeOut "argument = " & arg & " (3/16 of period)  parameter = " & param & _
  "  period = " & per
cn = jacobiCn(arg, param): cnEx = 0.30154364975464 + 8.929E-15
dn = jacobiDn(arg, param): dnEx = 0.64396328147504 + 1.82E-15
sn = jacobiSn(arg, param): snEx = 0.95345237284965 + 9.237E-15
teeOut "cn", cn, cnEx, CSng(cn / cnEx - 1#)
teeOut "dn", dn, dnEx, CSng(dn / dnEx - 1#)
teeOut "sn", sn, snEx, CSng(sn / snEx - 1#)
arg = 2.5
teeOut "argument = " & arg & " (5/16 of period)  parameter = " & param & _
  "  period = " & per
cn = jacobiCn(arg, param): cnEx = -0.30154364975464 - 8.929E-15
dn = jacobiDn(arg, param): dnEx = 0.64396328147504 + 1.82E-15
sn = jacobiSn(arg, param): snEx = 0.95345237284965 + 9.237E-15
teeOut "cn", cn, cnEx, CSng(cn / cnEx - 1#)
teeOut "dn", dn, dnEx, CSng(dn / dnEx - 1#)
teeOut "sn", sn, snEx, CSng(sn / snEx - 1#)
arg = 3.5
teeOut "argument = " & arg & " (7/16 of period)  parameter = " & param & _
  "  period = " & per
cn = jacobiCn(arg, param): cnEx = -0.88358957167486 - 9.103E-15
dn = jacobiDn(arg, param): dnEx = 0.92672649084087 + 3.698E-15
sn = jacobiSn(arg, param): snEx = 0.46826217958257 + 2.477E-15
teeOut "cn", cn, cnEx, CSng(cn / cnEx - 1#)
teeOut "dn", dn, dnEx, CSng(dn / dnEx - 1#)
teeOut "sn", sn, snEx, CSng(sn / snEx - 1#)
arg = 4.5
teeOut "argument = " & arg & " (9/16 of period)  parameter = " & param & _
  "  period = " & per
cn = jacobiCn(arg, param): cnEx = -0.88358957167486 - 9.103E-15
dn = jacobiDn(arg, param): dnEx = 0.92672649084087 + 3.698E-15
sn = jacobiSn(arg, param): snEx = -0.46826217958257 - 2.477E-15
teeOut "cn", cn, cnEx, CSng(cn / cnEx - 1#)
teeOut "dn", dn, dnEx, CSng(dn / dnEx - 1#)
teeOut "sn", sn, snEx, CSng(sn / snEx - 1#)
arg = 5.5
teeOut "argument = " & arg & " (11/16 of period)  parameter = " & param & _
  "  period = " & per
cn = jacobiCn(arg, param): cnEx = -0.30154364975464 - 8.929E-15
dn = jacobiDn(arg, param): dnEx = 0.64396328147504 + 1.82E-15
sn = jacobiSn(arg, param): snEx = -0.95345237284965 - 9.237E-15
teeOut "cn", cn, cnEx, CSng(cn / cnEx - 1#)
teeOut "dn", dn, dnEx, CSng(dn / dnEx - 1#)
teeOut "sn", sn, snEx, CSng(sn / snEx - 1#)
arg = 6.5
teeOut "argument = " & arg & " (13/16 of period)  parameter = " & param & _
  "  period = " & per
cn = jacobiCn(arg, param): cnEx = 0.30154364975464 + 8.929E-15
dn = jacobiDn(arg, param): dnEx = 0.64396328147504 + 1.82E-15
sn = jacobiSn(arg, param): snEx = -0.95345237284965 - 9.237E-15
teeOut "cn", cn, cnEx, CSng(cn / cnEx - 1#)
teeOut "dn", dn, dnEx, CSng(dn / dnEx - 1#)
teeOut "sn", sn, snEx, CSng(sn / snEx - 1#)
arg = 7.5
teeOut "argument = " & arg & " (15/16 of period)  parameter = " & param & _
  "  period = " & per
cn = jacobiCn(arg, param): cnEx = 0.88358957167486 + 9.103E-15
dn = jacobiDn(arg, param): dnEx = 0.92672649084087 + 3.698E-15
sn = jacobiSn(arg, param): snEx = -0.46826217958257 - 2.477E-15
teeOut "cn", cn, cnEx, CSng(cn / cnEx - 1#)
teeOut "dn", dn, dnEx, CSng(dn / dnEx - 1#)
teeOut "sn", sn, snEx, CSng(sn / snEx - 1#)
teeOut

teeOut "Test negative argument (should be same as 15/16 of period)"
arg = -0.5
teeOut "argument = " & arg & " (-1/16 of period)  parameter = " & param & _
  "  period = " & per
cn = jacobiCn(arg, param): cnEx = 0.88358957167486 + 9.103E-15
dn = jacobiDn(arg, param): dnEx = 0.92672649084087 + 3.698E-15
sn = jacobiSn(arg, param): snEx = -0.46826217958257 - 2.477E-15
teeOut "cn", cn, cnEx, CSng(cn / cnEx - 1#)
teeOut "dn", dn, dnEx, CSng(dn / dnEx - 1#)
teeOut "sn", sn, snEx, CSng(sn / snEx - 1#)
teeOut

Const Cycles_ As Double = 100000#
teeOut "Test multiple cycles (in 100,000 cycles, lose around 5 digits)"
arg = Cycles_ * per + 0.5 ' mathematically same as 0.5 case above
teeOut "argument = " & arg & " (" & Cycles_ & "+1/16 period)  parameter = " & _
  param & "  period = " & per
' see EllipticStuff.mw 2014-07-16 equations (8)
cn = jacobiCn(arg, param): cnEx = 0.88358957167486 + 3.655E-15
dn = jacobiDn(arg, param): dnEx = 0.92672649084087 + 3.536E-16
sn = jacobiSn(arg, param): snEx = 0.46826217958258 + 2.756E-15
teeOut "cn", cn, cnEx, CSng(cn / cnEx - 1#)
teeOut "dn", dn, dnEx, CSng(dn / dnEx - 1#)
teeOut "sn", sn, snEx, CSng(sn / snEx - 1#)
teeOut

teeOut "Test small parameter - period close to 2 * Pi"
arg = 0.4
param = 0.0001
per = 4# * completeElliptic1(param)
teeOut "argument = " & arg & "  parameter = " & param & _
  "  period = " & per
cn = jacobiCn(arg, param): cnEx = 0.92106139629029 + 5.52E-16
dn = jacobiDn(arg, param): dnEx = 0.99999241767604 + 9.93E-16
sn = jacobiSn(arg, param): snEx = 0.38941739080808 + 9.542E-15
teeOut "cn", cn, cnEx, CSng(cn / cnEx - 1#)
teeOut "dn", dn, dnEx, CSng(dn / dnEx - 1#)
teeOut "sn", sn, snEx, CSng(sn / snEx - 1#)
teeOut

teeOut "Test parameter close to unity"
arg = 2.6
param = 0.99999999
per = 4# * completeElliptic1(param)
teeOut "argument = " & arg & "  parameter = " & param & _
  "  period = " & per
cn = jacobiCn(arg, param): cnEx = 0.14773216672437 + 5.262E-15
dn = jacobiDn(arg, param): dnEx = 0.14773219983074 + 3.33E-15
sn = jacobiSn(arg, param): snEx = 0.9890274045318 + 6.384E-15
teeOut "cn", cn, cnEx, CSng(cn / cnEx - 1#)
teeOut "dn", dn, dnEx, CSng(dn / dnEx - 1#)
teeOut "sn", sn, snEx, CSng(sn / snEx - 1#)
teeOut

teeOut "Test negative parameter - note Dn(u|m) is >= 1"
arg = 0.3
param = -1.5
per = 4# * completeElliptic1(param)
teeOut "argument = " & arg & "  parameter = " & param & _
  "  period = " & per
cn = jacobiCn(arg, param): cnEx = 0.9533438347308 + 4.321E-15
dn = jacobiDn(arg, param): dnEx = 1.0661628858533 + 5.181E-15
sn = jacobiSn(arg, param): snEx = 0.3018866223945 + 8.99E-15
teeOut "cn", cn, cnEx, CSng(cn / cnEx - 1#)
teeOut "dn", dn, dnEx, CSng(dn / dnEx - 1#)
teeOut "sn", sn, snEx, CSng(sn / snEx - 1#)
teeOut

teeOut "Test parameter greater than unity"
arg = 0.3
param = 2.5
per = 4# * jacobiCycleOvr4(param)  ' can't use "completeElliptic1" here
teeOut "argument = " & arg & "  parameter = " & param & _
  "  period = " & per
cn = jacobiCn(arg, param): cnEx = 0.95851072192295 + 5.261E-15
dn = jacobiDn(arg, param): dnEx = 0.89266847715328 + 3.466E-15
sn = jacobiSn(arg, param): snEx = 0.28505647854194 + 6.374E-15
teeOut "cn", cn, cnEx, CSng(cn / cnEx - 1#)
teeOut "dn", dn, dnEx, CSng(dn / dnEx - 1#)
teeOut "sn", sn, snEx, CSng(sn / snEx - 1#)
teeOut

teeOut "-- Some simple timing tests for parameter = " & _
  "0.1, 0.5, 0.9, 0.99999999"
If Rnd(-1) >= 0! Then Randomize 1  ' replace 1 with any desired seed point
Dim elapse As Single, count As Double
count = 0#
elapse = Timer()
For j = 1& To 40000  ' note we don't do m < 0 or m > 1
  jacobiCnDnSn 0.4, 0.1, cn, dn, sn: jacobiCnDnSn 1.2, 0.1, cn, dn, sn
  jacobiCnDnSn 2#, 0.1, cn, dn, sn: jacobiCnDnSn 2.8, 0.1, cn, dn, sn
  jacobiCnDnSn 3.6, 0.1, cn, dn, sn: jacobiCnDnSn 4.4, 0.1, cn, dn, sn
  jacobiCnDnSn 5.2, 0.1, cn, dn, sn: jacobiCnDnSn 6#, 0.1, cn, dn, sn
    count = count + 8#
  jacobiCnDnSn 0.4, 0.5, cn, dn, sn: jacobiCnDnSn 1.2, 0.5, cn, dn, sn
  jacobiCnDnSn 2#, 0.5, cn, dn, sn: jacobiCnDnSn 2.8, 0.5, cn, dn, sn
  jacobiCnDnSn 3.6, 0.5, cn, dn, sn: jacobiCnDnSn 4.4, 0.5, cn, dn, sn
  jacobiCnDnSn 5.2, 0.5, cn, dn, sn: jacobiCnDnSn 6#, 0.5, cn, dn, sn
    count = count + 8#
  jacobiCnDnSn 0.65, 0.9, cn, dn, sn: jacobiCnDnSn 1.92, 0.9, cn, dn, sn
  jacobiCnDnSn 3.19, 0.9, cn, dn, sn: jacobiCnDnSn 4.46, 0.9, cn, dn, sn
  jacobiCnDnSn 5.73, 0.9, cn, dn, sn: jacobiCnDnSn 7#, 0.9, cn, dn, sn
  jacobiCnDnSn 8.27, 0.9, cn, dn, sn: jacobiCnDnSn 9.54, 0.9, cn, dn, sn
    count = count + 8#
  jacobiCnDnSn 2.8, 0.99999999, cn, dn, sn
  jacobiCnDnSn 8.4, 0.99999999, cn, dn, sn
  jacobiCnDnSn 14#, 0.99999999, cn, dn, sn
  jacobiCnDnSn 19.6, 0.99999999, cn, dn, sn
  jacobiCnDnSn 25.2, 0.99999999, cn, dn, sn
  jacobiCnDnSn 30.8, 0.99999999, cn, dn, sn
  jacobiCnDnSn 36.4, 0.99999999, cn, dn, sn
  jacobiCnDnSn 42#, 0.99999999, cn, dn, sn
    count = count + 8#
Next j
elapse = Timer() - elapse
teeOut "Elapsed time: " & elapse & " seconds   Calls: " & count
teeOut "Calculation rate: " & Format$(count / elapse, "0,0") & _
  " jacobiCnDnSn's per second (" & Round(1000000# * elapse / count, 3&) & _
  " µsec's each)"
teeOut

Done_L:  '<*><*><*><*><*><*><*><*><*><*><*><*><*><*><*><*><*><*><*><*><*><*><*>
teeOut "~~~~~~~~~~~~~~ unit tests complete ~~~~~~~~~~~~~~~~~"
Close #ƒ¤  ' close the file
End Sub

'===============================================================================
Public Sub teeOut(ParamArray arguments() As Variant)
' Prints 0 or more 'arguments' to the Immediate window if in IDE (so always in
' Excel), and also to the output file set up on unit "ƒ¤" if it is open (non-0).
' Each comma-delimited argument is sent to the next available 14-character tab
' zone, the same as Debug.Print's output (but you don't get ";" handling).
' It's best to send only strings or numbers to this Sub, to avoid evil type
' coercion. You can get away with (for example) Currency, Date, Empty (= null
' string), Null (= null string), and so forth, but be careful. Don't send an
' Array, Object, Collection, Dictionary, or non-simple Variants. Missing values
' are OK (even the first argument) and "print" as empty strings.
Dim ret As String  ' initialize result as first argument, or null string if none
ret = vbNullString  ' in case there is no argument at all
If 0 <= UBound(arguments) Then  ' there are 1 or more arguments, possibly Missing
  If IsMissing(arguments(0)) Then ret = vbNullString Else ret = arguments(0)
End If
Dim j As Long, k As Long
For j = 1 To UBound(arguments)  ' add on other arguments, tabbing 14 spaces
  k = Len(ret)
  ret = ret & Space$(14 * Int(k / 14 + 1) - k)
  If Not IsMissing(arguments(j)) Then ret = ret & arguments(j)
Next j
Debug.Print ret  ' send to Immediate window (no-op if not in IDE; unlikely)
If 0 <> ƒ¤ Then Print #ƒ¤, ret  ' try to print to file, if it has a unit number
End Sub

'===============================================================================
Private Function TimeStamp() As String
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

#End If  ' end of hide-expose unit tests

'~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


Attribute VB_Name = "Functions"
Attribute VB_Description = "Math Functions not available in the standard VB library. Devised & coded by John Trenholme."
'
'         _______                             _.
'        | ______)                       _   (_)
'        | |___   _   _  ____    ____  _| |_  _   ___   ____    ___
'        |  ___) | | | ||  _ \  / ___)(_   _)| | / _ \ |  _ \  /___)
'        | |     | |_| || | | |( (___   | |_ | || |_| || | | ||___ |
'        |_|     |____/ |_| |_| \____)   \__)|_| \___/ |_| |_|(___/
'
'
'###############################################################################
'#
'# Visual Basic 6 & Visual Basic for Applications Module "Functions.bas"
'#
'# Supplies mathematical Functions not available in the standard VB6 & VBA
'# library. Most of these Functions take Double arguments and return a Double
'# result.
'#
'# Note: there are functions here that duplicate some that VBA can supply
'# through a host's Application object, so (for example) Excel supplies cosh
'# (and acosh and ...) from Excel.Application.WorksheetFunction.cosh
'#
'# Where there is a name conflict, you can force the functions in this module
'# to be used by prefixing them with the module name. E.g., "Functions.Cosh".
'#
'# Note: where errors can occur, the rather terse default VB messages have been
'# augmented with additional information to aid the user.
'#
'# Devised & coded by John Trenholme - Started 2008-09-02
'#
'# This Module exports the routines:
'#   Function ArcCos
'#   Function ArcCoSec
'#   Function ArcCosh
'#   Function ArcCoTan
'#   Function ArcSec
'#   Function ArcSin
'#   Function ArcSinh
'#   Function ArcTanh
'#   Function Atan2
'#   Function CoSec
'#   Function Cosh
'#   Function CoTan
'#   Function Deg
'#   Function ExpM1
'#   Function ExpM1P
'#   Function FunctionsVersion
'#   Function Hypot
'#   Function IntSMooth
'#   Function InvPos
'#   Function LimitExp
'#   Function LimitHard
'#   Function LimitSqr
'#   Function Log10
'#   Function Log1P
'#   Function Log1plusP
'#   Function LogN
'#   Function Max
'#   Function MaxExp
'#   Function MaxSqr
'#   Function Min
'#   Function MinExp
'#   Function MinSqr
'#   Function NoZero
'#   Function Pow2
'#   Function Pow3
'#   Function Pow4
'#   Function Pow5
'#   Function Rad
'#   Function RankMedian
'#   Function Sec
'#   Function Sech
'#   Function SerialNum
'#   Function Sinh
'#   Function StdNormCum
'#   Function StdNormInvCum
'#   Function StepSqr
'#   Function Tanh
'#   Function TenTo
'#
'# This module requires the file:
'#   Formats.bas  (used by Hypot, Max, & Min to error-print Variant arg's)
'#
'###############################################################################

Option Base 0          ' array base value when not specified - the default
Option Compare Binary  ' string comparison based on Asc(char) - the default
Option Explicit        ' forces explicit variable declaration - changes default
'Option Private Module  ' globals project-only in VBA; no effect in VB6

Private Const Version_c As String = "2013-10-30"  ' date of latest revision
Private Const File_c As String = "Functions[" & Version_c & "]."

' All constants below are accurate to the last bit in IEEE 754 arithmetic
' They are written as sums to avoid precision loss when this file is written
' to, and then read from, a file (as when saving & loading a project). This
' impolite behaviour is a known defect in VB6 and VBA.

' Const values are supplied in each routine, rather than globally, so you can
' cut function code from here and paste it into your own code.

'===============================================================================
Public Function ArcCos(ByVal x As Double) As Double
Attribute ArcCos.VB_Description = "The angle in radians that gives the supplied cosine value 'x', which must obey -1 <= x <= 1. 0 <= ArcCos <= Pi"
' The angle in radians that gives the supplied cosine value 'x', which must obey
' -1 <= x <= 1. The result is in the range 0 <= ArcCos <= Pi.
' Note that 2 * Pi * N + ArcCos and 2 * Pi * N - ArcCos also give x, for any N.
Const ID_c As String = File_c & "ArcCos"
Const Pi As Double = 3.1415926 + 5.358979324E-08
Const PiOvr2 As Double = 1.5707963 + 2.67948965E-08
If Abs(x) < 1# Then
  ArcCos = PiOvr2 - Atn(x / Sqr(1# - x * x))
ElseIf x = 1# Then
  ArcCos = 0#
ElseIf x = -1# Then
  ArcCos = Pi
Else
  Const Invalid_c As Long = 5&  ' "Invalid procedure call or argument"
  Err.Raise Invalid_c, ID_c, _
    "Invalid function argument" & vbLf & _
    "Need -1 <= x <= 1 but x = " & x & vbLf & _
    "Problem in " & ID_c
End If
End Function

'===============================================================================
Public Function ArcCoSec(ByVal x As Double) As Double
Attribute ArcCoSec.VB_Description = "The angle in radians that gives the supplied cosecant (inverse of Sin) value 'x', which must obey x <= -1 or x >= 1. -Pi/2 <= ArcCoSec <= Pi/2"
' The angle in radians that gives the supplied cosecant value 'x', which must
' obey x <= -1 or x >= 1. -Pi/2 <= ArcCoSec <= Pi/2
Const ID_c As String = File_c & "ArcCoSec"
Const PiOvr2 As Double = 1.5707963 + 2.67948965E-08
Dim temp As Double
If x <> 0# Then
  temp = 1# / x
Else  ' x = 0; set to cause error
  temp = 2#  ' an arbitrary value > 1 that will cause an error
End If
' temp is now Sin, so do ArcSin
If Abs(temp) < 1# Then
  ArcCoSec = Atn(temp / Sqr(1# - temp * temp))
ElseIf x = 1# Then
  ArcCoSec = PiOvr2
ElseIf x = -1# Then
  ArcCoSec = -PiOvr2
Else
  Const Invalid_c As Long = 5&  ' "Invalid procedure call or argument"
  Err.Raise Invalid_c, ID_c, _
    "Invalid function argument" & vbLf & _
    "Need x <= -1 or x >= 1 but x = " & x & vbLf & _
    "Problem in " & ID_c
End If
End Function

'===============================================================================
Public Function ArcCosh(ByVal x As Double) As Double
Attribute ArcCosh.VB_Description = "The positive-branch value of the inverse hyperbolic cosine of the input argument 'x', which must obey x >= 1. Has 2E-8 granularity near x = 1"
' The inverse hyperbolic cosine of the input argument. Must have x >= 1.
' This returns the positive branch. Negate the return for the other branch.
' Because ArcCosh behaves as Sqr(2*(x-1)) near x = 1, the smallest return value
' greater than 0 is Sqr(2*2.22E-16) = 2.1E-8. This large granularity can't be
' avoided unless the problem is recast to return ArcCoshXm1(u) where u = x - 1,
' and a routine is written to use a Padé approximation for small u. The writing
' of such a routine is left as an exercise for the student.
Const ID_c As String = File_c & "ArcCosh"
Const Ln_2 As Double = 0.6931471 + 8.055994531E-08
If x > 10000000000# Then  ' -1 makes no change in square root
  ArcCosh = Log(x) + Ln_2  ' less overflow than Log(2# * x)
ElseIf x >= 1# Then
  ArcCosh = Log(x + Sqr(x * x - 1#))  ' use mathematically exact form
Else
  Const Invalid_c As Long = 5&  ' "Invalid procedure call or argument"
  Err.Raise Invalid_c, ID_c, _
    "Invalid function argument" & vbLf & _
    "Need x >= 1 but x = " & x & vbLf & _
    "Problem in " & ID_c
End If
End Function

'===============================================================================
Public Function ArcCoTan(ByVal x As Double) As Double
Attribute ArcCoTan.VB_Description = "The angle in radians that gives the supplied cotangent value 'x'. 0 <= ArcCoTan <= Pi"
' The angle in radians that gives the supplied cotangent. 0 <= ArcCoTan <= Pi
Const PiOvr2 As Double = 1.5707963 + 2.67948965E-08
Dim temp As Double
temp = PiOvr2 - Atn(x)
If temp < 0# Then temp = 0#  ' rare roundoff error for Huge positive argument
ArcCoTan = temp
End Function

'===============================================================================
Public Function ArcSec(ByVal x As Double) As Double
Attribute ArcSec.VB_Description = "The angle in radians that gives the supplied secant (inverse of Cos) value 'x', which must obey x <= -1 or x >= 1. 0 <= ArcSec <= Pi"
' The angle in radians that gives the supplied secant value 'x', which must obey
' x <= -1 or x >= 1. The result is in the range 0 <= ArcSec <= Pi.
Const ID_c As String = File_c & "ArcSec"
Const Pi As Double = 3.1415926 + 5.358979324E-08
Const PiOvr2 As Double = 1.5707963 + 2.67948965E-08
Dim temp As Double
If x <> 0# Then
  temp = 1# / x
Else  ' x = 0; set to cause error
  temp = 2#  ' an arbitrary value > 1 that will cause an error
End If
' temp is now Cos, so do ArcCos
If Abs(temp) < 1# Then
  ArcSec = PiOvr2 - Atn(temp / Sqr(1# - temp * temp))
ElseIf temp = 1# Then
  ArcSec = 0#
ElseIf temp = -1# Then
  ArcSec = Pi
Else
  Const Invalid_c As Long = 5&  ' "Invalid procedure call or argument"
  Err.Raise Invalid_c, ID_c, _
    "Invalid function argument" & vbLf & _
    "Need x <= -1 or x >= 1 but x = " & x & vbLf & _
    "Problem in " & ID_c
End If
End Function

'===============================================================================
Public Function ArcSin(ByVal x As Double) As Double
Attribute ArcSin.VB_Description = "The angle in radians that gives the supplied sine value 'x', which must obey -1 <= x <= 1. -Pi/2 <= ArcSin <= Pi/2"
' The angle in radians that gives the supplied sine value 'x', which must obey
' -1 <= x <= 1. The result is in the range -Pi/2 <= ArcSin <= Pi/2.
Const ID_c As String = File_c & "ArcSin"
Const PiOvr2 As Double = 1.5707963 + 2.67948965E-08
If Abs(x) < 1# Then
  ArcSin = Atn(x / Sqr(1# - x * x))
ElseIf x = 1# Then
  ArcSin = PiOvr2
ElseIf x = -1# Then
  ArcSin = -PiOvr2
Else
  Const Invalid_c As Long = 5&  ' "Invalid procedure call or argument"
  Err.Raise Invalid_c, ID_c, _
    "Invalid function argument" & vbLf & _
    "Need -1 <= x <= 1 but x = " & x & vbLf & _
    "Problem in " & ID_c
End If
End Function

'===============================================================================
Public Function ArcSinh(ByVal x As Double) As Double
Attribute ArcSinh.VB_Description = "Inverse hyperbolic sine of the input argument. Correct even for very small |x|."
' The inverse hyperbolic sine of the input argument. Correct even for very
' small |x|.
Const Ln_2 As Double = 0.6931471 + 8.055994531E-08
Dim absX As Double
absX = Abs(x)
If absX > 10000000000# Then  ' +1 makes no change in square root
  ArcSinh = Sgn(x) * (Log(absX) + Ln_2)  ' less overflow than Log(2# * ax)
ElseIf absX > 0.57 Then  ' roundoff low enough to use mathematically exact form
  ArcSinh = Sgn(x) * Log(absX + Sqr(absX * absX + 1#))
Else  ' use Padé-based form to avoid Log(1# + u) roundoff for small u
  ' worst relative error 4E-16 - see the Maple file "ArcSinhApprox.mws"
  Dim x2 As Double
  x2 = absX * absX
  ArcSinh = x - x * (14.529865069991 + (20.5299661518337 + _
      (7.84722767656954 + 0.68470342125032 * x2) * x2) * x2) * x2 / _
    (87.1791904199951 + (162.410432594456 + (96.8164920782699 + _
      (20.0648834021231 + x2) * x2) * x2) * x2)
End If
End Function

'===============================================================================
Public Function ArcTanh(ByVal x As Double) As Double
Attribute ArcTanh.VB_Description = "Inverse hyperbolic tangent of the input argument value 'x', which must obey -1 <= x <= 1. Correct even for very small |x|."
' The inverse hyperbolic tangent of the input argument value 'x', which must
' obey -1 <= x <= 1. Correct even for very small |x|.
Const ID_c As String = File_c & "ArcTanh"
Const HugeDbl As Double = 1.79769313486231E+308 + 5.7E+293
Dim absX As Double
absX = Abs(x)
If absX < 0.47 Then  ' use a Padé approximation
  Dim x2 As Double
  x2 = x * x
  ' worst relative error 3E-16 - see the Maple file "ArcTanhApprox.mws"
  ArcTanh = x + x * (10.3521776218498 - (16.4426422166074 - (7.3934364920134 - _
    0.82499936759 * x2) * x2) * x2) * x2 / (31.0565328655567 - _
    (67.9618463701179 - (49.647474684131 - (13.4891558533671 - x2) * _
    x2) * x2) * x2)
ElseIf absX < 1# Then  ' use the mathematically exact form
  ArcTanh = 0.5 * Log((1# + x) / (1# - x))
ElseIf x = 1# Then
  ArcTanh = HugeDbl
ElseIf x = -1# Then
  ArcTanh = -HugeDbl
Else
  Const Invalid_c As Long = 5&  ' "Invalid procedure call or argument"
  Err.Raise Invalid_c, ID_c, _
    "Invalid function argument" & vbLf & _
    "Need -1 <= x <= 1 but x = " & x & vbLf & _
    "Problem in " & ID_c
End If
End Function

'===============================================================================
Public Function Atan2( _
  ByVal y As Double, _
  ByVal x As Double, _
  Optional ByVal err00 As Boolean = False) _
As Double
Attribute Atan2.VB_Description = "The ANSI standard arctangent of two arguments in ""reverse"" order.  -Pi <= Atan2 <= Pi  Atan2(0,0) = 0 unless ""err00"" = True"
' The ANSI standard arc-tangent of two arguments in "reverse" order. The branch
' cut is just below the negative X axis, so the result is between -Pi and +Pi.
' The "undefined" value Atan2(0,0) is set to 0 without error unless the Optional
' argument 'err00' is True, in which case an error is raised.
Const ID_c As String = File_c & "Atan2"
Const Pi As Double = 3.1415926 + 5.358979324E-08
Const PiOvr2 As Double = 1.5707963 + 2.67948965E-08
On Error GoTo ErrHandler  ' y / x might overflow (no error on underflow)
If 0# = x Then  ' on the Y axis
  Const Invalid_c As Long = 5&  ' "Invalid procedure call or argument"
  If err00 And (0# = y) Then Err.Raise Invalid_c  ' raise error for 0,0 arg's
  Atan2 = Sgn(y) * PiOvr2  ' also takes care of 0,0 case since Sgn(0) = 0
ElseIf 0# = y Then  ' on the X axis; return 0 if x > 0, Pi if x < 0
  Atan2 = (1# - Sgn(x)) * PiOvr2
ElseIf x > 0# Then  ' in +X half-plane with y <> 0; use ordinary Atn
  Atan2 = Atn(y / x)
Else  ' x < 0; extend smoothly into -X half-plane with y <> 0
  Atan2 = Atn(y / x) + Sgn(y) * Pi  ' gives -Pi just below negative X axis
End If
Exit Function  '----------------------------------------------------------------
ErrHandler:  '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
Err.Raise Err.Number, ID_c, _
  Error$(Err.Number) & " caused by invalid arg('s)" & vbLf & _
  "Input values are: y = " & y & vbLf & _
  "x = " & x & vbLf & _
  "err00 = " & err00 & vbLf & _
  "Problem in " & ID_c
Resume  ' back to error line: 'Set Next Statement' here (Ctrl+F9) & step (F8)
End Function

'===============================================================================
Public Function CoSec(ByVal x As Double) As Double
Attribute CoSec.VB_Description = "The cosecant (inverse of Sin) of the argument 'x' (in radians). Won't overflow."
' The cosecant (inverse of Sin) of the argument 'x' (in radians) . Won't
' overflow.
Const HugeDbl As Double = 1.79769313486231E+308 + 5.7E+293
Dim temp As Double
temp = Sin(x)
If temp <> 0# Then
  CoSec = 1# / temp
Else
  CoSec = HugeDbl  ' barely avoid overflow with really huge value
End If
End Function

'===============================================================================
Public Function Cosh(ByVal x As Double) As Double
Attribute Cosh.VB_Description = "The hyperbolic cosine of the input argument. Cosh >= 1.0"
' The hyperbolic cosine of the input argument. Cosh >= 1.0
Const ID_c As String = File_c & "Cosh"
If Abs(x) > 709.782712893384 Then   ' Exp(x) will overflow or underflow
  Const Invalid_c As Long = 5&  ' "Invalid procedure call or argument"
  Err.Raise Invalid_c, ID_c, _
    "Invalid function argument (Exp overflow)" & vbLf & _
    "Need |x| <= 709.782712893384 but x = " & x & vbLf & _
    "Problem in " & ID_c
End If
Dim temp As Double
temp = Exp(x)
Cosh = 0.5 * (temp + 1# / temp)
End Function

'===============================================================================
Public Function CoTan(ByVal x As Double) As Double
Attribute CoTan.VB_Description = "The cotangent (inverse of Tan) of the argument (in radians). Won't overflow"
' The cotangent (inverse of Tan) of the argument (in radians). Won't overflow.
Const HugeDbl As Double = 1.79769313486231E+308 + 5.7E+293
Dim temp As Double
temp = Tan(x)
If temp <> 0# Then
  CoTan = 1# / temp
Else
  CoTan = HugeDbl  ' barely avoid overflow; return really huge value
End If
End Function

'===============================================================================
Public Function Deg(ByVal angleInRadians As Double) As Double
Attribute Deg.VB_Description = "The input angle (assumed to be in radians) converted to degrees"
' The input angle (assumed to be in radians) converted to degrees.
Const RadToDeg_ As Double = 180# / (3.1415926 + 5.358979324E-08)
Deg = angleInRadians * RadToDeg_
End Function

'===============================================================================
Public Function ExpM1(ByVal x As Double) As Double
Attribute ExpM1.VB_Description = "Does accurate evaluation of Exp(x) - 1 for any |x| < 702.875, even for very small x"
' Does accurate evaluation of Exp(x) - 1 for any x, even for very small x.
' Worst relative error is 3E-15. Meanwhile, Exp(x) - 1 has error going as
' about 9E-17 / x for x > 0, which is bigger when -0.02 < x < 0.03.
' Silently clamps input value to the range 0 <= |x| < 702.875, so the value will
' be wrong for huge (absolute) input values, but no error is raised there, so
' beware. See "Log1P" for a related routine.
Dim absX As Double
absX = Abs(x)
If absX < 1E-16 Then  ' only lead term in power series contributes
  ExpM1 = x
ElseIf absX < 0.028 Then  ' use an equal-ripple-adjusted series expansion
  Const NearOne As Double = 1# + 0.000000000000003  ' 1 + 3E-15
  ' see Maple worksheet "ExpM1 & Log1P Functions_C.mw"
  ExpM1 = (NearOne + (0.499999999999995 + (0.1666666665981 + (0.04166666669 + _
    (0.0083335664 + 0.001388889 * x) * x) * x) * x) * x) * x
Else  ' library math is good enough
  ' if result is close to overflow, clamp at upper end
  ' allow enough margin that result can be multiplied by 1000 before overflow
  Const MaxExpArg_c As Double = 702.8749576144  ' Log(1.79769313485905E+305)
  If Abs(x) > MaxExpArg_c Then x = Sgn(x) * MaxExpArg_c
  ExpM1 = Exp(x) - 1#
End If
End Function

'===============================================================================
Public Function ExpM1P(ByVal x As Double) As Double
Attribute ExpM1P.VB_Description = "Does accurate evaluation of Exp(x) - 1 for 0 <= x <= 702.875, even for very small x"
' Does accurate evaluation of Exp(x) - 1 for x >= 0, even for very small x.
' Worst relative error 5E-15 (near roundoff of Exp(x) - 1 for x < 0.228).
' Silently clamps input value to the range 0 <= x < 702.875, so the value will
' be wrong for negative or huge input values, but no error is raised there -
' beware. See "Log1plusP" for a related routine.
If x <= 0# Then  ' polynomial only good for x >= 0; clamp at zero
  ExpM1P = 0#
ElseIf x < 1E-16 Then  ' only lead term in power series contributes
  ExpM1P = x
ElseIf x < 0.228 Then
  ' Chebyshev polynomial fit found using Maple; worksheet "FrantzNodvik.mws"
  Dim temp As Double  ' need to make expression simpler; split into 2 parts
  temp = (0.04166666780939 + (0.00833331304887 + _
    (0.0013890810153 + (0.0001974165118 + _
    0.00002745392 * x) * x) * x) * x) * x
  ExpM1P = (1# + (0.500000000000273 + (0.166666666636214 + temp) * x) * x) * x
Else  ' library math is good enough
  ' if result is close to overflow, clamp at upper end
  ' allow enough margin that result can be multiplied by 1000 before overflow
  Const MaxExpArg_c As Double = 702.8749576144  ' Log(1.79769313485905E+305)
  If Abs(x) > MaxExpArg_c Then x = Sgn(x) * MaxExpArg_c
  ExpM1P = Exp(x) - 1#
End If
End Function

'===============================================================================
Public Function FunctionsVersion(Optional ByVal trigger As Variant) As String
Attribute FunctionsVersion.VB_Description = "Date of the latest revision to this code, as a string with format ""yyyy-mm-dd"""
' Date of the latest revision to this code, as a string with format "yyyy-mm-dd"
trigger = trigger  ' "use" the input variable
FunctionsVersion = Version_c
End Function

'===============================================================================
Public Function Hypot(ByVal x As Variant, _
  ParamArray otherArgs() As Variant) As Double
Attribute Hypot.VB_Description = "The hypotenuse, AKA root of sum of squares (RSS), AKA Euclidian norm, of the input argument(s), of which there can be 1 or more."
' The hypotenuse, AKA root of sum of squares (RSS), AKA Euclidian norm, of the
' input argument(s), of which there can be 1 or more. Hypot >= 0. Safe (mostly)
' against overflow, except when several sides are very close to the largest
' possible Double. Accepts anything numeric that can be converted to a Double,
' including Strings, so Hypot(3#, "4") = 5 but Hypot(3#, "four") is an error.
' Note that an "Empty" Variant is treated as if it were 0, with no error.
Const ID_c As String = File_c & "Hypot"
Dim argMax As Double, j As Long, relArg As Double, temp As Double, ub As Long, _
  v As Variant
On Error GoTo ErrHandler
ub = UBound(otherArgs)  ' ParamArray indexed from 0 to ub; if empty, ub = -1
' find longest side (in absolute value) and check for invalid arguments
argMax = 0#
For j = -1& To ub
  If j = -1& Then  ' get first arg
    If IsObject(x) Then Err.Raise 7734&  ' not even with a default Property
    v = x
  Else  ' get subsequent arg(s)
    If IsObject(otherArgs(j)) Then Err.Raise 7734&
    v = otherArgs(j)  ' get first or subsequent arg
  End If
  If Not IsNumeric(v) Then Err.Raise 7734&  ' Null, Error, non-numeric String
  temp = Abs(CDbl(v))  ' error here if it can't be converted to a Double
  If argMax < temp Then argMax = temp
Next j
If argMax = 0# Then ' handle special case of all sides = 0; avoid div-by-0
  Hypot = 0#
Else
  ' sum up squares of ratio to largest
  temp = 0#
  For j = -1& To ub
    If j = -1& Then v = x Else v = otherArgs(j)  ' get first or subsequent arg
    relArg = CDbl(v) / argMax  ' |relArg| <= 1
    temp = temp + relArg * relArg
  Next j
  Hypot = argMax * Sqr(temp)  ' temp <= argument count; might overflow
End If
Exit Function  '----------------------------------------------------------------
ErrHandler:  '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
Dim k As Long, s As String
s = toStr(x)  ' make printable versions of special Variant values
For k = 0& To ub
  s = s & "," & toStr(otherArgs(k))
Next k
If 7734& = Err.Number Then  ' this is a special "cannot-convert-to-Double" error
  Const Invalid_c As Long = 5&  ' "Invalid procedure call or argument"
  Err.Raise Invalid_c, ID_c, _
    "Non-numeric input argument " & (j + 2&) & vbLf & _
    "Arguments: " & s & vbLf & _
    "Problem in " & ID_c
Else
  Err.Raise Err.Number, ID_c, _
    Error$(Err.Number) & " caused by invalid function argument(s)" & vbLf & _
    "Arguments: " & s & vbLf & _
    "Problem in " & ID_c
End If
Resume  ' back to error line: 'Set Next Statement' here (Ctrl+F9) & step (F8)
End Function

'===============================================================================
Function IntSmooth( _
  ByVal x As Double, _
  Optional ByVal wide As Double = 0.001) _
As Double
Attribute IntSmooth.VB_Description = "Smoothed, continuous approximation to the Int function, which equals zero at X = 0, and steps upward by unity around x = N+1/2 (N integer). The ""wide"" parameter sets the X distance over which the smoothed unit jump extends."
' Smoothed, continuous approximation to VB's Int(x + 0.5) function, which equals
' zero at X = 0, and steps upward by unity around X = N + 1/2 (N integer). The
' "wide" parameter sets the X distance over which the smoothed jump extends,
' so the smaller "wide" is the more the function resembles an abrupt step.
' See the Function "StepSqr" for a single smooth step near x = 0.
Const Pi As Double = 3.1415926 + 5.358979324E-08

' impose sanity on width of transition
wide = Abs(wide)
Const MinWide_c As Double = 0.00000001
If wide < MinWide_c Then wide = MinWide_c Else If wide > 1# Then wide = 1#

' We get in trouble when Atn() and Int() both try to jump at the same point,
' since errors in Pi et. al. will make the jumps at slightly different points.
' Therefore we make two versions with the jump points offset by 1/2, and switch
' between them along 45 degree diagonals through the bottom and top of the jumps
' at X = N + 1/2 (N integer). Thus we never cross a double-jump point.
Dim dx As Double
dx = x - 0.5 - Int(x)  ' distance from point where Int() jumps
' the dx limit below gives equal distance to the two forms from X = 0 & X = 1/2
' it is a fast 2%-error approx. to the exact value Atn(Sqr(wide)) / Pi
If Abs(dx) > 0.31831 * Sqr(wide) * (1# - 0.229268 * wide) Then
  ' this approximation form is centered at X = N (N integer)
  IntSmooth = Atn(wide * Tan(x * Pi)) / Pi + Int(x + 0.5)
Else
  ' this approximation form is centered at X = N + 1/2 (N integer)
  IntSmooth = 0.5 + Atn(Tan(Pi * (x - 0.5)) / wide) / Pi + Int(x)
End If
End Function

'===============================================================================
Public Function InvPos( _
  ByVal x As Double, _
  Optional ByVal delta As Double = 1#) _
As Double
Attribute InvPos.VB_Description = "Equal to 1/x for x >= delta > 0, but then continues to increase steadily as x decreases below delta, finally becoming asymptotic to 2/delta for large negative x"
' Equal to 1/x for x >= delta, but then continues to increase steadily as x
' decreases below delta, finally becoming asymptotic to 2/delta for large
' negative x. Note you must have delta > 0. Use this when you don't want the
' sign flip and discontinuity that happen with 1/x as x passes through zero,
' don't mind that the result is inaccurate, but do want the result to keep
' increasing as x becomes smaller than delta, passes through zero, and
' continues on to negative values. Note that value and slope are continuous
' at x = delta, but higher derivatives are discontinuous.
' See 'NoZero' for an alternative.
Const ID_c As String = File_c & "InvPos"
If delta <= 0# Then
  Const Invalid_c As Long = 5&  ' "Invalid procedure call or argument"
  Err.Raise Invalid_c, ID_c, _
    "Invalid function argument" & vbLf & _
    "Need delta > 0 but delta = " & delta & vbLf & _
    "Problem in " & ID_c
ElseIf x >= delta Then  ' return the exact inverse
  InvPos = 1# / x
Else  ' return a fake continuation, forcing strictly decreasing behavior
  InvPos = 2# / delta + 1# / (x - 2# * delta)
End If
End Function

'===============================================================================
Public Function LimitExp( _
  ByVal x As Double, _
  ByVal xLo As Double, _
  ByVal xHi As Double, _
  Optional ByVal wide As Double = 0.0001) _
As Double
Attribute LimitExp.VB_Description = "Smoothly limit x between xLo and xHi, with a transition width near the value given by wide * (xHi - xLo)"
' Smoothly limit x between xLo and xHi, with a transition width near the
' value given by wide * (xHi - xLo). Use an exponential-based method, which
' will rapidly approach x as x moves away from the limits. Use this to apply
' simple lower and upper ("box") limits to input argument values. First, use
' limitExp(x, xLo, xHi) instead of x in the figure-of-merit coding. Then, run
' the optimizer, but make sure the initial argument values lie within the given
' limits. When the optimizer is done use limitExp again, with the same limits,
' to convert the returned x to its limited value. For one-sided limits, just set
' the other side well away from the value you want to limit at (but not too far
' away, to avoid roundoff problems).
wide = Abs(wide)
If wide > 10# Then wide = 10#  ' clip absurd values
If wide < 0.00000001 Then wide = 0.00000001
Dim eps As Double, uLo As Double, uHi As Double
eps = Abs(wide * (xHi - xLo))
Const MinEps_c As Double = 0.00000001
If eps < MinEps_c Then eps = MinEps_c  ' we are about to divide by eps
uLo = (x - xLo) / eps  ' use z as approx. to Log(1 + Exp(z)) for large z
Const Log_1E8_c As Double = 18.4206807439524
' note that Log(1 + eps) = eps - eps^2 / 2 + eps^3 / 3 +- ...
If uLo < -Log_1E8_c Then  ' Exp(uLo) less than 1E-8, next term < 5E-17
  uLo = Exp(uLo)  ' approx. Log(1 + Exp(x)) as Exp(x)
ElseIf uLo <= Log_1E8_c Then  ' Exp(uLo) will be less than 1E8
  uLo = Log(1# + Exp(uLo))  ' no need to approximate; use exact result
End If
uHi = (xHi - x) / eps
If uHi < -Log_1E8_c Then
  uHi = Exp(uHi)
ElseIf uHi <= Log_1E8_c Then
  uHi = Log(1# + Exp(uHi))
End If
LimitExp = xLo + xHi - x + eps * (uLo - uHi)
End Function

'===============================================================================
Public Function LimitHard( _
  ByVal x As Double, _
  ByVal xLo As Double, _
  ByVal xHi As Double) _
As Double
Attribute LimitHard.VB_Description = "Abruptly limit x between xLo and xHi, with no smooth transition"
' Abruptly limit x between xLo and xHi, with no smooth transition. For one-sided
' limits, just set the other side well away from the value you want to limit, or
' use the Max and Min functions.
If x < xLo Then
  LimitHard = xLo
ElseIf x > xHi Then
  LimitHard = xHi
Else
  LimitHard = x
End If
End Function

'===============================================================================
Public Function LimitSqr( _
  ByVal x As Double, _
  ByVal xLo As Double, _
  ByVal xHi As Double, _
  Optional ByVal wide As Double = 0.0001) _
As Double
Attribute LimitSqr.VB_Description = "Smoothly limit x between xLo and xHi, with a transition width near the value given by wide * (xHi - xLo)"
' Smoothly limit x between xLo and xHi, with a transition width near the
' value given by wide * (xHi - xLo). Use a square-root-based method, which
' will more slowly approach x as x moves away from the limits. Use this to apply
' simple lower and upper ("box") limits to argument values in NMreduce. First,
' use limitSqr(x, xLo, xHi) instead of x in the coding of NMfunc. Then, run
' NMreduce. Make sure the initial argument values lie within the supplied
' limits. When NMreduce is done use limitSqr again, with the same limits, to
' convert the returned x to its limited value. For one-sided limits, just set
' the other side well away from the value you want to limit at (but not too far
' away, to avoid roundoff problems).
wide = Abs(wide)
If wide > 10# Then wide = 10#  ' clip absurd values
If wide < 0.00000001 Then wide = 0.00000001
Dim dLo As Double, dHi As Double, eps As Double, eps2 As Double
dLo = x - xLo
dHi = xHi - x
eps = wide * (xHi - xLo)
eps2 = eps * eps
LimitSqr = 0.5 * (xLo + xHi + Sqr(dLo * dLo + eps2) - Sqr(dHi * dHi + eps2))
End Function

'===============================================================================
Public Function Log10(ByVal x As Double) As Double
Attribute Log10.VB_Description = "The logarithm, base ten, of the input argument 'x', which must obey x > 0"
' The logarithm, base ten, of the input argument 'x', which must obey x > 0.
Const ID_c As String = File_c & "Log10"
Const Log10_e As Double = 0.43429448 + 1.903251828E-09
If x > 0# Then
  Log10 = Log(x) * Log10_e
Else
  Const Invalid_c As Long = 5&  ' "Invalid procedure call or argument"
  Err.Raise Invalid_c, ID_c, _
    "Invalid function argument" & vbLf & _
    "Need x > 0 but x = " & x & vbLf & _
    "Problem in " & ID_c
End If
End Function

'===============================================================================
Public Function Log1P(ByVal x As Double) As Double
Attribute Log1P.VB_Description = "Does accurate evaluation of Log(1 + x) for any x, even for very small x"
' Does accurate evaluation of Log(1 + x) for any x, even for very small x.
' Worst relative error is 3E-15, except near x = 0.025 where it briefly rises to
' around 4.3E-15. Meanwhile, Log(1 + x) has roundoff error going as about
' 1.1E-16 / x for x > 0, which is bigger when -0.013 < x < 0.025.
Dim absX As Double
absX = Abs(x)
If absX < 5.5E-17 Then  ' only lead term in power series contributes
  Log1P = x
ElseIf absX < 0.0253 Then  ' use an equal-ripple-adjusted Padé [2,3] form
  ' see Maple worksheet "ExpM1 & Log1P Functions_C.mw"
  Const C1 As Double = 19.9927673651397 + 0.000000000000055  ' needs 16+ digits
  Log1P = (C1 + (19.9951779061522 + 3.6664657191017 * x) * x) * x / _
    (19.9927673651397 + (29.991561588723 + (11.99799072665 + x) * x) * x)
Else  ' library math is good enough
  Log1P = Log(1# + x)
End If
End Function

'===============================================================================
Public Function Log1plusP(ByVal x As Double) As Double
Attribute Log1plusP.VB_Description = "Does accurate evaluation of Log(1 + x) for x >= 0, even for very small x"
' Does accurate evaluation of Log(1 + x) for x >= 0, even for very small x.
' Worst relative error 6E-15 (near roundoff of Log(1 + x) for x < 0.0693).
' Silently clamps input value to the range 0 < x, so the value will be wrong
' for negative input values, but no error is raised there - beware.
' See "ExpM1P" for a related routine.
If x <= 0# Then  ' polynomial only good for x >= 0; clamp at zero
  Log1plusP = 0#
ElseIf x < 0.0693 Then
  ' Chebyshev polynomial fit found using Maple; worksheet "FrantzNodvik.mws"
  Dim temp As Double  ' need to make expression simpler; split into 2 parts
  temp = 0.19999489114 - (0.16650460939 - (0.140003324 - 0.0982997 * x) * x) * x
  Log1plusP = (1# - (0.499999999998088 - (0.333333332635105 - _
    (0.2499999134106 - temp * x) * x) * x) * x) * x
Else  ' library math is good enough
  Log1plusP = Log(1# + x)
End If
End Function

'===============================================================================
Public Function LogN(ByVal x As Double, ByVal N As Double) As Double
Attribute LogN.VB_Description = "The logarithm, base N, of the argument x. Both x and N must be > 0"
' The logarithm, base N, of the input argument x. Both x and N must be > 0.
' If N = 10, use Log10 instead of this function, since it's faster.
Const ID_c As String = File_c & "LogN"
If (x > 0#) And (N > 0#) Then
  LogN = Log(x) / Log(N)
Else
  Const Invalid_c As Long = 5&  ' "Invalid procedure call or argument"
  Err.Raise Invalid_c, ID_c, _
    "Invalid function argument(s)" & vbLf & _
    "Need x > 0 and N > 0 but x = " & x & vbLf & _
    "N = " & N & vbLf & _
    "Problem in " & ID_c
End If
End Function

'===============================================================================
Public Function Max(ByVal x1 As Variant, ByVal x2 As Variant, _
  ParamArray otherArgs() As Variant) As Variant
Attribute Max.VB_Description = "Maximum of two or more quantities. Works with any type(s) where ""<"" works."
' Maximum of two or more quantities. Works with any type(s) where "<" works.
' Accepts anything that can be converted to a Double, including Strings.
' Note that an "Empty" Variant is treated as if it were 0, with no error.
' Be careful when mixing types; for example Max("hello", 0.25) = "hello"
Const ID_c As String = File_c & "Max"
Dim res As Variant  ' result
Dim ub As Long
ub = UBound(otherArgs)  ' will have ub = -1 if otherArgs is empty
On Error GoTo ErrHandler
Const Invalid_c As Long = 5&  ' "Invalid procedure call or argument"
If IsObject(x1) Or IsObject(x2) Then Err.Raise Invalid_c
If x1 < x2 Then res = x2 Else res = x1
' check for more than 2 arguments; always have LBound = 0 for ParamArrays
Dim j As Long
For j = 0& To ub
  If IsObject(otherArgs(j)) Then Err.Raise Invalid_c
  If res < otherArgs(j) Then res = otherArgs(j)
Next j
Max = res
Exit Function  '----------------------------------------------------------------
ErrHandler:  '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
Dim k As Long, s As String
s = toStr(x1) & "," & toStr(x2)  ' printable versions of special Variant values
For k = 0& To ub
  s = s & "," & toStr(otherArgs(k))
Next k
Err.Raise Err.Number, ID_c, _
  Error$(Err.Number) & " caused by invalid function arg(s)" & vbLf & _
  "Arguments: " & s & vbLf & _
  "Problem in " & ID_c
Resume  ' back to error line: 'Set Next Statement' here (Ctrl+F9) & step (F8)
End Function

'===============================================================================
Public Function MaxExp( _
  ByVal arg1 As Double, _
  ByVal arg2 As Double, _
  Optional ByVal wide As Double = 0.01) _
As Double
Attribute MaxExp.VB_Description = "Smoothed maximum of two quantities, with absolute error from exact Max falling off as Exp(-x) away from the point(s) of equality"
' Smoothed maximum of two quantities, with absolute error from exact Max falling
' off as Exp(-x) away from the point(s) of equality. See LimitExp & LimitSqr.
wide = Abs(wide)
Const MinWide_c As Double = 0.0000000001
If wide < MinWide_c Then wide = MinWide_c  ' impose minimal sanity
Dim t As Double
t = (arg1 - arg2) / wide
Const MaxExpArg_c As Double = 702.8749576144  ' Log(1.79769313485905E+305)
If Abs(t) > MaxExpArg_c Then t = Sgn(t) * MaxExpArg_c
t = Exp(t)
MaxExp = 0.5 * (arg1 + arg2 + Log(2# + t + 1# / t))  ' Log(2 + 2 * Cosh(t))
End Function

'===============================================================================
Public Function MaxSqr( _
  ByVal arg1 As Double, _
  ByVal arg2 As Double, _
  Optional ByVal wide As Double = 0.001) _
As Double
Attribute MaxSqr.VB_Description = "Smoothed maximum of two quantities, with absolute error from exact Max falling off as 1/x away from the point(s) of equality"
' Smoothed maximum of two quantities, with absolute error from exact Max falling
' off as 1/x away from the point(s) of equality. Also see LimitExp & LimitSqr.
wide = Abs(wide)
Const MinWide_c As Double = 0.0000000001
If wide < MinWide_c Then wide = MinWide_c  ' impose minimal sanity
Dim t As Double
t = arg1 - arg2
' the constant here is 1 / ( 2^(2/3) - 1 ) (see "FudgeGallery.mws" for why)
' In this form, there can be overflow with arguments above 1E154 or so
MaxSqr = 0.5 * (arg1 + arg2 + Sqr(t * t + 1.702414384 * wide * wide))
End Function

'===============================================================================
Public Function Min(ByVal x1 As Variant, ByVal x2 As Variant, _
  ParamArray otherArgs() As Variant) As Variant
Attribute Min.VB_Description = "Minimum of two or more quantities. Works with any type(s) where "">"" works."
' Minimum of two or more quantities. Works with any type(s) where ">" works.
' Accepts anything that can be converted to a Double, including Strings.
' Note that an "Empty" Variant is treated as if it were 0, with no error.
' Be careful when mixing types; for example Min("hello", 0.25) = 0.25
Const ID_c As String = File_c & "Min"
Dim res As Variant  ' result
Dim ub As Long
ub = UBound(otherArgs)  ' will have ub = -1 if otherArgs is empty
On Error GoTo ErrHandler
Const Invalid_c As Long = 5&  ' "Invalid procedure call or argument"
If IsObject(x1) Or IsObject(x2) Then Err.Raise Invalid_c
If x1 > x2 Then res = x2 Else res = x1
' check for more than 2 arguments; always have LBound = 0 for ParamArrays
Dim j As Long
For j = 0& To ub
  If IsObject(otherArgs(j)) Then Err.Raise Invalid_c
  If res > otherArgs(j) Then res = otherArgs(j)
Next j
Min = res
Exit Function  '----------------------------------------------------------------
ErrHandler:  '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
Dim k As Long, s As String
s = toStr(x1) & "," & toStr(x2)  ' printable versions of special Variant values
For k = 0& To ub
  s = s & "," & toStr(otherArgs(k))
Next k
Err.Raise Err.Number, ID_c, _
  Error$(Err.Number) & " caused by invalid function arg(s)" & vbLf & _
  "Arguments: " & s & vbLf & _
  "Problem in " & ID_c
Resume  ' back to error line: 'Set Next Statement' here (Ctrl+F9) & step (F8)
End Function

'===============================================================================
Public Function MinExp( _
  ByVal arg1 As Double, _
  ByVal arg2 As Double, _
  Optional ByVal wide As Double = 0.01) _
As Double
Attribute MinExp.VB_Description = "Smoothed minimum of two quantities, with absolute error from exact Min falling off as Exp(-x) away from the point(s) of equality"
' Smoothed minimum of two quantities, with absolute error from exact Min falling
' off as Exp(-x) away from the point(s) of equality. See LimitExp & LimitSqr.
Const MinWide_c As Double = 0.0000000001
If wide < MinWide_c Then wide = MinWide_c  ' impose minimal sanity
Dim t As Double
t = (arg1 - arg2) / wide
Const MaxExpArg_c As Double = 702.8749576144  ' Log(1.79769313485905E+305)
If Abs(t) > MaxExpArg_c Then t = Sgn(t) * MaxExpArg_c
t = Exp(t)
MinExp = 0.5 * (arg1 + arg2 - Log(2# + t + 1# / t))  ' Log(2 + 2 * Cosh(t))
End Function

'===============================================================================
Public Function MinSqr( _
  ByVal arg1 As Double, _
  ByVal arg2 As Double, _
  Optional ByVal wide As Double = 0.001) _
As Double
Attribute MinSqr.VB_Description = "Smoothed minimum of two quantities, with absolute error from exact Min falling off as 1/x away from the point(s) of equality"
' Smoothed minimum of two quantities, with absolute error from exact Min falling
' off as 1/x away from the point(s) of equality. Also see LimitExp & LimitSqr.
wide = Abs(wide)
Const MinWide_c As Double = 0.0000000001
If wide < MinWide_c Then wide = MinWide_c  ' impose minimal sanity
Dim t As Double
t = arg1 - arg2
' the constant here is 1 / ( 2^(2/3) - 1 ) (see "FudgeGallery.mws" for why)
' In this form, there can be overflow with arguments above 1E154 or so
MinSqr = 0.5 * (arg1 + arg2 - Sqr(t * t + 1.702414384 * wide * wide))
End Function

'===============================================================================
Public Function NoZero(ByVal x As Double) As Double
Attribute NoZero.VB_Description = "Value that can be safely reciprocated, without division by zero"
' Value that can be safely reciprocated, without division by zero
' See 'InvPos' for an alternative.
Const MinValue_c As Double = 1E-300  ' allow some space for later arithmetic
If 0# = x Then
  NoZero = MinValue_c  ' special case, because Sgn(0) = 0
ElseIf Abs(x) < MinValue_c Then
  NoZero = Sgn(x) * MinValue_c  ' hard limit at minimum; keep sign
Else
  NoZero = x  ' unchanged; reciprocation is safe
End If
End Function

'===============================================================================
Public Function Pow2(ByVal x As Double) As Double
Attribute Pow2.VB_Description = "The square of the supplied argument 'x', which must obey |x| < 1.34078079299426E154"
' The square of the supplied argument 'x', which must obey
' |x| < 1.34078079299426E154. Unless 'x' is something like a complicated
' expression or function return, it's probably better to just use
' x * x inline. Will silently underflow to 0 if |x| < 1.572E-162
Const ID_c As String = File_c & "Pow2"
On Error GoTo ErrHandler  ' might overflow
Pow2 = x * x
Exit Function  '----------------------------------------------------------------
ErrHandler:  '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
Err.Raise Err.Number, ID_c, _
  Error$(Err.Number) & " caused by invalid function argument" & vbLf & _
  "Need |x| < 1.34078079299426E154 but" & vbLf & _
  "x = " & x & vbLf & _
  "Problem in " & ID_c
Resume  ' back to error line: 'Set Next Statement' here (Ctrl+F9) & step (F8)
End Function

'===============================================================================
Public Function Pow3(ByVal x As Double) As Double
Attribute Pow3.VB_Description = "The third power of the supplied argument 'x', which must obey |x| < 5.64380309412237E102"
' The third power of the supplied argument 'x', which must obey
' |x| < 5.64380309412237E102. Unless 'x' is something like a complicated
' expression or function return, it's probably better to just use
' x * x * x inline. Will silently underflow to 0 if |x| < 1.352E-108
Const ID_c As String = File_c & "Pow3"
On Error GoTo ErrHandler  ' might overflow
Pow3 = x * x * x
Exit Function  '----------------------------------------------------------------
ErrHandler:  '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
Err.Raise Err.Number, ID_c, _
  Error$(Err.Number) & " caused by invalid function argument" & vbLf & _
  "Need |x| < 5.64380309412237E102 but" & vbLf & _
  "x = " & x & vbLf & _
  "Problem in " & ID_c
Resume  ' back to error line: 'Set Next Statement' here (Ctrl+F9) & step (F8)
End Function

'===============================================================================
Public Function Pow4(ByVal x As Double) As Double
Attribute Pow4.VB_Description = "The fourth power of the supplied argument 'x', which must obey |x| < 1.15792089237317E77"
' The fourth power of the supplied argument 'x', which must obey
' |x| < 1.15792089237317E77. Unless 'x' is something like a complicated
' expression or function return, it's probably better to just use
' x * x * x * x inline, or y = x * x: z = y * y. Will silently underflow to 0
' if |x| < 1.254E-81
Const ID_c As String = File_c & "Pow4"
On Error GoTo ErrHandler  ' might overflow
Dim temp As Double
temp = x * x
Pow4 = temp * temp
Exit Function  '----------------------------------------------------------------
ErrHandler:  '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
Err.Raise Err.Number, ID_c, _
  Error$(Err.Number) & " caused by invalid function argument" & vbLf & _
  "Need |x| < 1.15792089237317E77 but" & vbLf & _
  "x = " & x & vbLf & _
  "Problem in " & ID_c
Resume  ' back to error line: 'Set Next Statement' here (Ctrl+F9) & step (F8)
End Function

'===============================================================================
Public Function Pow5(ByVal x As Double) As Double
Attribute Pow5.VB_Description = "The fifth power of the supplied argument 'x', which must obey |x| < 4.47654662275724E61"
' The fifth power of the supplied argument 'x', which must obey
' |x| < 4.47654662275724E61. Unless 'x' is something like a complicated
' expression or function return, it's probably better to just use
' x * x * x * x * x inline, or y = x * x: z = x * y * y. Will silently
' underflow to 0 if |x| < 1.900E-65
Const ID_c As String = File_c & "Pow5"
On Error GoTo ErrHandler  ' might overflow
Dim temp As Double
temp = x * x
Pow5 = x * temp * temp
Exit Function  '----------------------------------------------------------------
ErrHandler:  '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
Err.Raise Err.Number, ID_c, _
  Error$(Err.Number) & " caused by invalid function argument" & vbLf & _
  "Need |x| < 4.47654662275724E61 but" & vbLf & _
  "x = " & x & vbLf & _
  "Problem in " & ID_c
Resume  ' back to error line: 'Set Next Statement' here (Ctrl+F9) & step (F8)
End Function

'===============================================================================
Public Function Rad(ByVal angleInDegrees As Double) As Double
Attribute Rad.VB_Description = "The input angle (assumed to be in degrees) converted to radians"
' The input angle (assumed to be in degrees) converted to radians.
Const DegToRad_ As Double = (3.1415926 + 5.358979324E-08) / 180#
Rad = angleInDegrees * DegToRad_
End Function

'===============================================================================
Public Function RankMedian(ByVal rank As Long, ByVal outOf As Long) As Double
Attribute RankMedian.VB_Description = "Unbiased estimate of the rank (AKA quantile or order) median value corresponding to the Jth of N IID randoms taken from an arbitrary distribution, sorted into increasing order"
' The unbiased estimate of the rank (AKA quantile or order) median value
' corresponding to the Jth of N independent, identically-distributed random
' values taken from an arbitrary distribution, and then sorted into increasing
' order. The exact median value is given by equating the binomial distribution
' of the rank, which is equal to an incomplete beta function, to 1/2 (see, e.g.,
' "A Reliable Algorithm for the Exact Median Rank Function", Jacquelin, 1993).
' Here we are using Filliben's approximation (Filliben, "The Probability Plot
' Correlation Coefficient Test for Normality," Tecnometrics, 17, 1, 1975) which
' is good to 0.00028 for any J ond N. See also Wikipedia, "Q-Q Plot."
' This can be used with any inverse cumulative distribution to supply the
' median sample points for stratified sampling as someIC(RankMedian(J, N)).
' Note that using the median as an estimate minimizes the absolute error, just
' as using the mean minimizes the square of the error.
' Note that for the uniform distribution from 0 to 1, we have the exact results
' mean = J / (N + 1) and mode = (J - 1) / (N - 1). These may cautiously be used
' with any inverse cumulative distribution function. However, be careful using
' the mode with long-tailed distributions; that gives "infinity" at J=1 & J=N.
'
' Exact: (1,5)=0.129449 (2,5)=0.313810 (3,5)=0.5 (4,5)=0.686189 (5,5)=0.870551
' Here:  (1,5)=0.129449 (2,5)=0.313607 (3,5)=0.5 (4,5)=0.686393 (5,5)=0.870551
Const ID_c As String = File_c & "RankMedian"
Const Filliben As Double = 0.3175
Const Filliben2 As Double = 1# - 2# * Filliben
If (rank < 1&) Or (rank > outOf) Then
  Const Invalid_c As Long = 5&  ' "Invalid procedure call or argument"
  Err.Raise Invalid_c, ID_c, _
    "Invalid function argument" & vbLf & _
    "Need 1 <= rank <= outOf but rank = " & rank & _
      " outOf = " & outOf & vbLf & _
    "Problem in " & ID_c
End If
If rank = 1& Then
  RankMedian = 1# - 0.5 ^ (1# / outOf)  ' exact median for J = 1
ElseIf rank = outOf Then
  RankMedian = 0.5 ^ (1# / outOf)  ' exact median for J = N
Else  ' approximation is symmetric around center, where value = 1/2
  RankMedian = (rank - Filliben) / (outOf + Filliben2)
End If
End Function

'===============================================================================
Public Function Sec(ByVal x As Double) As Double
Attribute Sec.VB_Description = "The secant (inverse of Cos) of the argument (in radians). Won't overflow."
' The secant (inverse of Cos) of the argument (in radians). Won't overflow.
Const HugeDbl As Double = 1.79769313486231E+308 + 5.7E+293
Dim temp As Double
temp = Cos(x)
If temp <> 0# Then
  Sec = 1# / temp
Else
  Sec = HugeDbl  ' barely avoid overflow with really huge value
End If
End Function

'===============================================================================
Public Function Sech(ByVal x As Double) As Double
Attribute Sech.VB_Description = "The hyperbolic secant (inverse of Cosh) of the argument. 0 < Sech <= 1."
' The hyperbolic secant (inverse of Cosh) of the argument. 0 < Sech <= 1.
Const ID_c As String = File_c & "Sech"
If Abs(x) > 709.782712893384 Then   ' Exp(x) will overflow or underflow
  Const Invalid_c As Long = 5&  ' "Invalid procedure call or argument"
  Err.Raise Invalid_c, ID_c, _
    "Invalid function argument (Exp overflow)" & vbLf & _
    "Need |x| <= 709.782712893384 but x = " & x & vbLf & _
    "Problem in " & ID_c
End If
Dim temp As Double
temp = Exp(x)
Sech = 2# / (temp + 1# / temp)
End Function

'===============================================================================
Public Function SerialNum(Optional ByVal setTo As Variant) As Double
Attribute SerialNum.VB_Description = "Each call yields a Double that is 1 larger than the previous value, starting with 0. Supply optional argument to (re)start at that value."
' Supplies a sequence of increasing integer values, in a Double. On each call,
' the return is one more than on the previous call. First value returned is 0.
' After 9.007199254740992E+15 values, the sequence quits with an error. If the
' optional argument is supplied, the sequence is (re)set to the supplied value
' and that value is returned.
Const ID_c As String = File_c & "SerialNum"
Static id_s As Double   ' number to be issued; initialized to 0 on program start
Const maxCount_c As Double = 9.00719925474099E+15 + 2#  ' no unit add above here
If Not IsMissing(setTo) Then  ' user has supplied a starting value
  On Error Resume Next  ' ignore silly values; they will give 0
  id_s = CDbl(Int(setTo + 0.5))  ' set to closest integer value
  On Error GoTo 0
  If -maxCount_c > id_s Then id_s = -maxCount_c
  If maxCount_c - 1# < id_s Then id_s = maxCount_c - 1#  ' allow 1 return val.
End If
SerialNum = id_s  ' set to return the present value
If maxCount_c > id_s Then  ' increment Static variable for the next call
  id_s = id_s + 1#  ' stops adding at 9,007,199,254,740,992
Else
  Const Overflow_c As Long = 6&  ' Overflow
  Err.Raise Overflow_c, ID_c, _
    Error$(Overflow_c) & " incrementing value for next call" & vbLf & _
    "New value >= maximum value of " & maxCount_c & vbLf & _
    "Try starting with SerialNum(-" & maxCount_c & ")" & vbLf & _
    "to get twice as many values." & vbLf & _
    "Problem in " & ID_c
End If
End Function

'===============================================================================
Public Function Sinh(ByVal x As Double) As Double
Attribute Sinh.VB_Description = "The hyperbolic sine of the input argument. Correct even for very small |x|."
' The hyperbolic sine of the input argument.
Const ID_c As String = File_c & "Sinh"
If Abs(x) > 709.782712893384 Then   ' Exp(x) will overflow or underflow
  Const Invalid_c As Long = 5&  ' "Invalid procedure call or argument"
  Err.Raise Invalid_c, ID_c, _
    "Invalid function argument (Exp overflow)" & vbLf & _
    "Need |x| <= 709.782712893384 but x = " & x & vbLf & _
    "Problem in " & ID_c
End If
If Abs(x) < 0.639 Then  ' need to use Padé form to keep roundoff low
  Dim x2 As Double
  x2 = x * x
  ' worst relative error 3E-16 - see the Maple file "SinhApprox.mws"
  Sinh = x + x * (18.3930586940839 + (0.7529862681128 + (0.0135631642778 + _
    0.0001057093995 * x2) * x2) * x2) * x2 / (110.358352164517 - x2)
Else  ' minimal cancellation error in Exp(x) - Exp(-x)
  Dim temp As Double
  temp = Exp(x)
  Sinh = 0.5 * (temp - 1# / temp)
End If
End Function

'===============================================================================
Public Function StdNormCum(ByVal z As Double) As Double
Attribute StdNormCum.VB_Description = "Cumulative Distribution Function of a standard normal variate to 1E-14 worst relative error within 6 sigma of the mean. Gives 0 for z < -37.5 & 1 for z > 8.2923."
' Returns cumulative distribution function (probability) of a standard normal
' variate z with a relative error in the probability of 1E-14 or better within 6
' sigma of the mean (i.e., for -6 < z < 6). The relative error rises to 3E-13
' far out on the negative tail, where the probability is very small, because of
' roundoff in Exp(x). If z < -37.5, 0.0 is returned, due to approaching double-
' precision underflow (note that StdNormCum( -37.5) = 4.6E-308). For z > 8.2923,
' the returned value is exactly 1.0 because of roundoff in the subtraction.
' For large positive z, the difference from unity is small and suffers from
' roundoff error. If you want this difference to high accuracy, note that
' 1 - StdNormCum(z) = StdNormCum(-z). The quantity 1 - StdNormCum(z) (the
' exceedance) is small and accurate just where StdNormCum(z) gets into trouble.
'
' The CDF of a normal variate with non-zero mean and non-unit standard deviation
' is given by "StdNormCum((x - mean) / stdDev)".
' Note: CDF of the max of N independent Normal samples is this CDF to the Nth power.
' Version of 2008-11-08 - see Maple worksheet "CumulativeNormalApproxI.mws"
Dim a As Double
a = Abs(z)
Dim x2 As Double
x2 = z * z
Dim r As Double, s As Double, u As Double
If a < 0.786 Then
  r = 0.5 + z * (286274.853432297 + (13766.2029201158 + _
    2102.12025748516 * x2) * x2) / (717584.641929165 + _
    (154104.193799632 + (13013.6502555081 + (452.005967245251 + _
    x2) * x2) * x2) * x2)
ElseIf z > 8.3 Then  ' limited by granularity in floating point values near 1
  r = 1#
ElseIf z < -37.5 Then  ' limited by smallest possible floating point value
  r = 0#  ' huge relative error, but there's no choice
Else
  If a < 1.2633 Then
    r = (18.8692816870042 + (11.4479560860487 + (3.37052801140488 + _
      0.39779995873049 * a) * a) * a) / (37.7385790222798 + _
      (53.0067910675199 + (30.1657284206649 + (8.39621980769258 + _
      a) * a) * a) * a) + 1.6E-16
  ElseIf a < 1.8395 Then
    r = (15.3105344115503 + (9.99779797787253 + (3.08878063108747 + _
      0.39848536667058 * a) * a) * a) / (30.6212995567237 + _
      (44.4262834362372 + (26.318962736101 + (7.71750668140928 + _
      a) * a) * a) * a)
  Else
    u = 1# / x2
    ' split evaluation into parts to avoid "expression too complex"
    s = 3.16965563545039E-04 + (0.015227793862574 + (0.260705053546659 + _
      (2.00621176176922 + (7.22108903104585 + (11.5148288549791 + _
      (6.89233938349171 + u) * u) * u) * u) * u) * u) * u
    r = 0.712684696150189 + (2.31460584415985 + (3.07234263767005 + _
      (1.28623813814016 + 7.21836636375586E-02 * u) * u) * u) * u
    r = ((1.26450964729363E-04 + (5.94855984431045E-03 + _
      (9.83106106510839E-02 + r * u) * u) * u) / s + 4E-16) / a
  End If
  r = r * Exp(-0.5 * x2)  ' Exp(-0.5 * x2) = 0 if |x| >= 38.603; can't happen
  If z > 0# Then r = 1# - r
End If
StdNormCum = r
End Function

'===============================================================================
Public Function StdNormInvCum(ByVal prob As Double) As Double
Attribute StdNormInvCum.VB_Description = "Inverse cumulative distribution function (ICDF) of a standard normal distribution (Gaussian with zero mean and unit standard deviation). Almost the exact inverse of StdNormCum(z)."
' Inverse cumulative distribution function (ICDF) of a standard normal
' distribution (Gaussian with zero mean and unit standard deviation). Note
' that values more than +- 38.4674 will never be returned. The probability
' of such extreme values is less than 2E-16. To translate and scale the
' result to a desired non-zero mean and non-unit standard deviation, use the
' form "mean + stdDev * StdNormInvCum(prob)".
Const ID_c As String = File_c & "StdNormInvCum"
' check the input probability
If (prob < 0#) Or (prob > 1#) Then
  Const Invalid_c As Long = 5&  ' "Invalid procedure call or argument"
  Err.Raise Invalid_c, ID_c, _
    "Invalid function argument - out of valid domain" & vbLf & _
    "Need 0 <= prob <= 1 but prob = " & prob & vbLf & _
    "Problem in " & ID_c
End If
' Calculate the inverse cumulative distribution function (quantile). The initial
' approximation used here has equal-ripple absolute error of 6.3E-11 or
' less except for prob > 0.9999997 where granularity in 'prob' dominates.
' The worst relative error of the corresponding PDF is about 2E-8.
' Note that the positive tail has zero accuracy above 'prob' = 1 - 1.1E-16;
' If you want accurate tail values use the negative tail instead.
' See Maple worksheet "InverseCumulativeNormal4.mws" for derivation.
' This initial result is then (usually) subjected to one Newton iteration to
' make it be the inverse of StdNormCum(z) exactly, or to within a few bits.
' Thus StdNormCum(StdNormInvCum(prob)) = prob to high accuracy, and this
' function therefore shares the high accuracy of StdNormCum(z).
Dim u As Double
u = prob - 0.5
Dim ua As Double
ua = Abs(u)
Dim u2 As Double
u2 = u * u
Dim z As Double  ' standard normal variate
If ua < 0.37377 Then  ' prob > 0.12623 and prob < 0.87377
  ' unit test: snIC(0.3)=-0.5244005127080408   snIC(0.7)=0.5244005127080408
  z = (1.10320796594653 - (7.77941365082975 - (16.1360412312915 - _
    8.94247760684027 * u2) * u2) * u2) * u / _
    (0.440116302105953 - (3.56442583646134 - (9.15646709284907 - _
    (7.69878138754029 - u2) * u2) * u2) * u2)
ElseIf ua < 0.44286 Then  ' prob > 0.05714 and prob < 0.94286, except for above
  ' unit test: snIC(0.09)=-1.34075503369022   snIC(0.91)=1.34075503369022
  z = (0.317718558863025 - (2.70051978050927 - (7.20258279324852 - _
    5.82660777818178 * u2) * u2) * u2) * u / _
    (0.126757926972973 - (1.21032688875879 - (3.85234822216469 - _
    (4.38840255884193 - u2) * u2) * u2) * u2)
ElseIf (prob < 4.94065645841247E-324) Or (prob = 1#) Then  ' off ends
  ' negative max is snIC(4.94065645841247E-324)=-38.467405617106 so jump is small
  ' positive max is snIC(1 - 1E-16)=8.2095361515899 so jump is larger
  z = Sgn(u) * 38.4674056172733  ' make the two "off-the-end" values the same
Else  ' out in tails - use expansion in w = Sqr(-Log(p))  -->  p = Exp(-w^2)
  Dim w As Double
  If prob < 0.5 Then  ' 4.94E-324 <= p < 0.05714
    w = Sqr(-Log(prob))
  Else  ' 0.94286 <= p < 1
    w = Sqr(-Log(1# - prob))  ' has roundoff noise > error above 0.9999997
  End If
  If w < 3.769 Then  ' prob or 1-prob > 6.77158141318452E-07 and < 0.05714
    ' unit test: snIC(1E-4)=-snIC(1 - 1E-4)=-3.71901648545568
    w = (3.40265621744676 + (9.03080228605413 - (6.88823432035713 + _
      (9.47396446577765 + 1.41485388628381 * w) * w) * w) * w) / _
      (1.10738880205572 + (7.00041795498572 + (6.72600088945649 + w) * w) * w)
   ElseIf w < 8.371 Then  ' prob or 1-prob > 3.69321326562547E-31 and < 6.77E-7
    ' unit test: snIC(1E-8)=-snIC(1 - 1E-8)=-5.61200124417479
     w = (27.5896468790036 + (11.8481686174627 - (37.7133528390963 + _
       (18.6301980539071 + 1.41437483654701 * w) * w) * w) * w) / _
       (10.7729777720728 + (29.1330213184579 + (13.1871785457772 + w) * w) * w)
   Else  ' prob <= 3.69321326562547E-31
    ' unit test: snIC(1E-33)=-12.0474677869249
     w = (859575101.771399 - (167079541.087701 + (887823598.683122 + _
       (206626688.300811 + 7785160.41001698 * w) * w) * w) * w) / _
       (382471838.745491 + (643197787.097259 + (146148247.666043 + _
       (5504690.97847543 + w) * w) * w) * w)
   End If
   z = Sgn(-u) * w
End If
If Abs(z) <= 37.6771207204951 Then  ' Exp(0.5 * z * z) will not overflow
' do one Newton step to become almost the exact inverse of StdNormCum(z)
  Const Sqrt2pi_ As Double = 2.506628274631  ' the next 4 digits are 0005
  z = z - (StdNormCum(z) - prob) * Sqrt2pi_ * Exp(0.5 * z * z)
End If
StdNormInvCum = z
End Function

'===============================================================================
Public Function StepExp( _
  ByVal arg As Double, _
  Optional ByVal wide As Double = 0.001) _
As Double
Attribute StepExp.VB_Description = "Smoothed unit step function around zero, near 0 for arg < 0 and near 1 for x > 0. Absolute error falls off as Exp(-x)"
' Smoothed step function around zero, near 0 for arg < 0 and near 1 for x > 0.
' Absolute error from exact Step falls off as Exp(-x) well away from the origin.
' See the function "IntSmooth" for a smooth step at every integer.
Const ID_c As String = File_c & "StepExp"
' check the input width
wide = Abs(wide)
Const MinWide_c As Double = 0.0000000001
If wide < MinWide_c Then wide = MinWide_c  ' impose minimal sanity
If Abs(arg) >= 1000000000# * wide Then
  StepExp = 0.5 * (1# + Sgn(arg))  ' smoothing part is well below roundoff
Else
  Dim temp As Double
  temp = arg / wide  ' we have |temp| <= 1,000,000,000 by the test above
  ' The constant is 1 / ( 2^(2/3) - 1 ) (see "FudgeGallery.mws" for why)
  StepExp = 0.5 * (1# + temp / Sqr(temp * temp + 1.702414384))
End If
End Function

'===============================================================================
Public Function StepSqr( _
  ByVal arg As Double, _
  Optional ByVal wide As Double = 0.001) _
As Double
Attribute StepSqr.VB_Description = "Smoothed unit step function around zero, near 0 for arg < 0 and near 1 for x > 0. Absolute error falls off as 1/x^2"
' Smoothed step function around zero, near 0 for arg < 0 and near 1 for x > 0.
' When |arg| = wide, the result is 0.195845649775679 or 0.804154350224321
' When |arg| = 1.73969 * wide, the result is 0.01 or 0.99
' When |arg| = 20.5992 * wide, the result is 0.001 or 0.999
' When |arg| = 1000 * wide, the result is 4.2560305256356E-7 or 0.99999957439695
' Absolute error from exact Step falls off as 1/x^2 well away from the origin.
' See the function "IntSmooth" for a smooth step at every integer.
Const ID_c As String = File_c & "StepSqr"
' check the input width
wide = Abs(wide)
Const MinWide_c As Double = 0.0000000001
If wide < MinWide_c Then wide = MinWide_c  ' impose minimal sanity
If Abs(arg) >= 1000000000# * wide Then
  StepSqr = 0.5 * (1# + Sgn(arg))  ' smoothing part is well below roundoff
Else
  Dim temp As Double
  temp = arg / wide  ' we have |temp| <= 1,000,000,000 by the test above
  ' The constant is 1 / ( 2^(2/3) - 1 ) (see "FudgeGallery.mws" for why)
  StepSqr = 0.5 * (1# + temp / Sqr(temp * temp + 1.702414384))
End If
End Function

'===============================================================================
Public Function Tanh(ByVal x As Double) As Double
Attribute Tanh.VB_Description = "The hyperbolic tangent of the input argument.  -1 <= Tanh <= 1. Correct even for very small |x|"
' The hyperbolic tangent of the input argument.  -1 <= Tanh <= 1.
Dim ax As Double
ax = Abs(x)
If ax < 0.503 Then  ' roundoff would spoil mathematical identity
  Dim x2 As Double
  x2 = x * x
  ' worst relative error 3E-16 - see the Maple file "TanhApprox.mws"
  Tanh = x - x * (19.0563004085872 + (0.9231700917023 + _
    0.0003630096197 * x2) * x2) * x2 / _
    (57.1689012257678 + (25.637070765062 + x2) * x2)
ElseIf ax < 20# Then  ' mathematical identity with Exp(2 * x) is accurate
  Dim temp As Double
  temp = Exp(x + x)
  Tanh = (temp - 1#) / (temp + 1#)
Else  ' result is +1 or -1 to IEEE 754 accuracy
  Tanh = Sgn(x)
End If
End Function

'===============================================================================
Public Function TenTo(ByVal p As Double) As Double
Attribute TenTo.VB_Description = "Result is 10 to the supplied power 'p'. Must have |p| <= 308.254715559916"
' Result is 10 to the supplied power 'p'. Must have |p| <= 308.254715559916
Const ID_c As String = File_c & "TenTo"
Const Ln_10 As Double = 2.302585 + 9.29940457E-08  ' sum is exact to last bit
If Abs(p) > 308.254715559916 Then  ' Exp() will overflow or underflow
  Const Invalid_c As Long = 5&  ' "Invalid procedure call or argument"
  Err.Raise Invalid_c, ID_c, _
    "Invalid function argument (Exp overflow)" & vbLf & _
    "Need |p| < 308.254715559917 but" & vbLf & _
    "p = " & p & vbLf & _
    "Problem in " & ID_c
End If
TenTo = Exp(p * Ln_10)
End Function

'===============================================================================
#If False Then  ' True here to enable this routine, False to disable it
'#If True Then  ' True here to enable this routine, False to disable it
Public Function findErrorPoint(ByVal x1 As Double, ByVal x2 As Double)
' Use bisection to find the input argument where errors start being thrown.
' Start with x1 = no error, x2 = error. Insert function to test below.
' This routine is used during development
Dim x As Double, d As Double, f As Double
x = x1  ' low end of interval
d = x2 - x1  ' high end minus low end
Do While d > 1E-16 * x  ' relative error desired
'Do While d > 0.00000000000001  ' absolute error desired
  d = 0.5 * d
  On Error Resume Next
  ' ***** insert function to be tested here *****
  f = TenTo(x + d)
  ' *********************************************
  If 0& = Err.Number Then x = x + d
  On Error GoTo 0
Loop
findErrorPoint = x + 0.5 * d
End Function
#End If

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


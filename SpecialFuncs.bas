Attribute VB_Name = "SpecialFuncs"
'
'###############################################################################
'#
'# Special function code in VBA
'#
'# version of 19 Jan 2002 by John Trenholme
'#
'###############################################################################

Option Explicit

Public Function cumNormComp(x As Double) As Double
'===============================================================================
' Complement of the cumulative distribution of a standard normal distribution.
' This is also known as the exceedance.
' Relative error is 1.2e-6 or less in absolute magnitude everywhere.
' Devised and coded by John Trenholme - version of 11 Oct 2001
'===============================================================================
  Dim u As Double, v As Double
u = Abs(x)
If u > 37# Then
  If x > 0# Then cumNormComp = 0# Else cumNormComp = 1#
Else
  If u < 5.0688410353692 Then
    u = (0.5 + u * (0.21733275 + u * (0.0422167 - u * 0.000047146755))) / _
        (1# + u * (1.2325829 + u * (0.56760466 + u * 0.10335289)))
  Else
    v = 1# / (u * u)
    u = (0.39894228 - v * (0.39869609 - v * (1.1553681 - v * 3.8645875))) / u
  End If
  u = Exp(-x * x / 2#) * u
  If x >= 0# Then cumNormComp = u Else cumNormComp = 1# - u
End If
End Function

Public Function ASin(value As Double) As Double
'===============================================================================
' arc sine
' error if value is outside the range [-1,1]
'===============================================================================
If Abs(value) <> 1 Then
    ASin = Atn(value / Sqr(1 - value * value))
Else
    ASin = 1.5707963267949 * Sgn(value)
End If
End Function

Public Function ACos(ByVal number As Double) As Double
'===============================================================================
' arc cosine
' error if NUMBER is outside the range [-1,1]
'===============================================================================
If Abs(number) <> 1 Then
    ACos = 1.5707963267949 - Atn(number / Sqr(1 - number * number))
ElseIf number = -1 Then
    ACos = 3.14159265358979
End If
'elseif number=1 --> Acos=0 (implicit)
End Function

Public Function ACot(value As Double) As Double
'===============================================================================
' arc cotangent
' error if NUMBER is zero
'===============================================================================
ACot = Atn(1 / value)
End Function

Public Function ASec(value As Double) As Double
'===============================================================================
' arc secant
' error if value is inside the range [-1,1]
'===============================================================================
' NOTE: the following lines can be replaced by a single call
'            ASec = ACos(1 / value)
If Abs(value) <> 1 Then
    ASec = 1.5707963267949 - Atn((1 / value) / Sqr(1 - 1 / (value * value)))
Else
    ASec = 3.14159265358979 * Sgn(value)
End If
End Function

Public Function ACsc(value As Double) As Double
'===============================================================================
' arc cosecant
' error if value is inside the range [-1,1]
'===============================================================================
' NOTE: the following lines can be replaced by a single call
'            ACsc = ASin(1 / value)
If Abs(value) <> 1 Then
    ACsc = Atn((1 / value) / Sqr(1 - 1 / (value * value)))
Else
    ACsc = 1.5707963267949 * Sgn(value)
End If
End Function

'----------------------------- end of file -------------------------------------


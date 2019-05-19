Attribute VB_Name = "InvCum"
Attribute VB_Description = "Inverse cumulative distribution functions of common statistical distributions. Devised & coded by John Trenholme."
'
'###############################################################################
'#
'# Visual Basic (VBA or VB6, not .NET) Module "InvCum" in file "InvCum.bas"
'#
'# Inverse cumulative distribution functions of common statistical distributions
'#
'# This module exports the following routines:
'#
'# Function binaryIC
'# Function cauchyIC
'# Function cosineIC
'# Function cosSquaredIC
'# Function exponentialIC
'# Function gumbelIC
'# Function InvCumVersion
'# Function laplaceIC
'# Function logisticIC
'# Function logNormalIC
'# Function normalIC
'# Function parabolaIC
'# Function parabSquaredIC
'# Function paretoIC
'# Function pwrWarp
'# Function rayleighIC
'# Function sechIC
'# Function standardNormalIC
'# Function triangularIC
'# Function uniformIC
'# Function uniformLongIC
'# Function weibullIC
'#
'# ~~~~~~~~~~~~~~~~ Devised and coded by John Trenholme ~~~~~~~~~~~~~~~~
'#
'###############################################################################

Option Base 0          ' array base value when not specified - the default
Option Compare Binary  ' string comparison based on Asc(char) - the default
Option Explicit        ' forces explicit variable declaration - changes default

Private Const Version_c As String = "2013-05-07"  ' update manually on each edit
Private Const M_c As String = "InvCum[" & Version_c & "]."  ' Module ID

Private Const InvalidArg_c As Long = 5&  ' "Invalid procedure call or argument"

'###############################################################################
'#
'# Supplies inverse cumulative distribution functions of common statistical
'# distributions. Such functions return the value of a variate (quantile) that
'# gives the supplied (input) probability value of the cumulative distribution
'# function (CDF).
'#
'# If a sequence of uniform pseudo-random probability values (0 <= P <= 1)
'# is supplied to such a function, a pseudo-random sequence having the
'# corresponding probability distribution function (PDF) is produced. This
'# method of producing pseudo-random samples from statistical distributions has
'# several advantages compared to methods where calls to a uniform pseudo-random
'# source are built into the distribution functions:
'#   1) only one uniform variate is consumed per sample
'#   2) it is easy to switch among different sources of uniform variates
'#   3) it is easy to introduce correlations among distribution samples,
'#      by using correlated input probabilities. In variance reduction methods,
'#      negatively-correlated variates are very useful. Search online for
'#      "latin hypercube sampling with dependence" for more information.
'#   4) a reduced output range may be produced by supplying a reduced input
'#      range; for example, only right-side Gaussians by supplying P >= 0.5
'#   5) quantiles of the PDF are easily produced (these are points where the
'#      CDF has a specified value, such as 0.5, 0.75, 0.9, 0.95 etc.).
'# Of course, this method also has disadvantages:
'#   1) it may be slower than other methods
'#   2) an approximation to the inverse cumulative function may be required
'#      when no functional inverse is available
'#   3) some distributions have no known inverse cumulative function
'#      even in approximate form; if the cumulative distribution is available,
'#      nonlinear root-finding can give the inverse CDF, but that is slow
'#
'# The functions here have been written so that increasing values of the
'# probability always cause increasing values of the variate. That is, all
'# functions here are increasing functions of their argument P.
'#
'# The argument sequence to all functions consists of a probability (which
'# must obey 0 <= P <= 1), followed by one or more parameters of the PDF from
'# which the inverse cumulative function was calculated, followed by limit
'# points if the distribution is limited on the low side, or on both sides.
'# Useful default values for the parameters and limits are usually supplied.
'#
'# The distributions (PDF's) here come in three types:
'#
'#   2-tailed distributions extend, in principle, to + infinity and - infinity.
'#   In practice, floating-point representation restrictions limit the tails to
'#   finite values, but the truncated extreme tails only lose a small part of
'#   the cumulative probability (typically 1E-15 or less).
'#
'#   1-tailed distributions extend from a fixed lower limit to, in principle,
'#   + infinity.  Values below the lower limit will never be returned. As with
'#   the 2-tailed case, the tail is truncated at a finite value with very low
'#   lost probability. To make a 1-tailed distribution extend toward - infinity,
'#   negate the return value (it then becomes a decreasing function of P, so you
'#   may wish to supply 1-P as the input probability).
'#
'#   0-tailed distributions extend from a fixed lower limit to a fixed upper
'#   limit. Values beyond the limits will never be returned.
'#
'# Any of these distributions can be modified by use of the "pwrWarp" function.
'# An entire family of distributions with different kurtosis values can be
'# supplied this way. See the comments for "pwrWarp" for details
'#
'###############################################################################

' Note: if a routine in this Module conflicts with one of the same name in
' another Module, prefix the name with the name of this Module, plus a period.
' For example, use InvCum.exponentialIC(p, mu) instead of exponentialIC(p, mu).

' Note: the inverse of the maximum of N independent samples, all with the same
' probability P, is InvCum(P ^ (1# / N))

' Note: when a Function that is defined to return a numeric value raises an
' error, Excel reports the result as #VALUE! This is confusingly explained by
' Microsoft as "A value used in the formula is of the wrong data type" when the
' actual problem is (usually) an argument value error. When functions return
' a Variant, it can hold arbitrary errors, but then extracting the numeric value
' takes an extra amount of time, so we don't do that.

'===============================================================================
Public Function binaryIC( _
  ByVal prob As Double, _
  Optional ByVal success As Double = 0.5, _
  Optional ByVal loEnd As Double = 0#, _
  Optional ByVal hiEnd As Double = 1#) _
As Double
Attribute binaryIC.VB_Description = "ICDF of a binary (Bernoulli) distribution that models the flip of a biased coin. If 'prob' is less than 'success' the value is 'loEnd', otherwise it is 'hiEnd'."
' ICDF of a binary (Bernoulli) distribution. If 'prob' is less than 'success'
' (default 0.5) the value is 'loEnd' (default 0), otherwise it is 'hiEnd'
' (default 1). The PDF is a delta function of area 'success' at 'loEnd' and
' another delta function of area 1-success at 'hiEnd', and the CDF is a step
' function of height 'success' at 'loEnd' and another step function of height
' 1-success at 'hiEnd'.
' This is a 0-tailed discrete distribution.
' Devised and coded by John Trenholme.
Const ID As String = M_c & "binaryIC"
' check the input probability
If (prob < 0#) Or (prob > 1#) Then
  Err.Raise InvalidArg_c, ID, _
  "Argument ERROR. Need 0 <= prob <= 1 but prob = " & prob & vbLf & _
  "Problem in " & ID
End If
' calculate the inverse cumulative distribution function (quantile)
If prob <= success Then binaryIC = loEnd Else binaryIC = hiEnd
End Function

'===============================================================================
Public Function cauchyIC( _
  ByVal prob As Double, _
  Optional ByVal mean As Double = 0#, _
  Optional ByVal hwhm As Double = 1#) _
As Double
Attribute cauchyIC.VB_Description = "ICDF of a Cauchy distribution with the supplied mean and half-width-at-half-maximum values. This distribution has 'heavy' tails that fall off only as 1/x^2."
' ICDF of a Cauchy distribution with the supplied mean and half-width-at-half-
' maximum values. You can't specify a standard deviation because this
' distribution has none - it has "heavy tails" that drop off as 1 / x^2, so the
' standard deviation is infinity.
' Note: the values at P = 0 and P = 1 are properly +- infinity, but this
' Function returns +- 1.63317787283838E+16
' This is a 2-tailed symmetric continuous distribution.
' Devised and coded by John Trenholme.
Const ID As String = M_c & "cauchyIC"
' check the input probability
If (prob < 0#) Or (prob > 1#) Then
  Err.Raise InvalidArg_c, ID, _
  "Argument ERROR. Need 0 <= prob <= 1 but prob = " & prob & vbLf & _
  "Problem in " & ID
End If
' check the distribution parameters
If hwhm < 0# Then  ' we could allow negative values, but it's likely user error
  Err.Raise InvalidArg_c, ID, _
  "Argument ERROR. Need hwhm >= 0 but hwhm = " & hwhm & vbLf & _
  "Problem in " & ID
End If
' calculate the inverse cumulative distribution function (quantile)
Const Pi_c As Double = 3.1415926 + 5.358979324E-08  ' good to last bit
cauchyIC = mean + hwhm * Tan(Pi_c * (prob - 0.5))
End Function

'===============================================================================
Public Function cosineIC( _
  ByVal prob As Double, _
  Optional ByVal loEnd As Double = 0#, _
  Optional ByVal hiEnd As Double = 1#) _
As Double
Attribute cosineIC.VB_Description = "ICDF of a cosine distribution; PDF is a cosine positive half-cycle that is zero at 'loEnd' and below, and zero at 'hiEnd' and above."
' ICDF of a cosine distribution; PDF is a cosine positive half-cycle that is
' zero at 'loEnd' and below, maximum at (loEnd+hiEnd)/2, and zero at 'hiEnd'
' and above.
' This is a 0-tailed symmetric continuous distribution.
' Devised and coded by John Trenholme.
Const ID As String = M_c & "cosineIC"
' check the input probability
If (prob < 0#) Or (prob > 1#) Then
  Err.Raise InvalidArg_c, ID, _
  "Argument ERROR. Need 0 <= prob <= 1 but prob = " & prob & vbLf & _
  "Problem in " & ID
End If
' calculate the inverse cumulative distribution function (quantile)
If prob = 0# Then  ' avoid divide-by-zero
  cosineIC = loEnd
ElseIf prob = 1# Then  ' avoid divide-by-zero
  cosineIC = hiEnd
Else
  Const PiInv_c As Double = 0.31830988 + 6.1837907E-09  ' 1 / Pi
  cosineIC = 0.5 * (loEnd + hiEnd) + (hiEnd - loEnd) * _
    PiInv_c * Atn((prob - 0.5) / Sqr(prob * (1# - prob)))
End If
End Function

'===============================================================================
Public Function cosSquaredIC( _
  ByVal prob As Double, _
  Optional ByVal loEnd As Double = 0#, _
  Optional ByVal hiEnd As Double = 1#) _
As Double
Attribute cosSquaredIC.VB_Description = "ICDF of a cosine-squared distribution; PDF is a squared cosine that is zero at 'loEnd' and below, and zero at 'hiEnd' and above."
' ICDF of a cosine-squared distribution; PDF is a squared cosine that is zero
' at 'loEnd' and below, maximum at (loEnd+hiEnd)/2, and zero at 'hiEnd'
' and above.
' This is a 0-tailed symmetric continuous distribution.
' Devised and coded by John Trenholme.
Const ID As String = M_c & "cosSquaredIC"
' check the input probability
If (prob < 0#) Or (prob > 1#) Then
  Err.Raise InvalidArg_c, ID, _
  "Argument ERROR. Need 0 <= prob <= 1 but prob = " & prob & vbLf & _
  "Problem in " & ID
End If
' calculate the inverse cumulative distribution function (quantile)
Dim u As Double
If prob < 0.5 Then
  u = prob
Else
  u = (1# - prob)
End If
Const OneThird_c As Double = 1# / 3#
u = u ^ OneThird_c
Dim t As Double
' this approximation has equal-ripple absolute error of 1E-11 or less
' see the Maple worksheet "InverseCumulativeDistributions.mws"
t = (100.513937727401 - (166.490764348476 - (25.9299365731226 + _
  (65.3708595919531 - (26.6593395804447 - (2.86646254511049 - u) * _
  u) * u) * u) * u) * u) * u / _
  (188.348623507727 - (311.979716331963 - (13.295934698858 + _
  (180.947236835638 - (69.3953608593967 + 0.611277137973115 * u) * _
   u) * u) * u) * u)
If prob < 0.5 Then
  cosSquaredIC = loEnd + (hiEnd - loEnd) * t
Else
  cosSquaredIC = hiEnd - (hiEnd - loEnd) * t
End If
End Function

'===============================================================================
Public Function exponentialIC( _
  ByVal prob As Double, _
  Optional ByVal stdDev As Double = 1#, _
  Optional ByVal loEnd As Double = 0#) _
As Double
Attribute exponentialIC.VB_Description = "ICDF of an exponential distribution starting at 'loEnd' with a standard deviation of 'stdDev.' If you want a negative-going tail, negate the return value (and perhaps pass 1-P as the argument)."
' ICDF of an exponential distribution with the supplied standard deviation
' value. Note that values more than 37.43 times 'stdDev' above 'loEnd' will
' never be returned. The probability of such large values is less than 1E-16.
' The PDF tail drops off as Exp(-x) for x >= loEnd. If you want a negative-
' going tail, negate the return value (and perhaps pass 1-P as the argument).
' This is a 1-tailed asymmetric continuous distribution.
' Devised and coded by John Trenholme.
Const ID As String = M_c & "exponentialIC"
' check the input probability
If (prob < 0#) Or (prob > 1#) Then
  Err.Raise InvalidArg_c, ID, _
  "Argument ERROR. Need 0 <= prob <= 1 but prob = " & prob & vbLf & _
  "Problem in " & ID
End If
' check the distribution parameter
If stdDev < 0# Then
  Err.Raise InvalidArg_c, ID, _
  "Argument ERROR. Need stdDev >= 0 but stdDev = " & stdDev & vbLf & _
  "If you want a negative-going tail, negate the return value." & vbLf & _
  "Problem in " & ID
End If
' calculate the inverse cumulative distribution function (quantile)
If prob = 0# Then  ' avoid return of -0 when loEnd = 0
  exponentialIC = loEnd
ElseIf prob = 1# Then
  ' avoid Log(0); use 1# - prob = 2^(-54) (half a bit below 1.0)
  Const LogTiny_c As Double = -37.429947750237  ' Log(2^(-54))=Log(5.551115E-17)
  exponentialIC = loEnd - stdDev * LogTiny_c
Else
  ' use of "1-prob" makes this a strictly non-decreasing function
  exponentialIC = loEnd - stdDev * Log(1# - prob)
End If
End Function

'===============================================================================
Public Function gumbelIC( _
  ByVal prob As Double, _
  Optional ByVal mean As Double = 0#, _
  Optional ByVal stdDev As Double = 1#) _
As Double
Attribute gumbelIC.VB_Description = "ICDF of a Gumbel extreme-value distribution with the supplied mean and standard deviation values. For a 'reverse Gumbel,' negate the return value (and perhaps pass 1-P as the argument)."
' ICDF of a Gumbel extreme-value distribution with the supplied mean and
' standard deviation values. Ths will never return a value below 5.6059015
' standard deviations below the mean, or more than 28.193513 standard
' deviations above the mean. The probability of such extreme values is less
' than 1E-16. For a "reverse Gumbel," negate the return value (and perhaps
' pass 1-P as the argument for strictly increasing behavior).
' This is a 2-tailed continuous distribution.
' Devised and coded by John Trenholme.
Const ID As String = M_c & "gumbelIC"
' check the input probability
If (prob < 0#) Or (prob > 1#) Then
  Err.Raise InvalidArg_c, ID, _
  "Argument ERROR. Need 0 <= prob <= 1 but prob = " & prob & vbLf & _
  "Problem in " & ID
End If
' check the distribution parameters
If stdDev < 0# Then
  Err.Raise InvalidArg_c, ID, _
  "Argument ERROR. Need stdDev >= 0 but stdDev = " & stdDev & vbLf & _
  "If you want a reverse Gumbel, negate the return value." & vbLf & _
  "Problem in " & ID
End If
' avoid Log(0)
Const Tiny_c As Double = 4.94065645841247E-324
If prob < Tiny_c Then prob = Tiny_c
Const NearOne_c As Double = 1# - 1.1E-16
If prob > NearOne_c Then prob = NearOne_c
' calculate the inverse cumulative distribution function (quantile)
Const Sqr6OvrPi_c As Double = 0.779696801233678
Const Euler_c As Double = 0.57721566 + 4.9015329E-09
gumbelIC = mean - stdDev * Sqr6OvrPi_c * (Log(-Log(prob)) + Euler_c)
End Function

'===============================================================================
Public Function InvCumVersion(Optional ByVal trigger As Variant) As String
Attribute InvCumVersion.VB_Description = "The date of the latest revision to this Module, in the format 'yyyy-mm-dd'."
' The date of the latest revision to this Module, in the format "yyyy-mm-dd".
' We use 'trigger' instead of making this Volatile, for use with Excel.
InvCumVersion = Version_c
End Function

'===============================================================================
Public Function laplaceIC( _
  ByVal prob As Double, _
  Optional ByVal mean As Double = 0#, _
  Optional ByVal stdDev As Double = 1#) _
As Double
Attribute laplaceIC.VB_Description = "ICDF of a Laplace distribution with the supplied mean and standard deviation values. The tails drop off as Exp(-|x|)."
' ICDF of a Laplace distribution with the supplied mean and standard deviation
' values. This is a symmetric version of the exponential distribution, with a
' sharp peak at the center. Note that values more than 26.46697 times the
' standard deviation away from the mean will never be returned; the probability
' of such extreme values is less than 2E-16. The tails drop off as Exp(-|x|).
' This is a 2-tailed symmetric continuous distribution.
' Devised and coded by John Trenholme.
Const ID As String = M_c & "laplaceIC"
' check the input probability
If (prob < 0#) Or (prob > 1#) Then
  Err.Raise InvalidArg_c, ID, _
  "Argument ERROR. Need 0 <= prob <= 1 but prob = " & prob & vbLf & _
  "Problem in " & ID
End If
' check the distribution parameters
If stdDev < 0# Then
  Err.Raise InvalidArg_c, ID, _
  "Argument ERROR. Need stdDev >= 0 but stdDev = " & stdDev & vbLf & _
  "Problem in " & ID
End If
' calculate the inverse cumulative distribution function (quantile)
Const LogTiny_c As Double = -37.429947750237  ' Log(2^(-54))=Log(5.551115E-17)
Const SqrHalf_c As Double = 0.707106781186548  ' Sqr(0.5)
If prob < 1.1E-16 Then  ' adjust very low prob to match high end
  laplaceIC = mean + SqrHalf_c * stdDev * LogTiny_c  ' value for 5.5E-17
ElseIf prob = 1# Then  ' value for 1.0 - 5.5E-17 (impossible prob)
  laplaceIC = mean - SqrHalf_c * stdDev * LogTiny_c
Else  ' value is in central region, use analytic formula
  If prob < 0.5 Then  ' negative-going exponential decay
    laplaceIC = mean + SqrHalf_c * stdDev * Log(prob + prob)
  Else  ' positive-going exponential decay
    laplaceIC = mean - SqrHalf_c * stdDev * Log(2# - prob - prob)
  End If
End If
End Function

'===============================================================================
Public Function logisticIC( _
  ByVal prob As Double, _
  Optional ByVal mean As Double = 0#, _
  Optional ByVal stdDev As Double = 1#) _
As Double
Attribute logisticIC.VB_Description = "ICDF of a logistic distribution with the supplied mean and standard deviation values. The tails drop off as Exp(-|x|)."
' ICDF of a logistic distribution with the supplied mean and standard deviation
' values. Note that values more than 20.636212 times the standard deviation away
' from the mean will never be returned; the probability of such extreme values
' is less than 2E-16. The tails drop off as Exp(-|x|).
' This is a 2-tailed symmetric continuous distribution.
' Devised and coded by John Trenholme.
Const ID As String = M_c & "logisticIC"
' check the input probability
If (prob < 0#) Or (prob > 1#) Then
  Err.Raise InvalidArg_c, ID, _
  "Argument ERROR. Need 0 <= prob <= 1 but prob = " & prob & vbLf & _
  "Problem in " & ID
End If
' check the distribution parameters
If stdDev < 0# Then
  Err.Raise InvalidArg_c, ID, _
  "Argument ERROR. Need stdDev >= 0 but stdDev = " & stdDev & vbLf & _
  "Problem in " & ID
End If
' calculate the inverse cumulative distribution function (quantile)
Const LogTiny_c As Double = -37.429947750237  ' Log(2^(-54))=Log(5.551115E-17)
Const Sqr3ovrPi_c As Double = 0.551328895421793  ' Sqr(3)/Pi
If prob < 1.1E-16 Then  ' adjust very low prob to match high end
  logisticIC = mean + Sqr3ovrPi_c * stdDev * LogTiny_c ' value for 5.5E-17
ElseIf prob = 1# Then  ' value for 1.0 - 5.5E-17 (impossible prob)
  logisticIC = mean - Sqr3ovrPi_c * stdDev * LogTiny_c
Else  ' value is in central region, use analytic formula
  logisticIC = mean + Sqr3ovrPi_c * stdDev * Log(prob / (1# - prob))
End If
End Function

'===============================================================================
Public Function logNormalIC( _
  ByVal prob As Double, _
  Optional ByVal mean As Double = 1#, _
  Optional ByVal stdDev As Double = 1#, _
  Optional ByVal loEnd As Double = 0#) _
As Double
Attribute logNormalIC.VB_Description = "ICDF of a log-normal distribution with the supplied mean and standard deviation values (values for the log-normal distribution itself, not the normal distribution in the exponent). If you want a negative-going tail, negate the return value (and perhaps pass 1-P as the argument)."
' ICDF of a log-normal distribution with the supplied mean and standard
' deviation values. Note that these are the values for the log-normal
' distribution itself, not the normal distribution in the exponent. When the
' standard deviation is much less than the mean, this approaches a Gaussian.
' This is a 1-tailed continuous distribution.
' Devised and coded by John Trenholme.
Const ID As String = M_c & "logNormalIC"
' check the input probability
If (prob < 0#) Or (prob > 1#) Then
  Err.Raise InvalidArg_c, ID, _
  "Argument ERROR. Need 0 <= prob <= 1 but prob = " & prob & vbLf & _
  "Problem in " & ID
End If
' check the distribution parameters
If mean < loEnd Then
  Err.Raise InvalidArg_c, ID, _
  "Argument ERROR. Need mean >= " & loEnd & " but mean = " & mean & vbLf & _
  "If you want a negative-going tail, negate the return value" & vbLf & _
  "Problem in " & ID
End If
If stdDev < 0# Then
  Err.Raise InvalidArg_c, ID, _
  "Argument ERROR. Need stdDev >= 0 but stdDev = " & stdDev & vbLf & _
  "Problem in " & ID
End If
Dim mean0 As Double
mean0 = mean - loEnd  ' use distance above loEnd
' calculate the inverse cumulative distribution function (quantile)
If (prob = 0#) Or (mean0 = 0#) Then  ' avoid trouble later
  logNormalIC = loEnd
Else
  Dim c As Double, c2 As Double, t As Double ' "contrast" & Log(1 + c^2)
  c = stdDev / mean0
  c2 = c * c
  ' Log(1 + x) has a severe roundoff problem for small x - avoid it
  If c < 0.0484 Then  ' use series expansion; relative trunc. error < 4.5E-14
    t = (0.999999999999954 - (0.499999999358952 - (0.333331964738561 - _
         0.249064956072441 * c2) * c2) * c2) * c2
  Else  ' relative roundoff error < 4.5E-14
    t = Log(1# + c2)
  End If
  Dim m As Double, s As Double  ' mean & stdDev of Gaussian in exponent
  m = Log(mean0) - 0.5 * t
  s = Sqr(t)
  logNormalIC = loEnd + Exp(m + s * standardNormalIC(prob))
End If
End Function

'===============================================================================
Public Function normalIC( _
  ByVal prob As Double, _
  Optional ByVal mean As Double = 0#, _
  Optional ByVal stdDev As Double = 1#) _
As Double
Attribute normalIC.VB_Description = "ICDF of a normal (Gaussian) distribution with the supplied mean and standard deviation values. For a standard normal, use the 'standardNormalIC' function."
' ICDF of a normal (Gaussian) distribution with the supplied mean and standard
' deviation values. Note that values more than 38.47 times the standard
' deviation away from the mean will never be returned; the probability of such
' extreme values is less than 1E-16. The tails drop off as Exp(-x^2).
' If you just want the inverse CDF of a standard normal variate, call the
' "standardNormalIC" function directly; this is a wrapper for that function.
' This is a 2-tailed symmetric continuous distribution.
' Devised and coded by John Trenholme.
Const ID As String = M_c & "normalIC"
' check the input probability
If (prob < 0#) Or (prob > 1#) Then
  Err.Raise InvalidArg_c, ID, _
  "Argument ERROR. Need 0 <= prob <= 1 but prob = " & prob & vbLf & _
  "Problem in " & ID
End If
' check the distribution parameters
If stdDev < 0# Then
  Err.Raise InvalidArg_c, ID, _
  "Argument ERROR. Need stdDev >= 0 but stdDev = " & stdDev & vbLf & _
  "Problem in " & ID
End If
' calculate the inverse cumulative distribution function (quantile)
normalIC = mean + stdDev * standardNormalIC(prob)
End Function

'===============================================================================
Public Function parabolaIC( _
  ByVal prob As Double, _
  Optional ByVal loEnd As Double = 0#, _
  Optional ByVal hiEnd As Double = 1#) _
As Double
Attribute parabolaIC.VB_Description = "ICDF of a parabolic distribution; PDF starts at 0 at 'loEnd', rises quadratically to center, then drops back quadratically to 0 at 'hiEnd'."
' ICDF of a parabolic distribution; PDF(x) starts at 0 at x = loEnd, rises
' quadratically to center at (loEnd+hiEnd)/2, then drops back quadratically
' to 0 at x = hiEnd. PDF is zero below loEnd and above hiEnd.
' This is a 0-tailed symmetric continuous distribution.
' Devised and coded by John Trenholme.
Const ID As String = M_c & "parabolaIC"
' check the input probability
If (prob < 0#) Or (prob > 1#) Then
  Err.Raise InvalidArg_c, ID, _
  "Argument ERROR. Need 0 <= prob <= 1 but prob = " & prob & vbLf & _
  "Problem in " & ID
End If
' calculate the inverse cumulative distribution function (quantile)
If prob = 0# Then  ' avoid divide-by-zero
  parabolaIC = loEnd
ElseIf prob = 1# Then  ' avoid divide-by-zero
  parabolaIC = hiEnd
Else
  Const OneThird_c As Double = 1# / 3#
  parabolaIC = 0.5 * (loEnd + hiEnd) + (hiEnd - loEnd) * _
    Sin(OneThird_c * Atn((prob - 0.5) / Sqr(prob * (1# - prob))))
End If
End Function

'===============================================================================
Public Function parabSquaredIC( _
  ByVal prob As Double, _
  Optional ByVal loEnd As Double = 0#, _
  Optional ByVal hiEnd As Double = 1#) _
As Double
Attribute parabSquaredIC.VB_Description = "ICDF of a parabola-squared (or 'biweight') distribution. PDF is a squared parabola that is zero below 'loEnd' and above 'hiEnd.'"
' ICDF of a parabola-squared (or "biweight") distribution; PDF is the square
' of a parabola that is zero at 'loEnd', rises to center at (loEnd+hiEnd)/2,
' and is zero again at 'hiEnd'. PDF is zero below loEnd and above hiEnd.
' This is a 0-tailed symmetric continuous distribution.
' Devised and coded by John Trenholme.
Const ID As String = M_c & "parabSquaredIC"
' check the input probability
If (prob < 0#) Or (prob > 1#) Then
  Err.Raise InvalidArg_c, ID, _
  "Argument ERROR. Need 0 <= prob <= 1 but prob = " & prob & vbLf & _
  "Problem in " & ID
End If
' calculate the inverse cumulative distribution function (quantile)
Dim u As Double
If prob < 0.5 Then
  u = prob
Else
  u = (1# - prob)
End If
Const OneThird_c As Double = 1# / 3#
u = u ^ OneThird_c
Dim t As Double
' this approximation has equal-ripple absolute error of 2.3E-12 or less
' see the Maple worksheet "InverseCumulativeDistributions.mws"
t = (113.2352821494 - (322.476228095615 - (334.709997563481 - _
  (150.979373326071 - (26.6238411054252 - u) * u) * u) * u) * u) * u / _
  (243.958019863441 - (751.371608103629 - (866.581115791067 - _
  (455.857498344098 - (104.217059558827 - 7.40019588193932 * u) * _
  u) * u) * u) * u)
If prob <= 0.5 Then
  parabSquaredIC = loEnd + (hiEnd - loEnd) * t
Else
  parabSquaredIC = hiEnd - (hiEnd - loEnd) * t
End If
End Function

'===============================================================================
Public Function paretoIC( _
  ByVal prob As Double, _
  ByVal power As Double, _
  Optional ByVal fwhm As Double = 1#, _
  Optional ByVal loEnd As Double = 0#) _
As Double
Attribute paretoIC.VB_Description = "ICDF of a Pareto (type I) distribution with the supplied power-law exponent and full width at half maximum (FWHM) values. If you want a negative-going tail, negate the return value (and perhaps pass 1-P as the argument)."
' ICDF of a Pareto (type I) distribution with the supplied power-law exponent
' and full width at half maximum (FWHM) values. If power <= 2, the standard
' deviation is formally infinite, and if power <= 1 the mean value is formally
' infinite, so you might want to avoid low values of 'power' unless you want
' the very heavy tails that result. The PDF tail drops off as x^(power + 1).
' If you want a negative-going tail, negate the return value (and perhaps
' pass 1-P as the argument).
' This is a 1-tailed asymmetric continuous distribution.
' Devised and coded by John Trenholme.
Const ID As String = M_c & "paretoIC"
' check the input probability
If (prob < 0#) Or (prob > 1#) Then
  Err.Raise InvalidArg_c, ID, _
  "Argument ERROR. Need 0 <= prob <= 1 but prob = " & prob & vbLf & _
  "Problem in " & ID
End If
' check the distribution parameters
If power <= 0# Then
  Err.Raise InvalidArg_c, ID, _
  "Argument ERROR. Need power > 0 but power = " & power & vbLf & _
  "Problem in " & ID
End If
If fwhm < 0# Then
  Err.Raise InvalidArg_c, ID, _
  "Argument ERROR. Need fwhm >= 0 but fwhm = " & fwhm & vbLf & _
  "If you want a negative-going tail, negate the return value." & vbLf & _
  "Problem in " & ID
End If
' calculate the inverse cumulative distribution function (quantile)
If 1# = prob Then prob = 1# - 1.1E-16  ' silently truncate the tail
If prob = 0# Then  ' avoid extra work at 0
  paretoIC = loEnd
Else
  Dim t1 As Double, t2 As Double
  On Error Resume Next  ' avoid problems with large exponents
  t1 = 1# / ((1# - prob) ^ (1# / power)) - 1#
  t2 = 2# ^ (1# / (power + 1#)) - 1#
  On Error GoTo 0
  If t2 < 1E-150 Then t2 = 1E-150  ' silently avoid x/0
  paretoIC = loEnd + fwhm * t1 / t2
End If
End Function

'===============================================================================
Public Function pwrWarp( _
  ByVal prob As Double, _
  ByVal power As Double) _
As Double
Attribute pwrWarp.VB_Description = "Warps the input value 'prob', stretching it ('power' > 1) or compressing it ('power' < 1) around 'prob' = 0.5 while maintaining the end points. Must have 0 <= 'prob' <= 1 and 'power' > 0."
' This warps the input probability to change the fourth moment (kurtosis) of
' a PDF implied by an ICDF. When power > 1, the PDF is stretched around
' CDF = 0.5, and compressed near CDF = 0 and CDF = 1. When power < 1, the
' stretch is in the center and the compression is near the ends. Note that when
' power < 1, there can be a sharp peak near the center. Requires power > 0.
' Usage: someInvCumFunc(pwrWarp(prob, power))
Const ID As String = M_c & "pwrWarp"
' check the input probability
If (prob < 0#) Or (prob > 1#) Then
  Err.Raise InvalidArg_c, ID, _
  "Argument ERROR. Need 0 <= prob <= 1 but prob = " & prob & vbLf & _
  "Problem in " & ID
End If
If power <= 0# Then
  Err.Raise InvalidArg_c, ID, _
  "Argument ERROR. Need 0 < power but power = " & power & vbLf & _
  "Problem in " & ID
End If
pwrWarp = 0.5 * (1 + Sgn(prob - 0.5) * Abs(2# * prob - 1#) ^ power)
End Function

'===============================================================================
Public Function rayleighIC( _
  ByVal prob As Double, _
  Optional ByVal stdDev As Double = 1#, _
  Optional ByVal loEnd As Double = 0#) _
As Double
Attribute rayleighIC.VB_Description = "ICDF of a Rayleigh distribution with the supplied standard deviation value. If you want a negative-going tail, negate the return value (and perhaps pass 1-P as the argument)."
' ICDF of a Rayleigh distribution with the supplied standard deviation value.
' Note that values more than 13.2067 times 'stdDev' above 'loEnd' will never be
' returned; the probability of such large values is less than 1E-16. The PDF
' tail drops off as x * Exp( -x^2)
' This is a 1-tailed asymmetric continuous distribution.
' Devised and coded by John Trenholme.
Const ID As String = M_c & "rayleighIC"
' check the input probability
If (prob < 0#) Or (prob > 1#) Then
  Err.Raise InvalidArg_c, ID, _
  "Argument ERROR. Need 0 <= prob <= 1 but prob = " & prob & vbLf & _
  "Problem in " & ID
End If
' check the distribution parameter
If stdDev < 0# Then
  Err.Raise InvalidArg_c, ID, _
  "Argument ERROR. Need stdDev >= 0 but stdDev = " & stdDev & vbLf & _
  "If you want a negative-going tail, negate the return value" & vbLf & _
  "(and perhaps pass 1-P as the argument)" & vbLf & _
  "Problem in " & ID
End If
' calculate the inverse cumulative distribution function (quantile)
Const RayCon_c As Double = 2.15865522173539  ' 1 / Sqr(1 - Pi / 4)
If prob = 0# Then  ' avoid return of -0 when loEnd = 0
  rayleighIC = loEnd
ElseIf prob = 1# Then
  ' avoid Log(0); use 1# - prob = 2^(-54) (half a bit below 1.0)
  rayleighIC = loEnd + stdDev * RayCon_c * Sqr(-Log(5.55111512312578E-17))
Else
  ' use of "1-prob" forces positive-slope function
  rayleighIC = loEnd + stdDev * RayCon_c * Sqr(-Log(1# - prob))
End If
End Function

'===============================================================================
Public Function sechIC( _
  ByVal prob As Double, _
  Optional ByVal mean As Double = 0#, _
  Optional ByVal stdDev As Double = 1#) _
As Double
Attribute sechIC.VB_Description = "ICDF of a sech distribution (inverse cosh) with the supplied mean and standard deviation values. The tails drop off as Exp(-|x|)."
' ICDF of a sech distribution with the supplied mean and standard deviation
' values. Note that values more than 22.3 times the standard deviation away
' from the mean will never be returned; the probability of such extreme values
' is less than 1E-15. The tails drop off as Exp(-|x|).
' This is a 2-tailed symmetric continuous distribution.
' Unit test value: sechIC(0.75) = 0.56109985233918
' Devised and coded by John Trenholme.
Const ID As String = M_c & "sechIC"
' check the input probability
If (prob < 0#) Or (prob > 1#) Then
  Err.Raise InvalidArg_c, ID, _
  "Argument ERROR. Need 0 <= prob <= 1 but prob = " & prob & vbLf & _
  "Problem in " & ID
End If
' check the distribution parameter
If stdDev < 0# Then
  Err.Raise InvalidArg_c, ID, _
  "Argument ERROR. Need stdDev >= 0 but stdDev = " & stdDev & vbLf & _
  "Problem in " & ID
End If
' calculate the inverse cumulative distribution function (quantile)
Const Pi_c As Double = 3.1415926 + 5.358979324E-08  ' good to last bit
Const TwoOvrPi_c As Double = 2# / Pi_c
Dim u As Double, ua As Double
u = prob - 0.5  ' function is antisymmetric around midpoint
ua = Abs(u)
Dim w As Double
If ua <= (0.5 - 5E-16) Then ' value is in central region; use analytic formula
  Dim v As Double
  v = Tan(ua * Pi_c)  ' 0 <= v <= 544191874731457
  w = mean + Sgn(u) * Log(v + Sqr(1# + v * v)) * TwoOvrPi_c * stdDev
Else  ' result is getting granular due to roundoff error
  w = Sgn(u) * 22.3  ' truncate; lose probability of 5E-16 on each tail
End If
sechIC = mean + stdDev * w
End Function

'===============================================================================
Public Function standardNormalIC( _
  ByVal prob As Double) _
As Double
Attribute standardNormalIC.VB_Description = "ICDF of a standard normal distribution (Gaussian with zero mean and unit standard deviation)."
' ICDF of a standard normal distribution (Gaussian with zero mean and unit
' standard deviation). Note that values more than +- 38.47 will never be
' returned; the probability of such extreme values is less than 2E-16.
' This is a 2-tailed symmetric continuous distribution.
' Devised and coded by John Trenholme.
Const ID As String = M_c & "standardNormalIC"
' check the input probability
If (prob < 0#) Or (prob > 1#) Then
  Err.Raise InvalidArg_c, ID, _
  "Argument ERROR. Need 0 <= prob <= 1 but prob = " & prob & vbLf & _
  "Problem in " & ID
End If
' Calculate the inverse cumulative distribution function (quantile)
' the approximation used here has equal-ripple absolute error of 6.3E-11 or
' less except for prob > 0.9999997 where granularity in 'prob' dominates.
' The worst relative error of the corresponding PDF is about 2E-8.
' Note that the positive tail has zero accuracy above 'prob' = 1 - 1.1E-16;
' If you want accurate tail values use the negative tail instead.
' see Maple worksheet "InverseCumulativeNormal4.mws"
Dim u As Double
u = prob - 0.5  ' antisymmetric around center
Dim ua As Double
ua = Abs(u)  ' distance from center
Dim u2 As Double
u2 = u * u  ' common factor
Dim z As Double  ' standard normal variate
If ua < 0.37377 Then  ' prob > 0.12623 and prob < 0.87377
  ' Padé fit at center
  z = (1.10320796594653 - (7.77941365082975 - (16.1360412312915 - _
    8.94247760684027 * u2) * u2) * u2) * u / _
    (0.440116302105953 - (3.56442583646134 - (9.15646709284907 - _
    (7.69878138754029 - u2) * u2) * u2) * u2)
ElseIf ua < 0.44286 Then  ' prob > 0.05714 and prob < 0.94286
  ' Padé fit in second zone out
  z = (0.317718558863025 - (2.70051978050927 - (7.20258279324852 - _
    5.82660777818178 * u2) * u2) * u2) * u / _
    (0.126757926972973 - (1.21032688875879 - (3.85234822216469 - _
    (4.38840255884193 - u2) * u2) * u2) * u2)
ElseIf (prob < 4.94065645841247E-324) Or (prob = 1#) Then  ' off ends
  z = Sgn(u) * 38.4674056172733
Else  ' need to use Padé approximants with variable Sqr(Log(P))
  Dim w As Double
  If prob < 0.5 Then
    w = Sqr(-Log(prob))  ' on negative tail
  Else
    w = Sqr(-Log(1# - prob))  ' has roundoff noise > error above 0.9999997
  End If
  If w < 3.769 Then  ' prob or 1-prob > 6.77158141318452E-07
    ' Padé fit in third zone out
    w = (3.40265621744676 + (9.03080228605413 - (6.88823432035713 + _
      (9.47396446577765 + 1.41485388628381 * w) * w) * w) * w) / _
      (1.10738880205572 + (7.00041795498572 + (6.72600088945649 + w) * w) * w)
   ElseIf w < 8.371 Then  ' prob or 1-prob > 3.69321326562547E-31
    ' Padé fit in fourth zone out
     w = (27.5896468790036 + (11.8481686174627 - (37.7133528390963 + _
       (18.6301980539071 + 1.41437483654701 * w) * w) * w) * w) / _
       (10.7729777720728 + (29.1330213184579 + (13.1871785457772 + w) * w) * w)
   Else  ' prob <= 3.69321326562547E-31
    ' Padé fit in fifth zone out (only in negative tail)
     w = (859575101.771399 - (167079541.087701 + (887823598.683122 + _
       (206626688.300811 + 7785160.41001698 * w) * w) * w) * w) / _
       (382471838.745491 + (643197787.097259 + (146148247.666043 + _
       (5504690.97847543 + w) * w) * w) * w)
   End If
   z = Sgn(-u) * w
End If
standardNormalIC = z
End Function

'===============================================================================
Public Function triangularIC( _
  ByVal prob As Double, _
  Optional ByVal peak As Double = 0.5, _
  Optional ByVal loEnd As Double = 0#, _
  Optional ByVal hiEnd As Double = 1#) _
As Double
Attribute triangularIC.VB_Description = "ICDF of a triangular distribution; PDF(x) starts at 0 at 'x' = 'loEnd', rises linearly to 'x' = 'peak', then drops back linearly to 0 at 'x' = 'hiEnd'."
' ICDF of a triangular distribution; PDF(x) starts at 0 at 'x' = 'loEnd', rises
' linearly to 'x' = 'peak', drops back linearly to 0 at 'x' = 'hiEnd'.
' PDF is zero below loEnd and above hiEnd.
' This is a 0-tailed asymmetric (in general) continuous distribution.
' Devised and coded by John Trenholme.
Const ID As String = M_c & "triangularIC"
' check the input probability
If (prob < 0#) Or (prob > 1#) Then
  Err.Raise InvalidArg_c, ID, _
  "Argument ERROR. Need 0 <= prob <= 1 but prob = " & prob & vbLf & _
  "Problem in " & ID
End If
' check the distribution parameters
If (loEnd = hiEnd) Then  ' silently avoid zero-length interval (so span <> 0)
  loEnd = loEnd - 0.000000000000001 * Abs(loEnd) - 1E-290
  hiEnd = hiEnd + 0.00000000000001 * Abs(hiEnd) + 1E-290
End If
Dim span As Double, u As Double
span = hiEnd - loEnd
u = (peak - loEnd) / span  ' goes from 0 to 1 across distribution's base
If (u < 0#) Or (u > 1#) Then
  Err.Raise InvalidArg_c, ID, _
  "Argument ERROR. Need peak between loEnd and hiEnd but" & vbLf & _
  "  loEnd = " & loEnd & vbLf & _
  "  peak = " & peak & vbLf & _
  "  hiEnd = " & hiEnd & vbLf & _
  "Problem in " & ID
End If
' calculate the inverse cumulative distribution function (quantile)
If prob <= u Then
  triangularIC = loEnd + span * Sqr(prob * u)
Else
  triangularIC = hiEnd - span * Sqr((1# - prob) * (1# - u))
End If
End Function

'===============================================================================
Public Function uniformIC( _
  ByVal prob As Double, _
  Optional ByVal loEnd As Double = 0#, _
  Optional ByVal hiEnd As Double = 1#) _
As Double
Attribute uniformIC.VB_Description = "ICDF of a uniform distribution. All this does is translate and scale the input uniform distribution."
' ICDF of a uniform distribution. All this does is translate and scale the input
' probability value.
' This is a 0-tailed symmetric continuous distribution.
' Devised and coded by John Trenholme.
Const ID As String = M_c & "uniformIC"
' check the input probability
If (prob < 0#) Or (prob > 1#) Then
  Err.Raise InvalidArg_c, ID, _
  "Argument ERROR. Need 0 <= prob <= 1 but prob = " & prob & vbLf & _
  "Problem in " & ID
End If
' calculate the inverse cumulative distribution function (quantile)
uniformIC = loEnd + (hiEnd - loEnd) * prob
End Function

'===============================================================================
Public Function uniformLongIC( _
  ByVal prob As Double, _
  Optional ByVal loEnd As Long = 0&, _
  Optional ByVal hiEnd As Long = 1&) _
As Double
Attribute uniformLongIC.VB_Description = "ICDF of a uniform distribution of Long integer values from 'loEnd' to 'hiEnd' (inclusive)."
' ICDF of a uniform distribution of Long integer values from 'loEnd' to 'hiEnd'
' (inclusive). This is a 0-tailed symmetric discrete distribution.
' Devised and coded by John Trenholme.
Const ID As String = M_c & "uniformLongIC"
' check the input probability
If (prob < 0#) Or (prob > 1#) Then
  Err.Raise InvalidArg_c, ID, _
  "Argument ERROR. Need 0 <= prob <= 1 but prob = " & prob & vbLf & _
  "Problem in " & ID
End If
' calculate the inverse cumulative distribution function (quantile)
' note: CLng snaps to closest integer value, hence the "- 0.5"
uniformLongIC = CLng(loEnd - 0.5 + (hiEnd - loEnd + 1#) * prob)
End Function

'===============================================================================
Public Function weibullIC( _
  ByVal prob As Double, _
  Optional ByVal size As Double = 1#, _
  Optional ByVal power As Double = 1#, _
  Optional ByVal loEnd As Double = 0#) _
As Double
Attribute weibullIC.VB_Description = "ICDF of a Weibull distribution with the supplied size and power parameter values, starting at 'loEnd'. If you want a negative-going tail, negate the return value (and perhaps pass 1-P as the argument)."
' ICDF of a Weibull distribution with the supplied size and power parameter
' values. When power = 1, this is an exponential distribution. The PDF is:
'   PDF = power / size * (x / size) ^ (power - 1) * Exp(-(x / size) ^ power)
' If you want a negative-going tail, negate the return value (and perhaps
' pass 1-P as the argument).
' This is a 1-tailed asymmetric continuous distribution.
' Devised and coded by John Trenholme.
Const ID As String = M_c & "weibullIC"
' check the input probability
If (prob < 0#) Or (prob > 1#) Then
  Err.Raise InvalidArg_c, ID, _
  "Argument ERROR. Need 0 <= prob <= 1 but prob = " & prob & vbLf & _
  "Problem in " & ID
End If
' check the distribution parameters
If (size <= 0#) Or (power <= 0#) Then
  Err.Raise InvalidArg_c, ID, _
  "Argument ERROR. Need size > 0 and power > 0 but" & vbLf & _
  "size = " & size & " power = " & power & vbLf & _
  "If you want a negative-going tail, negate the return value." & vbLf & _
  "Problem in " & ID
End If
' calculate the inverse cumulative distribution function (quantile)
If prob = 1# Then
  ' avoid Log(0); use 1# - prob = 2^(-54) (half a bit below 1.0)
  Const LogTiny_c As Double = -37.429947750237  ' Log(2^(-54))=Log(5.551115E-17)
  weibullIC = loEnd + size * (-LogTiny_c) ^ (1# / power)
Else
  ' use of "1-prob" makes this a strictly non-decreasing function
  weibullIC = loEnd + size * (-Log(1# - prob)) ^ (1# / power)
End If
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


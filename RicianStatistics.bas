Attribute VB_Name = "RicianStatistics"
Attribute VB_Description = "Functions related to the Rician intensity distribution. Devised & coded by John Trenholme."
'
'###############################################################################
'#
'#  Visual Basic 6 & VBA code module "RicianStatistics.bas"
'#
'#  Devised and coded by John Trenholme
'#
'#  Exports the routines:
'#    Function riCDF
'#    Function riInvCDF
'#    Function riMode
'#    Function riPDF
'#    Function riPDFN
'#    Function riPDFNv
'#    Function zFromRician
'#
'#  Requires the module "StandardNormal.bas" to supply:
'#    Function snCDF
'#    Function snInvCDF
'#    Function snPDF
'#
'#  Requires the module "BesselI0.bas" to supply:
'#    Function BesselI0
'#
'#  Note: when called from Excel, Err.Raise causes Excel's #VALUE! error
'#
'###############################################################################

Option Explicit
Option Private Module  ' Don't allow visibility outside this Project

Private Const Version_c As String = "2006-06-06"
Private Const m_c As String = "RicianStatistics!"  ' module name + separator

Private Const EOL As String = vbNewLine  ' short form; works on both PC and Mac

Private Const ConGauss_c As Double = 0.000001  ' contrast below this = Gaussian
Private Const ConMin_c As Double = 1E-300      ' contrast can't fall below this

' we keep module-global transform coefficients to avoid needless recalculation
Private con_m As Double  ' previous contrast; will initialize to 0.0
Private a_m As Double    ' u-to-z transform coefficient
Private b_m As Double    ' u-to-z transform coefficient
Private c_m As Double    ' u-to-z transform coefficient
Private d_m As Double    ' u-to-z transform coefficient
Private p_m As Double    ' z-to-u transform coefficient
Private q_m As Double    ' z-to-u transform coefficient
Private r_m As Double    ' z-to-u transform coefficient
Private s_m As Double    ' z-to-u transform coefficient

Private ofi_m As Integer  ' output file index used by unit-test routine

'*******************************************************************************
Public Function riCDF( _
  ByVal x As Double, _
  ByVal mean As Double, _
  ByVal stdDev As Double) _
As Double
Attribute riCDF.VB_Description = "Cumulative Distribution Function of a Rician intensity to 0.0004 worst relative error for CDF values from 0.000001 to 0.999999999, for values of stdDev/mean ('contrast') from 0 to near 0.55."
' Returns an approximation to the cumulative distribution function of a variate
' with the Rician intensity (or "modified Rician") distribution. The maximum
' relative error is of order 0.0004 for CDF values from 0.000001 to 0.999999999,
' for values of stdDev / mean ("contrast") from near 0 to near 0.55.

' Raises error 5 ("Invalid procedure call or argument") for impossible input
' values.

' See the Maple file "WarpRicianExceedance.mws" for calculations.

' Note: CDF of the max of N independent samples is this CDF to the Nth power.

' Note: this function and riInvCDF are only approximate inverses (sorry).

Static calls_s As Double  ' number of times this has been called
calls_s = calls_s + 1#    ' stops adding at 9,007,199,254,740,992 calls

If x < 0# Then
  Err.Raise 5&, "RicianStatistics!riCDF", _
    "Argument error in RicianStatistics!riCDF on call " & _
    Format$(calls_s, "#,##0") & vbLf & vbLf & _
    "x must be >= 0 but got:" & vbLf & _
    "  x = " & x & vbLf & _
    "  mean = " & mean & vbLf & _
    "  stdDev = " & stdDev
ElseIf mean <= 0# Then
  Err.Raise 5&, "RicianStatistics!riCDF", _
    "Argument error in RicianStatistics!riCDF on call " & _
    Format$(calls_s, "#,##0") & vbLf & vbLf & _
    "mean must be > 0 but got:" & vbLf & _
    "  x = " & x & vbLf & _
    "  mean = " & mean & vbLf & _
    "  stdDev = " & stdDev
ElseIf stdDev <= 0# Then
  Err.Raise 5&, "RicianStatistics!riCDF", _
    "Argument error in RicianStatistics!riCDF on call " & _
    Format$(calls_s, "#,##0") & vbLf & vbLf & _
    "stdDev must be > 0 but got:" & vbLf & _
    "  x = " & x & vbLf & _
    "  mean = " & mean & vbLf & _
    "  stdDev = " & stdDev
ElseIf stdDev > mean Then
  Err.Raise 5&, "RicianStatistics!riCDF", _
    "Argument error in RicianStatistics!riCDF on call " & _
    Format$(calls_s, "#,##0") & vbLf & vbLf & _
    "stdDev must be <= mean but got:" & vbLf & _
    "  x = " & x & vbLf & _
    "  mean = " & mean & vbLf & _
    "  stdDev = " & stdDev
ElseIf x = 0# Then
  riCDF = 0#
Else
  Dim con As Double
  con = stdDev / mean
  If con > 0.55 Then  ' caller is beyond numeric fit validity range
    Err.Raise 5&, "RicianStatistics!riCDF", _
      "Result-invalid error in RicianStatistics!riCDF on call " & _
      Format$(calls_s, "#,##0") & vbLf & vbLf & _
      "Approximation valid when stdDev <= 0.55 * mean but got:" & vbLf & _
      "  x = " & x & vbLf & _
      "  mean = " & mean & vbLf & _
      "  stdDev = " & stdDev & vbLf & _
      "             = " & stdDev / mean & " * mean"
  Else  ' all is well, so proceed with the approximation
    Dim u As Double
    u = x / mean
    Dim z As Double
    If con > ConGauss_c Then  ' warp u into standard normal variate z
      If con_m <> con Then Call coeffs(con)
      Dim sqrU As Double
      sqrU = Sqr(u)
      z = a_m + b_m * sqrU + c_m * Log(u) + d_m / sqrU
    Else  ' tiny contrast - use Gaussian approximation
      ' no zero contrast (will give divide-by-zero error)
      If con < ConMin_c Then con = ConMin_c
      z = (u - riMode(con)) / con
    End If
    riCDF = snCDF(z)
  End If
End If
End Function

'*******************************************************************************
Public Function ricianStatisticsVersion() As String
Attribute ricianStatisticsVersion.VB_Description = "The date of the latest revision to this module as a string in the format 'YYYY-MM-DD' such as 2004-06-18. It's a function so Excel etc. can use it."
' The date of the latest revision to this module as a string in the format
' 'YYYY-MM-DD' such as 2004-06-18. It's a function so Excel etc. can use it.
ricianStatisticsVersion = Version_c
End Function

'*******************************************************************************
Public Function riInvCDF( _
  ByVal probability As Double, _
  ByVal mean As Double, _
  ByVal stdDev As Double) _
As Double
Attribute riInvCDF.VB_Description = "Inverse of Cumulative Distribution Function of a Rician intensity to 0.04*mean worst absolute error for probability values from 0.000001 to 0.999999999, for values of stdDev/mean ('contrast') from 0 to near 0.55."
' Returns the Rician intensity argument corresponding to the supplied CDF
' probability.

' Raises error 5 ("Invalid procedure call or argument") for impossible input
' values.

' Note: inverse of max of N independent samples is riInvCDF (p ^ (1# / N), ...)

' Note: this function and riCDF are only approximate inverses (sorry).

Static con_s As Double  ' previous contrast value; will initialize to 0.0

Static calls_s As Double  ' number of times this has been called
calls_s = calls_s + 1#    ' stops adding at 9,007,199,254,740,992 calls

If (probability < 0#) Or (probability > 1#) Then
  Err.Raise 5&, "RicianStatistics!riInvCDF", _
    "Argument error in RicianStatistics!riInvCDF on call " & _
    Format$(calls_s, "#,##0") & vbLf & vbLf & _
    "probability must be >= 0 and <= 1 but got:" & vbLf & _
    "  probability = " & probability & vbLf & _
    "  mean = " & mean & vbLf & _
    "  stdDev = " & stdDev
ElseIf mean <= 0# Then
  Err.Raise 5&, "RicianStatistics!riInvCDF", _
    "Argument error in RicianStatistics!riInvCDF on call " & _
    Format$(calls_s, "#,##0") & vbLf & vbLf & _
    "mean must be > 0 but got:" & vbLf & _
    "  probability = " & probability & vbLf & _
    "  mean = " & mean & vbLf & _
    "  stdDev = " & stdDev
ElseIf stdDev <= 0# Then
  Err.Raise 5&, "RicianStatistics!riInvCDF", _
    "Argument error in RicianStatistics!riInvCDF on call " & _
    Format$(calls_s, "#,##0") & vbLf & vbLf & _
    "stdDev must be > 0 but got:" & vbLf & _
    "  probability = " & probability & vbLf & _
    "  mean = " & mean & vbLf & _
    "  stdDev = " & stdDev
ElseIf stdDev > mean Then
  Err.Raise 5&, "RicianStatistics!riInvCDF", _
    "Argument error in RicianStatistics!riInvCDF on call " & _
    Format$(calls_s, "#,##0") & vbLf & vbLf & _
    "stdDev must be <= mean but got:" & vbLf & _
    "  probability = " & probability & vbLf & _
    "  mean = " & mean & vbLf & _
    "  stdDev = " & stdDev
ElseIf probability = 0# Then
  riInvCDF = 0#
Else
  Dim con As Double
  con = stdDev / mean
  ' no zero contrast (will give divide-by-zero error)
  If con > 0.55 Then  ' caller is beyond numeric fit range; punt
    Err.Raise 5&, "RicianStatistics!riInvCDF", _
      "Result-invalid error in RicianStatistics!riInvCDF on call " & _
      Format$(calls_s, "#,##0") & vbLf & vbLf & _
      "Approximation valid when stdDev <= 0.55 * mean but got:" & vbLf & _
      "  probability = " & probability & vbLf & _
      "  mean = " & mean & vbLf & _
      "  stdDev = " & stdDev & vbLf & _
      "             = " & stdDev / mean & " * mean"
  Else
    If con < ConMin_c Then con = ConMin_c
    Dim z As Double
    z = snInvCDF(probability)  ' Gaussian result
    If con >= ConGauss_c Then
      If con_s <> con Then  ' contrast has changed; recalculate coefficients
        con_s = con
        Dim c2 As Double
        c2 = con * con
        ' See Maple file "WarpRicianExceedance.mws" for origin of these coeff's
        p_m = 1# + c2 * (0.30899 + c2 * 0.228253)
        q_m = con * (0.997275 + c2 * (0.764933 - c2 * (3.78672 - c2 * 6.15794)))
        r_m = c2 / (-3.35964 - c2 * 6.2545)
        s_m = c2 * (0.0149985 - c2 * 0.0146975)
      End If
      Dim u As Double
      u = Exp(z * (q_m + z * (r_m + z * s_m))) / p_m  ' un-warp from z to u
      riInvCDF = mean * u
    Else  ' use Gaussian approximation
      riInvCDF = mean + stdDev * z
    End If
  End If
End If
End Function

'*******************************************************************************
Public Function riMode( _
  ByVal contrast As Double) _
As Double
Attribute riMode.VB_Description = "Mode of unit-mean Rician intensity distribution. Approximation good to 1.14E-7 absolute; much better for low contrast (standard deviation / mean)."
' Returns an approximation to the location of the mode (peak) value of the
' Rician intensity (or "modified Rician") distribution, when the mean value
' is unity. The value depends on the "contrast" (standard deviation divided
' by mean). Found by a Padé approximation to numerically calculated results.
' The maximum absolute error is about 1.14E-7. Error for small contrast values
' is much smaller.

' Raises error 5 ("Invalid procedure call or argument") for impossible input
' value.

Static calls_s As Double  ' number of times this has been called
calls_s = calls_s + 1#    ' stops adding at 9,007,199,254,740,992 calls

If (contrast < 0#) Or (contrast > 1#) Then
  Err.Raise 5&, "RicianStatistics!riMode", _
    "Argument error in RicianStatistics!riMode on call " & _
    Format$(calls_s, "#,##0") & vbLf & vbLf & _
    "contrast must be >= 0 and <= 1 but got:" & vbLf & _
    "  contrast = " & contrast
ElseIf contrast = 0# Then
  riMode = 1#
ElseIf contrast >= 0.86602558704572 Then  ' where approximation goes to zero
  riMode = 0#  ' maximum value is at zero for contrast > Sqr(3)/2
Else
  Dim x As Double
  x = contrast  ' just to have a shorter name
  Dim A As Double
  ' written out this way to avoid VB "Expression too complex" errors
  A = 6.2614730546 - x * (11.542554 - x * (6.4616106851 - x * 1.24278285665))
  A = x * (5.2570415927 - x * (9.97246924215 - x * (5.781874401 + x * A)))
  Dim B As Double
  B = 3.09932214714 - x * (2.4945377228 - x * (0.81282758326 - x * 0.157454873))
  B = x * (10.722473396 - x * (9.7247162059 - x * (1.99947603112 + x * B)))
  riMode = (1# - A) / (1# - x * (5.2570416928 - B))
End If
End Function

'*******************************************************************************
Public Function riPDF( _
  ByVal x As Double, _
  ByVal mean As Double, _
  ByVal stdDev As Double) _
As Double
Attribute riPDF.VB_Description = "Probability Density Function of a Rician intensity to 3.9E-8 worst relative error."
' Returns the probability density function of a variate with the Rician
' intensity (or "modified Rician") distribution. Maximum relative error is
' determined by the I0(x) approximation routine, which is about 3.9E-8 here.

' Raises error 5 ("Invalid procedure call or argument") for impossible input
' values.

Const ExpLimit_c As Double = 700#  ' don't do exp(x) if x bigger than this

Static calls_s As Double  ' number of times this has been called
calls_s = calls_s + 1#    ' stops adding at 9,007,199,254,740,992 calls

If x < 0# Then
  Err.Raise 5&, "RicianStatistics!riPDF", _
    "Argument error in RicianStatistics!riPDF on call " & _
    Format$(calls_s, "#,##0") & vbLf & vbLf & _
    "x must be >= 0 but got:" & vbLf & _
    "  x = " & x & vbLf & _
    "  mean = " & mean & vbLf & _
    "  stdDev = " & stdDev
ElseIf mean <= 0# Then
  Err.Raise 5&, "RicianStatistics!riPDF", _
    "Argument error in RicianStatistics!riPDF on call " & _
    Format$(calls_s, "#,##0") & vbLf & vbLf & _
    "mean must be > 0 but got:" & vbLf & _
    "  x = " & x & vbLf & _
    "  mean = " & mean & vbLf & _
    "  stdDev = " & stdDev
ElseIf stdDev <= 0# Then
  Err.Raise 5&, "RicianStatistics!riPDF", _
    "Argument error in RicianStatistics!riPDF on call " & _
    Format$(calls_s, "#,##0") & vbLf & vbLf & _
    "stdDev must be > 0 but got:" & vbLf & _
    "  x = " & x & vbLf & _
    "  mean = " & mean & vbLf & _
    "  stdDev = " & stdDev
ElseIf stdDev > mean Then
  Err.Raise 5&, "RicianStatistics!riCDF", _
    "Argument error in RicianStatistics!riCDF on call " & _
    Format$(calls_s, "#,##0") & vbLf & vbLf & _
    "stdDev must be <= mean but got:" & vbLf & _
    "  x = " & x & vbLf & _
    "  mean = " & mean & vbLf & _
    "  stdDev = " & stdDev
ElseIf x = 0# Then
  riPDF = 0#
Else
  Dim u As Double
  u = x / mean
  Dim con As Double
  con = stdDev / mean
  If con < ConGauss_c Then  ' calculation of t1 would lose precision
    ' no zero contrast (will give divide-by-zero error)
    If con < ConMin_c Then con = ConMin_c
    ' at low contrast, Gaussian is very close to Rician so use Gaussian
    riPDF = snPDF((u - riMode(con)) / con) / (con * mean)
  Else
    Dim t1 As Double
    t1 = Sqr(1# - con * con)
    Dim t2 As Double
    t2 = 1# / (1# - t1)  ' note that t2 becomes large for small contrast
    Dim A As Double
    A = (u + t1) * t2
    Dim B As Double
    B = 2# * Sqr(u * t1) * t2
    ' We exceed the limits of Exp when con < Sqr((u+1)*(2*E+1-u))/(E+1) where E
    ' is the largest number that gives good results for Exp(-E): about 709.7
    ' This gives con = 0.075 at u = 1
    ' We will exceed the limits of BesselI0 when
    '   con < 2/E^2*Sqr((E^2+2*u)*Sqr(u*(E^2+u))-2*u*(E^2+u))
    ' This also gives con = 0.075 at u = 1
    If (A < ExpLimit_c) And (B < ExpLimit_c) Then  ' no underflow or overflow
      riPDF = Exp(-(u + t1) * t2) * besselI0(2# * Sqr(u * t1) * t2) * t2 / mean
    Else  ' Exp() will underflow, or BesselI0 will overflow (or both)
      ' use PDF we get from approximate CDF
      riPDF = riPDFfromCDF(x, mean, stdDev)
    End If
  End If
End If
End Function

'*******************************************************************************
Public Function riPDFfromCDF( _
  ByVal x As Double, _
  ByVal mean As Double, _
  ByVal stdDev As Double) _
As Double
Attribute riPDFfromCDF.VB_Description = "Probability Density Function of a variate with the Rician intensity (or 'modified Rician') distribution, derived from the derivative of the approximate Rician CDF (as returned by riCDF)."
' Returns the probability density function of a variate with the Rician
' intensity (or "modified Rician") distribution, as derived from the derivative
' of the approximate Rician CDF (as returned by riCDF).
' This function is not as accurate as the result of riPDF, but it works over a
' broader range, and is used when a PDF consistent with the approximate CDF is
' needed.

' Raises error 5 ("Invalid procedure call or argument") for impossible input
' values.

Static calls_s As Double  ' number of times this has been called
calls_s = calls_s + 1#    ' stops adding at 9,007,199,254,740,992 calls

If x < 0# Then
  Err.Raise 5&, "RicianStatistics!riPDFfromCDF", _
    "Argument error in RicianStatistics!riPDFfromCDF on call " & _
    Format$(calls_s, "#,##0") & vbLf & vbLf & _
    "x must be >= 0 but got:" & vbLf & _
    "  x = " & x & vbLf & _
    "  mean = " & mean & vbLf & _
    "  stdDev = " & stdDev
ElseIf mean <= 0# Then
  Err.Raise 5&, "RicianStatistics!riPDFfromCDF", _
    "Argument error in RicianStatistics!riPDFfromCDF on call " & _
    Format$(calls_s, "#,##0") & vbLf & vbLf & _
    "mean must be > 0 but got:" & vbLf & _
    "  x = " & x & vbLf & _
    "  mean = " & mean & vbLf & _
    "  stdDev = " & stdDev
ElseIf stdDev <= 0# Then
  Err.Raise 5&, "RicianStatistics!riPDFfromCDF", _
    "Argument error in RicianStatistics!riPDFfromCDF on call " & _
    Format$(calls_s, "#,##0") & vbLf & vbLf & _
    "stdDev must be > 0 but got:" & vbLf & _
    "  x = " & x & vbLf & _
    "  mean = " & mean & vbLf & _
    "  stdDev = " & stdDev
ElseIf stdDev > mean Then
  Err.Raise 5&, "RicianStatistics!riPDFfromCDF", _
    "Argument error in RicianStatistics!riPDFfromCDF on call " & _
    Format$(calls_s, "#,##0") & vbLf & vbLf & _
    "stdDev must be <= mean but got:" & vbLf & _
    "  x = " & x & vbLf & _
    "  mean = " & mean & vbLf & _
    "  stdDev = " & stdDev
ElseIf x = 0# Then
  riPDFfromCDF = 0#
Else
  Dim u As Double
  u = x / mean
  Dim con As Double
  con = stdDev / mean
  If con_m <> con Then Call coeffs(con)
  Dim sqrU As Double
  sqrU = Sqr(u)
  Dim z As Double
  z = a_m + b_m * sqrU + c_m * Log(u) + d_m / sqrU
  riPDFfromCDF = snPDF(z) * ((b_m - d_m / u) / (2# * sqrU) + c_m / u) / mean
End If
End Function

'*******************************************************************************
Public Function riPDFN( _
  ByVal x As Double, _
  ByVal mean As Double, _
  ByVal stdDev As Double, _
  ByVal N As Long) _
As Double
Attribute riPDFN.VB_Description = "Probability Density Function the maximum of N independent Rician intensities to 0.01 worst relative error for x/mean from 0.4 to 6, for stdDev/mean ('contrast') from 0 to near 0.55 (better for lower contrast)."
' Returns the probability density function of the maximum of N independent
' samples of a variate with the Rician intensity (or "modified Rician")
' distribution. Uses the fact that the CDF for N is the CDF for 1 to the N'th
' power, then takes the derivative of this relationship. Good to 1% relative
' error for x/mean from 0.4 to 6, for stdDev/mean ("contrast") from 0 to 0.55
' (better for lower contrast).

' See the Maple file "WarpRicianExceedance.mws"

Static calls_s As Double  ' number of times this has been called
calls_s = calls_s + 1#    ' stops adding at 9,007,199,254,740,992 calls

If N <= 0& Then  ' caller is confused and has supplied impossible input; punt
  Err.Raise 5&, "RicianStatistics!riPDFN", _
    "Argument error in RicianStatistics!riPDFN on call " & _
    Format$(calls_s, "#,##0") & vbLf & vbLf & _
    "N must be > 0 but got:" & vbLf & _
    "  x = " & x & vbLf & _
    "  mean = " & mean & vbLf & _
    "  stdDev = " & stdDev & vbLf & _
    "  N = " & N
ElseIf N = 1& Then  ' easy case; use the simpler 1-sample function
  riPDFN = riPDF(x, mean, stdDev)
Else
  If x < 0# Then
    Err.Raise 5&, "RicianStatistics!riPDFN", _
      "Argument error in RicianStatistics!riPDFN on call " & _
      Format$(calls_s, "#,##0") & vbLf & vbLf & _
      "x must be >= 0 but got:" & vbLf & _
      "  x = " & x & vbLf & _
      "  stdDev = " & stdDev & vbLf & _
      "  N = " & N
  ElseIf mean <= 0# Then
    Err.Raise 5&, "RicianStatistics!riPDFN", _
      "Argument error in RicianStatistics!riPDFN on call " & _
      Format$(calls_s, "#,##0") & vbLf & vbLf & _
      "mean must be > 0 but got:" & vbLf & _
      "  x = " & x & vbLf & _
      "  mean = " & mean & vbLf & _
      "  stdDev = " & stdDev & vbLf & _
      "  N = " & N
  ElseIf stdDev <= 0# Then
    Err.Raise 5&, "RicianStatistics!riPDFN", _
      "Argument error in RicianStatistics!riPDFN on call " & _
      Format$(calls_s, "#,##0") & vbLf & vbLf & _
      "stdDev must be > 0 but got:" & vbLf & _
      "  x = " & x & vbLf & _
      "  mean = " & mean & vbLf & _
      "  stdDev = " & stdDev & vbLf & _
      "  N = " & N
  ElseIf stdDev > mean Then
    Err.Raise 5&, "RicianStatistics!riPDFN", _
      "Argument error in RicianStatistics!riPDFN on call " & _
      Format$(calls_s, "#,##0") & vbLf & vbLf & _
      "stdDev must be <= mean but got:" & vbLf & _
      "  x = " & x & vbLf & _
      "  mean = " & mean & vbLf & _
      "  stdDev = " & stdDev & vbLf & _
      "  N = " & N
  ElseIf x = 0# Then
    riPDFN = 0#
  Else
    ' note pdfN = N * CDF ^ (N - 1) * PDF; use PDF we get from approximate CDF
    riPDFN = N * riCDF(x, mean, stdDev) ^ (N - 1) * _
      riPDFfromCDF(x, mean, stdDev)
    End If
End If
End Function

'*******************************************************************************
Public Function riPDFNv( _
  ByVal x As Double, _
  ByVal mean As Double, _
  ByVal stdDev As Double, _
  ByVal N As Long, _
  ByVal vary As Double) _
As Double
Attribute riPDFNv.VB_Description = "Probability Density Function of the maximum of N Rician intensities with fraction 'vary' of noise variance independent from sample to sample, to a few percent worst relative error for x/mean from 0.4 to 6, for stdDev/mean ('contrast') from 0 to near 0.55."
' Returns the probability density function of the maximum of N independent
' samples of a variate with the Rician intensity (or "modified Rician")
' distribution that results from having a portion "vary" of the total variance
' that changes from sample to sample, and a portion (1-vary) that is fixed from
' sample to sample. This corresponds to a convolution of the fixed part (which
' has a single-sample Rician intensity PDF) with the varying part (which has
' a maximum-of-N-sample Rician intensity distribution).

Const c_Mmax As Long = 33
Const c_Mmin As Long = 2
Const c_Mfact As Double = c_Mmax - c_Mmin
Const c_shift As Double = 0.15

Dim dx As Double
Dim f2 As Double
Dim j As Long
Dim invN As Double
Dim lnN As Double
Dim M As Long
Dim mF As Double
Dim mV As Double
Dim prob As Double
Dim p0 As Double
Dim p1 As Double
Dim s2 As Double
Dim sdE As Double
Dim sdF As Double
Dim sdN As Double
Dim sdV As Double
Dim sf As Double
Dim sv As Double
Dim sum As Double
Dim tilt As Double
Dim xA As Double
Dim xB As Double
Dim xC As Double
Dim xD As Double
Dim xF As Double
Dim xV As Double

If (x < 0#) Or (mean <= 0#) Or (stdDev <= 0#) Or (stdDev > mean) Or (N < 1) Then
  ' caller is confused and has supplied impossible input; punt
  riPDFNv = 0#
Else
  If (vary <= 0#) Or (N = 1) Then
    ' if "vary" is zero (or less), or N = 1, we get the 1-sample result from
    ' the input mean and standard deviation no matter what N is
    riPDFNv = riPDF(x, mean, stdDev)
  ElseIf vary >= 1# Then
    ' if "vary" is unity (or more), we get the maximum-of-N-independent-samples
    ' distribution resulting from the input mean and standard deviation
    riPDFNv = riPDFN(x, mean, stdDev, N)
  Else  ' 0.0 < vary < 1.0, so we have to do the hard work
    ' determine which distribution is the narrower one (approximately)
        
    ' get mean & standard deviation of fixed part
    f2 = Sqr(mean * mean - stdDev * stdDev)
    sdE = Sqr((mean - f2) * 0.5)
    s2 = sdE * sdE
    mF = mean - 2# * vary * s2
    sdF = 2# * sdE * Sqr((1# - vary) * (f2 + s2 * (1# - vary)))
    
    ' get standard deviation of varying part, based on mean of fixed part
    sdV = 2# * sdE * Sqr(vary * (mF + s2 * vary))
    
    ' get the standard deviation of the varying-part max-of-N by using the fact
    ' that it is 2.36 standard deviations from the 12.1% point to the 87.9%
    ' point of a standard normal cumulative distribution (points selected to
    ' span FWHM of standard normal PDF)
    invN = 1# / N
    sdN = (riInvCDF(0.879 ^ invN, mean, sdV) _
           - riInvCDF(0.121 ^ invN, mean, sdV)) / 2.36
    
    ' we approximate the convolution by using a sum of distributions
    ' pick a number M of points based on fractiles of the narrower
    ' distribution, and add up wider distributions based on those points
    sum = 0#
    If sdF < 2.5 * sdN Then
      ' fixed part is narrower, or not "too much" wider; use it as base
      M = c_Mmin + CInt(c_Mfact * sdF / sdN)
      If M > c_Mmax Then M = c_Mmax
      ' probability fractional point coefficients, shifted a bit near the ends
      p0 = (1# - c_shift) / (2# * M) - 1# / (M - c_shift)
      p1 = 1# / (M - c_shift)
      For j = 1 To M
        prob = p0 + p1 * j
        xF = riInvCDF(prob, mF, sdF)
        sum = sum + riPDFN(x, xF + 2# * vary * s2, _
                           2# * sdE * Sqr(vary * (xF + vary * s2)), N)
      Next j
    Else
      ' varying part is much narrower; use it as base
      M = c_Mmin + CInt(c_Mfact * sdN / sdF)
      If M > c_Mmax Then M = c_Mmax
      ' probability fractional point coefficients, shifted a bit near the ends
      p0 = (1# - c_shift) / (2# * M) - 1# / (M - c_shift)
      p1 = 1# / (M - c_shift)
      ' apply a kludge factor to allow for the fact that we are doing
      ' the convolution "backwards"
      xA = riInvCDF(0.1, mF, sdF)
      xB = riInvCDF(0.9, mF, sdF)
      xC = riInvCDF(0.5 ^ invN, xA + 2# * vary * s2, _
                    2# * sdE * Sqr(vary * (xA + vary * s2)))
      xD = riInvCDF(0.5 ^ invN, xB + 2# * vary * s2, _
                    2# * sdE * Sqr(vary * (xB + vary * s2)))
      tilt = (xD - xC) / (xB - xA)  ' slope in 2D space before collapse
      tilt = 1# + 0.4 * (tilt - 1#)
      For j = 1 To M
        prob = p0 + p1 * j
        xV = riInvCDF(prob ^ invN, mF, sdV)
        sum = sum + riPDF(x, xV, tilt * Sqr(sdF * sdF + sdN * sdN))
      Next j
    End If
    riPDFNv = sum / M
  End If
End If
End Function

'*******************************************************************************
Public Function zFromRician(ByVal x As Double, _
                            ByVal mean As Double, _
                            ByVal stdDev As Double) _
  As Double
Attribute zFromRician.VB_Description = "Coordinate that results from mapping a variate with the Rician intensity distribution to the standard normal distribution. When used to find the CDF, the maximum relative error is 0.0004 or better for CDF values from 0.000001 to 0.999999999, for values of stdDev/mean ('contrast') from 0 to near 0.55. Absolute error in z varies from 1 around z = -4 with contrast 0.5, to better than 0.0001 for z > 1 at any contrast."
' Returns an approximation to the coordinate that results from mapping a variate
' with the Rician intensity (or "modified Rician") distribution to the standard
' normal distribution. When used to find the exceedance, the maximum relative
' error is 0.0004 or better for exceedance values from 1.0 to 1E-9, for values
' of stdDev / mean ("contrast") from 0 to near 0.55. Absolute error in z varies
' from 1 around z = -4 with contrast 0.5, to better than 0.0001 for z > 1 and
' any contrast.

' See the Maple file "WarpRicianExceedance.mws"

' Version of 24 Jul 2002 - John Trenholme

Dim con As Double
Dim sqrU As Double
Dim u As Double

If (x < 0#) Or (mean <= 0#) Or (stdDev <= 0#) Or (stdDev > mean) Then
  ' caller is confused and has supplied impossible input; punt
  zFromRician = -37.7
Else
  con = stdDev / mean
  If (con < 0.01) Or (con > 0.55) Then  ' caller is overly ambitious; punt
    zFromRician = -37.7
  Else  ' all is well, so proceed with the approximation
    If con_m <> con Then Call coeffs(con)
    u = x / mean
    sqrU = Sqr(u)
    zFromRician = a_m + b_m * sqrU + c_m * 0.0560801 * Log(u) + d_m / sqrU
  End If
End If
End Function

'===============================================================================
Private Sub coeffs(ByVal con As Double)
Attribute coeffs.VB_Description = "Internal routine to evaluate approximation coefficients if contrast changes."
' Find transform coefficients from Rician intensity to standard normal and store
' them in module-global variables. Call this only if the contrast changes, to
' avoid needless work.

' no zero contrast (will give divide-by-zero error)
If con < ConMin_c Then con = ConMin_c
' save the contrast that these coefficients correspond to
con_m = con
' re-evaluate the transform coefficients; start with common factor
Dim c2 As Double
c2 = con * con
' See Maple file "WarpRicianExceedance.mws" for origin of these coefficients
a_m = (-2# + c2 * (0.563534 + c2 * (0.0638227 + c2 * 0.186412))) / con
b_m = (2# - c2 * (0.272642 + c2 * (0.0236441 + c2 * 0.112908))) / con
c_m = con * 0.0560801
d_m = con * (-0.0415165 + c2 * 0.00439764)
End Sub

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'&
'& Unit test
'&
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

' set this to "True" to use unit test routines
' set it to "False" to avoid compiling unit test routines into code
#If True Then
' #If False Then

#Const VBA = True  ' set True in Excel (etc.) VBA project; False in VB6

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Public Sub Test_RicianStatistics()
Attribute Test_RicianStatistics.VB_Description = "Unit-test routine for this module. Sends results to a file and to the immediate window (if in IDE)."
' Main unit test routine for this module.

' To run the test from VB, enter this routine's name (above) in the Immediate
' window (if the Immediate window is not open, use View.. or Ctrl-G to open it).
' To run this routine from VBA, put the cursor somewhere in it and hit F5.
' To run this routine from somewhere in a code, call it.

' The output will be in the file 'Test_RicianStatistics.txt' on disk, and in the
' immediate window (use Ctrl-G to open if not visible) if in the VB[A] editor.

' Version of 3 Sep 2002 - John Trenholme

Dim limit As Double
Dim nWarn As Long
Dim worst As Double

' get path to current directory, and prepend to file name
Dim ofs As String
Dim path As String
#If VBA Then
' note: in Excel, save workbook at least once so path exists
  path = Excel.Workbooks(1).path
  If path = "" Then
    MsgBox "Warning! Workbook has no disk location!" & EOL & _
           "Save workbook to disk before proceeding because" & EOL & _
           "Unit-test routine needs a known location to write to." & EOL & _
           "No unit test carried out.", _
           vbOKOnly Or vbCritical, m_c & " Unit Test"
    Exit Sub
  End If
#Else
  ' note: this is the project folder if in VB6 IDE; EXE folder if stand-alone
  path = App.path
#End If
If Right$(path, 1) <> "\" Then path = path & "\"  ' only C:\ etc. have "\"
ofs = path & "Test_" & Left$(m_c, Len(m_c) - 1&) & ".txt"

ofi_m = FreeFile
On Error Resume Next
Open ofs For Output As #ofi_m  ' output file
If Err.Number <> 0 Then
  ofi_m = 0  ' file did not open - don't use it
  MsgBox "ERROR - unable to open output file:" & EOL & EOL & _
    """" & ofs & """" & EOL & EOL & _
    "No unit test output will be written to file.", _
    vbOKOnly Or vbExclamation, Left$(m_c, Len(m_c) - 1&) & " Unit Test"
End If
On Error GoTo 0

teeOut "######## Test of " & Left$(m_c, Len(m_c) - 1&) & " routines at " & Now()

nWarn = 0

' riCDF
worst = 0#
compareRel "riCDF(0.87428,1.0,0.2)", riCDF(0.87428, 1#, 0.2), _
  1# - 0.72498668888586, worst
compareRel "riCDF(0.96732*2,2.0,0.2*2)", riCDF(0.96732 * 2#, 2#, 0.2 * 2#), _
  1# - 0.54535041349896, worst
compareRel "1 - riCDF(2.4098,1.0,0.2)", 1# - riCDF(2.4098, 1#, 0.2), _
  1.377453921085E-08, worst
teeOut "Largest riCDF relative error: " & Format(worst, "0.000000E-0")
limit = 0.00031
If Abs(worst) > limit Then
  teeOut "WARNING! That's too large - should be less than " & _
         Format(limit, "0.0000E-0")
  nWarn = nWarn + 1
End If
teeOut

' riInvCDF
worst = 0#
compareAbs "riInvCDF(0.275013,1.0,0.2)", riInvCDF(0.275013, 1#, 0.2), _
  0.87428, worst
compareAbs "riInvCDF(0.454650,2.0,0.2*2)", riInvCDF(0.45465, 2#, 0.2 * 2#), _
  0.96732 * 2#, worst
compareAbs "riInvCDF(0.999999986225,1.0,0.2)", _
  riInvCDF(0.999999986225, 1#, 0.2), 2.4098, worst
teeOut "Largest riInvCDF absolute error: " & Format(worst, "0.000000E-0")
limit = 0.0055
If Abs(worst) > limit Then
  teeOut "WARNING! That's too large - should be less than " & _
         Format(limit, "0.0000E-0")
  nWarn = nWarn + 1
End If
teeOut

' riPDF
worst = 0#
compareRel "riPDF(0.87428,1.0,0.2)", riPDF(0.87428, 1#, 0.2), _
  1.780258565632, worst
compareRel "riPDF(0.96732*2,2.0,0.2*2)", riPDF(0.96732 * 2#, 2#, 0.2 * 2#), _
  1.0050228968412, worst
compareRel "riPDF(2.4098,1.0,0.3)", riPDF(2.4098, 1#, 0.3), _
  8.0298597082478E-04, worst
teeOut "Largest riPDF relative error: " & Format(worst, "0.000000E-0")
limit = 0.000000039
If Abs(worst) > limit Then
  teeOut "WARNING! That's too large - should be less than " & _
         Format(limit, "0.0000E-0")
  nWarn = nWarn + 1
End If
teeOut

' riPDFN
worst = 0#
' special case
compareRel "riPDFN(0.87428,1.0,0.2,1)", riPDFN(0.87428, 1#, 0.2, 1&), _
  1.7802586, worst  ' reversion to 1-sample routine for N = 1
' regular cases
compareRel "riPDFN(1.293,1.0,0.2,10)", riPDFN(1.293, 1#, 0.2, 10&), _
  3.0786381, worst
compareRel "riPDFN(1.2574,1.0,0.1,100)", riPDFN(1.2574, 1#, 0.1, 100&), _
  8.6933868, worst
compareRel "riPDFN(1.8049*2,2.0,0.2*2,1000)", _
  riPDFN(1.8049 * 2#, 2#, 0.2 * 2#, 1000&), 2.7375608 / 2#, worst
teeOut "Largest riPDFN relative error: " & Format(worst, "0.000000E-0")
limit = 0.0001
If Abs(worst) > limit Then
  teeOut "WARNING! That's too large - should be less than " & _
         Format(limit, "0.0000E-0")
  nWarn = nWarn + 1
End If
teeOut

' riPDFNv
worst = 0#
' special cases
compareRel "riPDFNv(1.0,1.0,0.3,1,1)", riPDFNv(1#, 1#, 0.3, 1&, 1#), _
  riPDF(1#, 1#, 0.3), worst  ' reversion to 1-sample routine for N = 1
compareRel "riPDFNv(2.0,1.0,0.2,10,0.0)", riPDFNv(2#, 1#, 0.2, 10&, 0#), _
  riPDF(2#, 1#, 0.2), worst  ' reversion to 1-sample routine for vary = 0.0
compareRel "riPDFNv(2.0,1.0,0.1,20,1.0)", riPDFNv(2#, 1#, 0.1, 20&, 1#), _
  riPDFN(2#, 1#, 0.1, 20&), worst  ' reversion to simpler routine for vary = 1.0
' regular cases
' note: these are a sanity check only, for use when code is changed, because the
' only test values available are from Monte Carlo results, which were used to
' give some assurance that the results are correct
teeOut "Sanity check only:"
compareRel "riPDFNv(1.2,1.0,0.26,100,0.1)", _
  riPDFNv(1.2, 1#, 0.26, 100&, 0.1), 1.52402, worst
compareRel "riPDFNv(0.825,0.5,0.15,300,0.5)", _
  riPDFNv(0.825, 0.5, 0.15, 300&, 0.5), 2.66983, worst
compareRel "riPDFNv(1.4,1.0,0.13,1000,0.9)", _
  riPDFNv(1.4, 1#, 0.13, 1000&, 0.9), 5.36591, worst
teeOut "Largest riPDFNv relative error: " & Format(worst, "0.000000E-0")
limit = 0.0000013
If Abs(worst) > limit Then
  teeOut "WARNING! That's too large - should be less than " & _
         Format(limit, "0.0000E-0")
  nWarn = nWarn + 1
End If
teeOut

' zFromRician
worst = 0#
compareAbs "zFromRician(1.0138,1.0,0.2)", _
  zFromRician(1.0138, 1#, 0.2), 0.1190136929, worst
compareAbs "zFromRician(1.2*2,2.0,0.2*2)", _
  zFromRician(1.2 * 2, 2#, 0.2 * 2#), 1.002325221, worst
compareAbs "zFromRician(1.3253,1.0,0.1)", _
  zFromRician(1.3253, 1#, 0.1), 3.04732109, worst
compareAbs "zFromRician(3.5919,1.0,0.3)", _
  zFromRician(3.5919, 1#, 0.3), 5.997760738, worst
teeOut "Largest zFromRician absolute error: " & Format(worst, "0.000000E-0")
limit = 0.021
If Abs(worst) > limit Then
  teeOut "WARNING! That's too large - should be less than " & _
         Format(limit, "0.0000E-0")
  nWarn = nWarn + 1
End If
teeOut

If nWarn = 0 Then
  teeOut "Success - all errors were within limits."
Else
  teeOut "FAILURE! - warning count: " & nWarn
End If
teeOut "--- Test complete ---"
Close #ofi_m
End Sub

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Private Sub compareAbs(ByVal str As String, _
                       ByVal approx As Double, _
                       ByVal exact As Double, _
                       ByRef worst As Double)
' unit test support routine - John Trenholme - 9 Jul 2002

Dim absErr As Double

absErr = approx - exact
If Abs(worst) < Abs(absErr) Then worst = absErr
teeOut str & "  approx " & Format(approx, "0.000000000E-0") & _
       "  exact " & Format(exact, "0.000000000E-0") & _
       "  absErr " & Format(absErr, "0.000E-0")
End Sub

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Private Sub compareRel(ByVal str As String, _
                       ByVal approx As Double, _
                       ByVal exact As Double, _
                       ByRef worst As Double)
' unit test support routine - John Trenholme - 25 Jul 2002

Dim relErr As Double

If exact <> 0# Then
  relErr = approx / exact - 1#
Else
  If approx = 0# Then
    relErr = 0#
  Else
    relErr = 1000#  ' an arbitrary large value
  End If
End If
If Abs(worst) < Abs(relErr) Then worst = relErr
teeOut str & "  approx " & Format(approx, "0.000000000E-0") & _
       "  exact " & Format(exact, "0.000000000E-0") & _
       "  relErr " & Format(relErr, "0.000E-0")
End Sub

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Private Sub teeOut(Optional ByRef str As String = "")
' Unit-test support routine. Sends 'str' to Immediate window (if in IDE) and to
' output file (if open).
' John Trenholme - 2006-07-20

Debug.Print str  ' works only if in VB[6,A] IDE editor environment
If ofi_m <> 0 Then Print #ofi_m, str
End Sub

#End If

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


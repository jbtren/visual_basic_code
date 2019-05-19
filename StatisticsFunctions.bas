Attribute VB_Name = "StatisticsFunctions"
'
'###############################################################################
'#
'# Statistical Function Routines. Visual Basic module "StatisticsFunctions.bas"
'#
'#  Exports the routines:
'#    Function binomialBandHi
'#    Function binomialBandHi_Test
'#    Function binomialBandLo
'#    Function binomialBandLo_Test
'#    Function logNormalPDF
'#    Function logNormalPDF_Test
'#
'#  Requires the module "StandardNormal.bas" to supply:
'#    Function snInvCDF
'#
'# Devised and coded by John Trenholme - initial version 15 May 2003
'#
'# Revisions to 20 May 2003
'#
'###############################################################################
'
Option Explicit

'===== binomialBandHi ==========================================================
' Upper extent of a confidence band for a binomial variate.
'
' An unknown but fixed fraction of a large population has some trait. Suppose we
' draw a random sample of size N from the population, and observe that J items in
' the sample have the trait. The expectation value of the fraction is J / N. The
' confidence band, within which the actual value of the fraction lies with
' probability 1 - error, extends from a value below J / N to a value above J / N.
' Call this routine with the observed number that "have" the trait, the "total"
' sample count, plus the "error" probability you are willing to accept that the
' true mean is outside the band. You get back the amount which should be added to
' the expected mean J / N to get the upper end of the confidence band.
'
' You can supply non-integer values for "have" and "total" if needed, although
' most uses will involve integer inputs. The confidence band is only an
' approximation, because of the discrete nature of the binomial distribution.
' It is based on a normal approximation to the binomial, so it is not especially
' good for very small J.
'
' Note: to get "1-sigma" type bands, use error = 0.31731 so that the probalility
' of being inside the band limits is 68.269%, just as for 1-sigma normal limits.
'
' The lower end of the band is given by calling binomialBandLo with the same
' arguments used with this function.
'
' Revisions to 20 May 2003
'
Public Function binomialBandHi(ByRef have As Double, _
                               ByRef total As Double, _
                               ByRef error As Double)
Dim z As Double, z2 As Double
If have >= 0# And have < total And total > 0# And error > 0# And error < 1# Then
  ' find the standard normal variate that corresponds to the given error
  ' split the error equally between two tails of the normal distribution
  z = snInvCDF(1# - 0.5 * error)
  z2 = z * z
  z = 2# * have + z2 + z * Sqr(4# * have * (1# - have / total) + z2)
  binomialBandHi = z / (2# * (total + z2)) - have / total
Else  ' insane input, so silently supply default value
  binomialBandHi = 0#
End If
End Function

'===== binomialBandHi_Test =====================================================
' This is a simple test at a few points to be sure the calculations are correct.
' Because it uses Debug.Print, it only works in the VB or VBA Program Editor.
' Output goes to Immediate window (Ctrl-G to open).
' In VBA: place cursor in this routine's body, press F5.
' In VB: type this routine's name in the Immediate window, and press Enter.
Public Sub binomialBandHi_Test()
Debug.Print "=== Quick tests of binomialBandHi function: "; Date; time
Debug.Print "At 0, 1, 0.05 want 0.793471842031731 get"; _
  binomialBandHi(0#, 1#, 0.05)
Debug.Print "At 1, 1, 0.0.05 want 0 get"; binomialBandHi(1#, 1#, 0.05)
Debug.Print "At 10, 100, 0.31731 want 3.40757853534393E-02 get"; _
  binomialBandHi(10#, 100#, 0.31731)
Debug.Print "--- binomialBandHi test done ---"
End Sub

'===== binomialBandLo ==========================================================
' Lower extent of a confidence band for a binomial variate. See binomialBandHi.
Public Function binomialBandLo(ByRef have As Double, _
                               ByRef total As Double, _
                               ByRef error As Double)
Dim z As Double, z2 As Double
If have > 0# And have <= total And total > 0# And error > 0# And error < 1# Then
  ' find the standard normal variate that corresponds to the given error
  ' split the error equally between two tails of the normal distribution
  z = snInvCDF(1# - 0.5 * error)
  z2 = z * z
  z = 2# * have + z2 - z * Sqr(4# * have * (1# - have / total) + z2)
  binomialBandLo = have / total - z / (2# * (total + z2))
Else  ' insane input, so silently supply default value
  binomialBandLo = 0#
End If
End Function

'===== binomialBandLo_Test =====================================================
' This is a simple test at a few points to be sure the calculations are correct.
' Because it uses Debug.Print, it only works in the VB or VBA Program Editor.
' Output goes to Immediate window (Ctrl-G to open).
' In VBA: place cursor in this routine's body, press F5.
' In VB: type this routine's name in the Immediate window, and press Enter.
Public Sub binomialBandLo_Test()
Debug.Print "=== Quick tests of binomialBandLo function: "; Date; time
Debug.Print "At 0, 1, 0.05 want 0 get"; binomialBandLo(0#, 1#, 0.05)
Debug.Print "At 1, 1, 0.0.05 want 0.793471842031731 get"; _
  binomialBandLo(1#, 1#, 0.05)
Debug.Print "At 90, 100, 0.31731 want 3.40757853534393E-02 get"; _
  binomialBandLo(90#, 100#, 0.31731)
Debug.Print "--- binomialBandHi test done ---"
End Sub

'===== logNormalPDF ============================================================
' Probability density function of the log-normal distribution at a specified
' point 'x' given the distribution's mean and standard deviation. Arguments
' should all be positive.
'
' Note that the user-supplied mean and standard deviation are those of the
' actual log-normal distribution, not (as some foolish authors use) those of
' the Gaussian in the exponent.
'
' Revisions to 19 May 2003
'
Public Function logNormalPDF(ByRef x As Double, _
                             ByRef mean As Double, _
                             ByRef stdDev As Double) _
  As Double
' numeric constants must be written as sums to maintain full accuracy in files
Const Pi As Double = 3.1415926 + 5.35897932E-08
Dim t1 As Double, t2 As Double
If x <= 0# Or mean <= 0# Or stdDev <= 0# Then
  ' caller is befuddled, so silently return a "sensible" default value
  logNormalPDF = 0#
Else
  t2 = stdDev / mean
  t1 = 1# + t2 * t2
  t2 = Log(t1)
  logNormalPDF = Exp(-((Log(x * Sqr(t1) / mean)) ^ 2 / (2# * t2))) / _
    (x * Sqr(2# * Pi * t2))
End If
End Function

'===== logNormalPDF_Test =======================================================
' This is a simple test at a few points to be sure the calculations are correct.
' Because it uses Debug.Print, it only works in the VB or VBA Program Editor.
' Output goes to Immediate window (Ctrl-G to open).
' In VBA: place cursor in this routine's body, press F5.
' In VB: type this routine's name in the Immediate window, and press Enter.
Public Sub logNormalPDF_Test()
Debug.Print "=== Quick tests of logNormalPDF function: "; Date; time
Debug.Print "At 0, 1, 1 want 0 get"; logNormalPDF(0#, 1#, 1#)
Debug.Print "At 1, 1, 1 want 0.439408633656720 get"; logNormalPDF(1#, 1#, 1#)
Debug.Print "At 2.1, 1.9, 0.6 want 0.549440555382819 get"; _
  logNormalPDF(2.1, 1.9, 0.6)
Debug.Print "--- abnormal input ---"
Debug.Print "At 1, 0, 1 want 0 get"; logNormalPDF(1#, 0#, 1#)
Debug.Print "At 1, 1, 0 want 0 get"; logNormalPDF(1#, 1#, 0#)
Debug.Print "At -1, 1, 1 want 0 get"; logNormalPDF(-1#, 1#, 1#)
Debug.Print "At 1, -1, 1 want 0 get"; logNormalPDF(1#, -1#, 1#)
Debug.Print "At 1, 1, -1 want 0 get"; logNormalPDF(1#, 1#, -1#)
Debug.Print "--- logNormalPDF test done ---"
End Sub

'-------------------------------- end of file ----------------------------------

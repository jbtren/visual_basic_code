Attribute VB_Name = "LogNormal"
Attribute VB_Description = "Routines supporting use of the log-normal distribution. Devised & coded by John Trenholme."
'
'###############################################################################
'#
'# Visual Basic 6 & VBA code module "LogNormal.bas"
'#
'# Functions connected with the log-normal distribution.
'#
'# Exports the routines:
'#   Function logNormalCDF
'#   Function logNormalPDF
'#   Sub Test_LogNormal (if UnitTest is True)
'#
'# Requires the module "ErfErfc.bas" to supply:
'#   Function erfcT
'#
'# Devised and coded by John Trenholme
'# Started: 2006-06-12
'#
'###############################################################################

Option Base 0
Option Compare Binary
Option Explicit
Option Private Module  ' Don't allow visibility outside this Project

Private Const Version_c As String = "2006-07-19"
Private Const m_c As String = "LogNormal"  ' module name

#Const UnitTest = True  ' set True to enable unit test code

#If UnitTest Then
#Const VBA = True       ' set True in Excel (etc.) VBA ; False in VB6
Private ofi_m As Integer  ' output file index used by unit-test routines
#End If

Private Const EOL As String = vbNewLine  ' short form; works on both PC and Mac

'===============================================================================
Public Function logNormalCDF( _
  ByVal x As Double, _
  ByVal mean As Double, _
  ByVal stdDev As Double) _
As Double
Attribute logNormalCDF.VB_Description = "Cumulative Distribution Function of the log-normal distribution at a specified point 'x' given the distribution's actual mean and standard deviation."
' Cumulative probability density function of the log-normal distribution at a
' specified point 'x' given the distribution's mean and standard deviation.
' Arguments should all be positive.
'
' Accuracy: define the "contrast" to be the standard deviation divided by the
' mean. As the contrast becomes small, a number of accuracy problems arise.
' Because of roundoff in Sqr(2 * Log(1 + c2)) that equals 5E-17/contrast^2 for
' small contrast values, and because of truncation error in the series
' approximation to the above quantity used for contrast values below 0.05,
' relative error cannot be better than 2E-14. Also, because of roundoff
' in the evaluation of Log(x*Sqr(1+contrast^2)/mean), there is a relative error
' that increases linearly with distance below the zero of the argument to the
' Log, and that increases as the reciprocal of contrast. For less than 6
' standard deviations from the mean, this error is less than about
' 4E-16 / contrast. It is therefore unwise to specify a contrast below about
' 4E-10 if you want part-per-million accuracy out to 6 sigma. You get
' part-per-thousand at 4E-13, and no accuracy at all at 4E-16. If you want
' results beyond 6 sigma, apply even tighter contrast lower limits. Closer to
' the mean, you can relax a bit.
'
' Note that the 'mean' and 'standard deviation' arguments are those of the
' actual log-normal distribution, not (as some foolish authors use) those of
' the Gaussian in the exponent.
'
' See the Maple file "logNormal.mws"

Const p_c As String = "logNormalCDF"  ' procedure name

Static calls_s As Double  ' number of times this has been called
calls_s = calls_s + 1#    ' stops adding at 9,007,199,254,740,992 calls

If x = 0# Then
  logNormalCDF = 0#
ElseIf x < 0# Then
  ' Note: when called from Excel, Err.Raise causes Excel's #VALUE! error
  Err.Raise 5&, m_c & p_c, _
    "Argument error in " & m_c & "!" & p_c & " on call " & _
    Format$(calls_s, "#,##0") & EOL & EOL & _
    "x must be >= 0 but got:" & EOL & _
    "  x = " & x & EOL & _
    "  mean = " & mean & EOL & _
    "  stdDev = " & stdDev
ElseIf mean <= 0# Then
  Err.Raise 5&, m_c & p_c, _
    "Argument error in " & m_c & "!" & p_c & " on call " & _
    Format$(calls_s, "#,##0") & EOL & EOL & _
    "mean must be > 0 but got:" & EOL & _
    "  x = " & x & EOL & _
    "  mean = " & mean & EOL & _
    "  stdDev = " & stdDev
ElseIf stdDev < 0# Then
  Err.Raise 5&, m_c & p_c, _
    "Argument error in " & m_c & "!" & p_c & " on call " & _
    Format$(calls_s, "#,##0") & EOL & EOL & _
    "stdDev must be >= 0 but got:" & EOL & _
    "  x = " & x & EOL & _
    "  mean = " & mean & EOL & _
    "  stdDev = " & stdDev
ElseIf stdDev = 0# Then
  If x < mean Then
    logNormalCDF = 0#
  Else
    logNormalCDF = 1#
  End If
Else
  Dim contrast As Double, c2 As Double, t1 As Double
  contrast = stdDev / mean
  c2 = contrast * contrast
  If contrast > 0.05 Then  ' roundoff < 2E-14 (relative); use formula
    ' the formula has roundoff that goes as 5.5E-17 / contrast^2
    t1 = Sqr(2# * Log(1# + c2))
  Else  ' low contrast; too much roundoff; use series for Sqr(2 * Log(1 + c2))
    ' the series has error at or below 2E-14 for 0 < contrast < 0.05
    t1 = (1.41421356237307 - (0.353553390228032 - _
      (0.191507343543864 - 0.128421019600657 * c2) * c2) * c2) * contrast
  End If
  logNormalCDF = 0.5 * erfcT(-Log(x * Sqr(1# + c2) / mean) / t1)
End If
End Function

'===============================================================================
Public Function logNormalPDF( _
  ByRef x As Double, _
  ByRef mean As Double, _
  ByRef stdDev As Double) _
As Double
Attribute logNormalPDF.VB_Description = "Probability Density Function of the log-normal distribution at a specified point 'x' given the distribution's actual mean and standard deviation."
' Probability density function of the log-normal distribution at a specified
' point 'x' given the distribution's mean and standard deviation. Arguments
' should all be positive (0 is allowed for 'x').
'
' Accuracy: define the "contrast" to be the standard deviation divided by the
' mean. As the contrast becomes small, a number of accuracy problems arise.
' Because of roundoff in 2*Log(1+contrast^2) that equals 1E-16/contrast^2 for
' small contrast values, and because of truncation error in the series
' approximation to the above quantity used for contrast values below 0.0484,
' relative error cannot be better than 4.5E-14. Also, because of roundoff
' in the evaluation of Log(x*Sqr(1+contrast^2)/mean), there is a relative error
' that rises linearly with distance from the zero of the argument to the Log,
' and that increases as the reciprocal of contrast. For less than 6 standard
' deviations from the mean, this error is less than about 8E-16 / contrast.
' It is therefore unwise to specify a contrast below about 1E-9 if you
' want part-per-million accuracy out to 6 sigma. You get part-per-thousand
' at 1E-12, and no accuracy at all at 1E-17. If you want results beyond 6 sigma,
' apply even tighter contrast lower limits. Closer to the mean, you can relax
' a bit.
'
' Note that the 'mean' and 'standard deviation' arguments are those of the
' actual log-normal distribution, not (as some foolish authors use) those of
' the Gaussian in the exponent.
'
' See the Maple file "logNormal.mws"

Const p_c As String = "logNormalPDF"  ' procedure name

' numeric constants must be written as sums to maintain full accuracy in files
Const Pi_c As Double = 3.1415926 + 5.35897932E-08

Static calls_s As Double  ' number of times this has been called
calls_s = calls_s + 1#    ' stops adding at 9,007,199,254,740,992 calls

If x = 0# Then
  logNormalPDF = 0#
ElseIf x < 0# Then
  ' Note: when called from Excel, Err.Raise causes Excel's #VALUE! error
  Err.Raise 5&, m_c & p_c, _
    "Argument error in " & m_c & "!" & p_c & " on call " & _
    Format$(calls_s, "#,##0") & EOL & EOL & _
    "x must be >= 0 but got:" & EOL & _
    "  x = " & x & EOL & _
    "  mean = " & mean & EOL & _
    "  stdDev = " & stdDev
ElseIf mean <= 0# Then
  Err.Raise 5&, m_c & p_c, _
    "Argument error in " & m_c & "!" & p_c & " on call " & _
    Format$(calls_s, "#,##0") & EOL & EOL & _
    "mean must be > 0 but got:" & EOL & _
    "  x = " & x & EOL & _
    "  mean = " & mean & EOL & _
    "  stdDev = " & stdDev
ElseIf stdDev <= 0# Then
  Err.Raise 5&, m_c & p_c, _
    "Argument error in " & m_c & "!" & p_c & " on call " & _
    Format$(calls_s, "#,##0") & EOL & EOL & _
    "stdDev must be > 0 but got:" & EOL & _
    "  x = " & x & EOL & _
    "  mean = " & mean & EOL & _
    "  stdDev = " & stdDev
Else
  Dim contrast As Double, c2 As Double, t1 As Double, t2 As Double
  contrast = stdDev / mean
  c2 = contrast * contrast
  If contrast > 0.0484 Then  ' roundoff < 4.5E-14 (relative); use formula
    ' the formula has roundoff that goes as 1.1E-16 / contrast^2
    t1 = 2# * Log(1# + c2)
  Else  ' low contrast; too much roundoff; use series for 2 * Log(1 + c^2)
    ' the series has error at or below 4.5E-14 for 0 < contrast < 0.0484
    t1 = (1.99999999999991 - (0.999999998717904 - _
      (0.666663929477122 - 0.498129912144896 * c2) * c2) * c2) * c2
  End If
  t2 = Log(x * Sqr(1# + c2) / mean)
  logNormalPDF = Exp(-t2 * t2 / t1) / (x * Sqr(Pi_c * t1))
End If
End Function

'===============================================================================
Public Function logNormalVersion() As String
Attribute logNormalVersion.VB_Description = "The date of the latest revision to this module as a string in the format 'YYYY-MM-DD' such as 2004-06-18. It's a function so Excel etc. can use it."
' The date of the latest revision to this module as a string in the format
' 'YYYY-MM-DD' such as 2004-06-18. It's a function so Excel etc. can use it.
logNormalVersion = Version_c
End Function

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
#If UnitTest Then

Public Sub Test_LogNormal()
Attribute Test_LogNormal.VB_Description = "Unit-test routine for this module. Sends results to a file and to the immediate window (if in IDE)."
' Main unit test routine for this module.

' To run the test from VB6, enter this routine's name (above) in the Immediate
' window (if the Immediate window is not open, use View.. or Ctrl-G to open it).
' To run this routine from VBA, put the cursor somewhere in it and hit F5.
' To run this routine from somewhere in a code, call it.

' The output will be in the file 'Test_LogNormal.txt' on disk, and in the
' immediate window if in the VB[6,A] editor.

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
    MsgBox "Warning! Workbook #1 has no disk location!" & vbNewLine & _
           "Save workbook to disk before proceeding because" & vbNewLine & _
           "routine needs a known location to write to." & vbNewLine & _
           "No unit test output will be written to file.", _
           vbOKOnly Or vbCritical, _
           m_c & " Unit Test"
    Exit Sub
  End If
#Else
  ' note: this is the project folder if in VB6 IDE; EXE folder if stand-alone
  path = App.path
#End If
If Right$(path, 1) <> "\" Then path = path & "\"  ' only C:\ etc. have "\"
ofs = path & "Test_" & m_c & ".txt"

ofi_m = FreeFile
On Error Resume Next
Open ofs For Output As #ofi_m  ' output file
If Err.Number <> 0 Then
  ofi_m = 0  ' file did not open - don't use it
  MsgBox "ERROR - unable to open output file:" & EOL & EOL & _
         """" & ofs & """" & EOL & EOL & _
         "No unit test output will be written to file.", _
         vbOKOnly Or vbExclamation, _
         m_c & " Unit Test"
End If
On Error GoTo 0

teeOut "######## Test of " & m_c & " routines at " & Now()
teeOut "Code version: " & Version_c
teeOut

nWarn = 0&

teeOut "=== logNormalCDF ==="
' check logNormalCDF against accurate numeric values calculated by Maple
worst = 0#
compareRel "logNormalCDF(0.0,1.0,1.0)", logNormalCDF(0#, 1#, 1#), 0#, worst
compareRel "logNormalCDF(1.0,1.0,1.0)", logNormalCDF(1#, 1#, 1#), _
  0.661396451413337, worst
compareRel "logNormalCDF(2.1,1.9,0.6)", logNormalCDF(2.1, 1.9, 0.6), _
  0.683949452144803, worst
compareRel "logNormalCDF(1.0,1.0,0.02)", logNormalCDF(1#, 1#, 0.02), _
  0.503988957478706, worst
compareRel "logNormalCDF(0.99,1.0,0.02)", logNormalCDF(0.99, 1#, 0.02), _
  0.311158918130987, worst
compareRel "logNormalCDF(1.0001,1.0,0.0001)", _
  logNormalCDF(1.0001, 1#, 0.0001), 0.841344747479918, worst
compareRel "logNormalCDF(0.999999995,1.0,0.00000001)", _
  logNormalCDF(0.999999995, 1#, 0.00000001), 0.308537540046232, worst
compareRel "logNormalCDF(0.999999999999999,1.0,0.0)", _
  logNormalCDF(0.999999999999999, 1#, 0#), 0#, worst
compareRel "logNormalCDF(1.0,1.0,0.0)", _
  logNormalCDF(1#, 1#, 0#), 1#, worst

Dim f As Double
teeOut "--- abnormal input ---"
On Error Resume Next
Err.Clear: f = logNormalCDF(-1#, 1#, 1#)
  If Err.Number <> 0& Then f = Err.Number
  compareRel "logNormalCDF(-1.0,1.0,1.0)", f, 5#, worst
Err.Clear: f = logNormalCDF(1#, 0#, 1#)
  If Err.Number <> 0& Then f = Err.Number
  compareRel "logNormalCDF(1.0,0.0,1.0)", f, 5#, worst
Err.Clear: f = logNormalCDF(1#, 1#, -1#)
  If Err.Number <> 0& Then f = Err.Number
  compareRel "logNormalCDF(1.0,1.0,-1.0)", f, 5#, worst
On Error GoTo 0

teeOut "Largest logNormalCDF relative error: " & Format(worst, "0.000000E-0")
limit = 0.0000000025
If Abs(worst) > limit Then
  teeOut "WARNING! That's too large - should be less than " & _
         Format(limit, "0.0000E-0")
  nWarn = nWarn + 1&
End If
teeOut

teeOut "=== logNormalPDF ==="
' check logNormalPDF against accurate numeric values calculated by Maple
worst = 0#
compareRel "logNormalPDF(0.0,1.0,1.0)", logNormalPDF(0#, 1#, 1#), 0#, worst
compareRel "logNormalPDF(1.0,1.0,1.0)", logNormalPDF(1#, 1#, 1#), _
  0.43940863365672, worst
compareRel "logNormalPDF(2.1,1.9,0.6)", logNormalPDF(2.1, 1.9, 0.6), _
  0.549440555382819, worst
compareRel "logNormalPDF(1.0,1.0,0.02)", logNormalPDF(1#, 1#, 0.02), _
  19.9481112677461, worst
compareRel "logNormalPDF(1.12,1.0,0.02)", logNormalPDF(1.12, 1#, 0.02), _
  1.78818782733664E-06, worst
compareRel "logNormalPDF(1.0006,1.0,0.0001)", _
  logNormalPDF(1.0006, 1#, 0.0001), 6.13629750571736E-05, worst
compareRel "logNormalPDF(1.00000006,1.0,0.00000001)", _
  logNormalPDF(1.00000006, 1#, 0.00000001), 0.607588886494993, worst

teeOut "--- abnormal input ---"
On Error Resume Next
Err.Clear: f = logNormalPDF(-1#, 1#, 1#)
  If Err.Number <> 0& Then f = Err.Number
  compareRel "logNormalPDF(-1.0,1.0,1.0)", f, 5#, worst
Err.Clear: f = logNormalPDF(1#, 0#, 1#)
  If Err.Number <> 0& Then f = Err.Number
  compareRel "logNormalPDF(1.0,0.0,1.0)", f, 5#, worst
Err.Clear: f = logNormalPDF(1#, 1#, 0#)
  If Err.Number <> 0& Then f = Err.Number
  compareRel "logNormalPDF(1.0,1.0,0.0)", f, 5#, worst
On Error GoTo 0

teeOut "Largest logNormalPDF relative error: " & Format(worst, "0.000000E-0")
limit = 0.00000002
If Abs(worst) > limit Then
  teeOut "WARNING! That's too large - should be less than " & _
         Format(limit, "0.0000E-0")
  nWarn = nWarn + 1&
End If
teeOut

If nWarn = 0& Then
  teeOut "Success - all errors were within limits."
Else
  teeOut "FAILURE! - warning count: " & nWarn
End If
teeOut "--- Test complete ---"
Close #ofi_m
End Sub

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Private Sub compareAbs( _
  ByVal str As String, _
  ByVal approx As Double, _
  ByVal exact As Double, _
  ByRef worst As Double)
' Unit-test support routine. Makes an absolute comparison of 'approx' to
' 'exact', updates 'worst', and sends results to 'teeOut' prefixed by 'str'.
' John Trenholme - 2002-07-09

Dim absErr As Double

absErr = approx - exact
If Abs(worst) < Abs(absErr) Then worst = absErr
teeOut str
teeOut "  approx " & Format(approx, "0.000000000000000E-0") & _
       "  exact " & Format(exact, "0.000000000000000E-0") & _
       "  absErr " & Format(absErr, "0.000E-0")
End Sub

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Private Sub compareRel( _
  ByVal str As String, _
  ByVal approx As Double, _
  ByVal exact As Double, _
  ByRef worst As Double)
' Unit-test support routine. Makes a relative comparison of 'approx' to 'exact',
' updates 'worst', and sends results to 'teeOut' prefixed by 'str'.
' John Trenholme - 2006-07-12

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
teeOut str
teeOut "  approx " & Format(approx, "0.000000000000000E-0") & _
       "  exact " & Format(exact, "0.000000000000000E-0") & _
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

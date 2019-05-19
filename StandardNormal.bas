Attribute VB_Name = "StandardNormal"
Attribute VB_Description = "Functions related to the standard normal (Gaussian) distribution. Devised & coded by John Trenholme."
'
'###############################################################################
'#
'# Visual Basic 6 & VBA code module "StandardNormal.bas"
'#
'# Standard Normal functions.
'#
'# Exports the routines:
'#   Function snCDF
'#   Function snInvCDF
'#   Function snPDF
'#   Function StandardNormalVersion
'#
'# Devised and coded by John Trenholme - initial version 5 Jul 2002
'#
'###############################################################################

Option Explicit

Private Const Version_c As String = "2009-06-18"
Private Const M_c As String = "StandardNormal"

Private Const EOL As String = vbNewLine  ' short form; works on both PC and Mac

' set this to "True" to use unit test routines
' set it to "False" to avoid compiling unit test routines into code
#Const UnitTest = True
' #Const UnitTest = False

' set this "True" if this is VBA under Excel (or others)
' set this False if this is VB6
' #Const VBA = True
#Const VBA = False

#If UnitTest Then
  Private ofi_m As Integer  ' output file index used by unit test routine
#End If

'===============================================================================
Public Function snCDF(ByVal z As Double) As Double
Attribute snCDF.VB_Description = "Cumulative Distribution Function of a standard normal variate to 1E-14 worst relative error within 6 sigma of the mean. Gives 0 for z < -37.5 & 1 for z > 8.2923."
' Returns the cumulative distribution function of a standard normal variate
' with a relative error of 1E-14 or better within 6 sigma of the mean. This
' rises to 3E-13 far out on the negative tail, because of roundoff in Exp(x).
' If z < -37.5, 0.0 is returned, due to approaching double-precision underflow
' (note that snCDF( -37.5) = 4.6E-308).

' For large positive z, the difference from unity is small and suffers from
' roundoff error. If you want this difference to high accuracy, note that
' 1 - snCDF(z) = snCDF(-z). This quantity (the exceedance) is small and accurate
' just where snCDF(z) gets into trouble. For z > 8.2923, the returned value is
' exactly 1.0 because of roundoff in the subtraction.

' The CDF of a normal variate with non-zero mean and non-unit standard deviation
' is given by snCDF((x - mean) / stdDev).

' Note: CDF of the max of N independent samples is this CDF to the Nth power.

' Version of 2008-11-08 - see Maple file "CumulativeNormalApproxI.mws"

Dim a As Double
Dim r As Double, s As Double
Dim x2 As Double
Dim z As Double
a = Abs(x)
x2 = x * x
If a < 0.786 Then
  r = 0.5 + x * (286274.853432297 + (13766.2029201158 + _
    2102.12025748516 * x2) * x2) / (717584.641929165 + _
    (154104.193799632 + (13013.6502555081 + (452.005967245251 + _
    x2) * x2) * x2) * x2)
ElseIf x > 8.3 Then  ' limited by granularity in floating point values near 1
  r = 1#
ElseIf x < -37.5 Then  ' limited by smallest possible floating point value
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
    z = 1# / x2
    ' split evaluation into parts to avoid "expression too complex"
    s = 3.16965563545039E-04 + (0.015227793862574 + (0.260705053546659 + _
      (2.00621176176922 + (7.22108903104585 + (11.5148288549791 + _
      (6.89233938349171 + z) * z) * z) * z) * z) * z) * z
    r = 0.712684696150189 + (2.31460584415985 + (3.07234263767005 + _
      (1.28623813814016 + 7.21836636375586E-02 * z) * z) * z) * z
    r = ((1.26450964729363E-04 + (5.94855984431045E-03 + _
      (9.83106106510839E-02 + r * z) * z) * z) / s + 4E-16) / a
  End If
  r = r * Exp(-0.5 * x2)
  If x > 0# Then r = 1# - r
End If
snCDF = r
End Function

'===============================================================================
Public Function snInvCDF(ByVal probability As Double) As Double
Attribute snInvCDF.VB_Description = "Inverse of Cumulative Distribution Function of a standard normal variate to 6.3E-11 worst absolute error. Gives -38.4674 for prob <= 0 & 38.4674 for prob >= 1."
' Returns the standard normal argument corresponding to the supplied CDF
' probability. The maximum absolute error is around 6.3E-11, except for
' very small probabilities (less than 5E-324) where underflow occurs, or for
' probabilities above about 0.9999997, where floating-point granularity
' limits accuracy. Return values will always be less than 38.4674 in absolute
' value. If the supplied probability is <= 0 or >= 1, it is silently put back
' "in bounds" with no error, and -38.4674 or 38.4674 is returned.

' Note: inverse of max of N independent samples is snInvCDF(p ^ (1# / N))

' Version of 2008-02-21 - see Maple file "InverseCumulativeNormal4.mws"

Dim r As Double
Dim u As Double
Dim ua As Double
Dim u2 As Double
Dim w As Double

u = probability - 0.5
ua = Abs(u)
u2 = u * u

If ua < 0.37377 Then  ' P > 0.12623 and P < 0.87377
  r = (1.10320796594653 - (7.77941365082975 - (16.1360412312915 - _
    8.94247760684027 * u2) * u2) * u2) * u / _
    (0.440116302105953 - (3.56442583646134 - (9.15646709284907 - _
    (7.69878138754029 - u2) * u2) * u2) * u2)
ElseIf ua < 0.44286 Then  ' P > 0.05714 and P < 0.94286
  r = (0.317718558863025 - (2.70051978050927 - (7.20258279324852 - _
    5.82660777818178 * u2) * u2) * u2) * u / _
    (0.126757926972973 - (1.21032688875879 - (3.85234822216469 - _
    (4.38840255884193 - u2) * u2) * u2) * u2)
ElseIf (probability < 4.94065645841247E-324) Or (probability >= 1#) Then
  r = Sgn(u) * 38.4674
Else
  If probability < 0.5 Then
    w = Sqr(-Log(probability))
  Else
    w = Sqr(-Log(1# - probability))  ' roundoff noise above 0.9999997
  End If
  If w < 3.769 Then  ' P or 1-P > 6.77158141318452E-07
    w = (3.40265621744676 + (9.03080228605413 - (6.88823432035713 + _
      (9.47396446577765 + 1.41485388628381 * w) * w) * w) * w) / _
      (1.10738880205572 + (7.00041795498572 + (6.72600088945649 + w) * w) * w)
   ElseIf w < 8.371 Then  ' P or 1-P > 3.69321326562547E-31
     w = (27.5896468790036 + (11.8481686174627 - (37.7133528390963 + _
       (18.6301980539071 + 1.41437483654701 * w) * w) * w) * w) / _
       (10.7729777720728 + (29.1330213184579 + (13.1871785457772 + w) * w) * w)
   Else  ' P <= 3.69321326562547E-31
     w = (859575101.771399 - (167079541.087701 + (887823598.683122 + _
       (206626688.300811 + 7785160.41001698 * w) * w) * w) * w) / _
       (382471838.745491 + (643197787.097259 + (146148247.666043 + _
       (5504690.97847543 + w) * w) * w) * w)
   End If
   r = Sgn(-u) * w
End If
snInvCDF = r
End Function

'===============================================================================
Public Function snPDF(ByVal z As Double) As Double
Attribute snPDF.VB_Description = "Probability Distribution Function of a standard normal variate to 2E-15 worst relative error. Gives 0 for |z| > 38.56804."
' Returns the probability density function of a standard normal variate. If the
' absolute magnitude of z is greater than about 38.56804 (where the result is
' 4.94E-324), the result will be exactly zero because of the way VB underflows.

' Note: the PDF of a normal variate with non-zero mean and non-unit standard
' deviation is given by snPDF((x - mean) / stdDev) / stdDev.

' Version of 5 Jul 2002

' written as sum to avoid truncation error when writing to & reading from file
Const NormPDF_c As Double = 0.39894228 + 4.014327E-10  ' 1 / Sqr(2 * pi)

snPDF = Exp(-0.5 * z * z) * NormPDF_c
End Function

'===============================================================================
Public Function standardNormalVersion() As String
Attribute standardNormalVersion.VB_Description = "The date of the latest revision to this module as a string in the format 'YYYY-MM-DD' such as 2004-06-18. It's a function so Excel etc. can use it."
' The date of the latest revision to this module as a string in the format
' 'YYYY-MM-DD' such as 2004-06-18. It's a function so Excel etc. can use it.
standardNormalVersion = Version_c
End Function

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'&
'& Unit test
'&
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

#If UnitTest Then

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Public Sub Test_StandardNormal()
Attribute Test_StandardNormal.VB_Description = "Unit test routine. Test results go to file 'Test_StandardNormal.txt' & to Immediate window (if in VB[A] editor)."
' Main unit test routine for this module.

' To run the test from VB, enter this routine's name (above) in the Immediate
' window (if the Immediate window is not open, use View.. or Ctrl-G to open it).
' To run this routine from VBA, put the cursor somewhere in it and hit F5.
' To run this routine from somewhere in a code, call it.

' The output will be in a file on disk (see path\name below) and in the
' immediate window (use Ctrl-G to open if not visible) if in the VB[A] editor.

' Version of 2008-02-21 - John Trenholme

Dim limit As Double
Dim nWarn As Long
Dim worst As Double

' get path to current directory, and prepend to file name
Dim ofs As String
Dim path As String
' note: in Excel, save workbook at least once so path exists
#If VBA Then
  path = ThisWorkbook.path  ' VBA in Excel (or others)
#Else
  path = App.path  ' VB6
#End If
If Right$(path, 1) <> "\" Then path = path & "\"  ' only C:\ etc. have "\"
ofs = path & "Test_" & M_c & ".txt"

ofi_m = FreeFile
On Error Resume Next
Open ofs For Output As #ofi_m  ' output file
If Err.Number <> 0 Then
  ofi_m = 0  ' file did not open - don't use it
  MsgBox "ERROR - unable to open output file:" & EOL & EOL & _
    """" & ofs & """" & EOL & EOL & _
    "No unit test output will be written to file.", _
    vbOKOnly Or vbExclamation, M_c & " Unit Test"
End If
On Error GoTo 0

teeOut "########## Test of " & M_c & " routines at " & Now()

nWarn = 0

' snCDF - argument values are specific to the approximation used
worst = 0#
compareRel "snCDF(-37.5000000000001)", snCDF(-37.5000000000001), 0#, worst
compareRel "snCDF(-37.5)", snCDF(-37.5), 4.60535300958196E-308, worst
compareRel "snCDF(-19.5)", snCDF(-19.5), 5.48911547566041E-85, worst
compareRel "snCDF(-6.2)", snCDF(-6.2), 2.82315803704328E-10, worst
compareRel "snCDF(-5.53)", snCDF(-5.53), 1.6011539388091E-08, worst
compareRel "snCDF(-1.04)", snCDF(-1.04), 0.149169950330981, worst
compareRel "snCDF(0.0)", snCDF(0#), 0.5, worst
compareRel "snCDF(1.04)", snCDF(1.04), 0.850830049669019, worst
teeOut "Largest snCDF relative error: " & Format(worst, "0.000000E-0")
limit = 0.00000004
If Abs(worst) > limit Then
  teeOut "WARNING! That's too large - should be less than " & _
         Format(limit, "0.0000E-0")
  nWarn = nWarn + 1
End If
teeOut

' snInvCDF - argument values are not specific to the approximation used
worst = 0#
compareAbs "snInvCDF(-1.0)", snInvCDF(-1#), -38.4674056172733, worst
compareAbs "snInvCDF(3E-308)", snInvCDF(3E-308), -37.511419674256, worst
compareAbs "snInvCDF(1E-200)", snInvCDF(1E-200), -30.20559417958, worst
compareAbs "snInvCDF(1E-100)", snInvCDF(1E-100), -21.273453560966, worst
compareAbs "snInvCDF(1E-30)", snInvCDF(1E-30), -11.464024688444, worst
compareAbs "snInvCDF(1E-20)", snInvCDF(1E-20), -9.2623400897985, worst
compareAbs "snInvCDF(1E-10)", snInvCDF(0.0000000001), -6.3613409024041, worst
compareAbs "snInvCDF(0.215)", snInvCDF(0.215), -0.7891916527, worst
compareAbs "snInvCDF(0.388)", snInvCDF(0.388), -0.2845355427, worst
compareAbs "snInvCDF(0.5)", snInvCDF(0.5), 0#, worst
compareAbs "snInvCDF(0.612)", snInvCDF(0.612), 0.2845355427, worst
compareAbs "snInvCDF(0.785)", snInvCDF(0.785), 0.7891916527, worst
compareAbs "snInvCDF(2.0)", snInvCDF(2#), 38.4674056172733, worst
teeOut "Largest snInvCDF absolute error: " & Format(worst, "0.000000E-0")
limit = 0.000000000125
If Abs(worst) > limit Then
  teeOut "WARNING! That's too large - should be less than " & _
         Format(limit, "0.0000E-0")
  nWarn = nWarn + 1
End If
teeOut "Check in the region of large granularity error:"
compareAbs "snInvCDF(1.0-1E-15)", snInvCDF(1# - 0.000000000000001), _
  7.941345326, worst
teeOut

' snInvPDF - just testing that correct expression is used
worst = 0#
compareRel "snPDF(-1.0)", snPDF(-1#), 0.241970724519143, worst
compareRel "snPDF(0.0)", snPDF(0#), 0.398942280401433, worst
compareRel "snPDF(1.0)", snPDF(1#), 0.241970724519143, worst
compareRel "snPDF(2.0)", snPDF(2#), 0.053990966513188, worst
compareRel "snPDF(3.0)", snPDF(3#), 0.004431848411938, worst
compareRel "snPDF(9.0)", snPDF(9#), 1.02797735716689E-18, worst
compareRel "snPDF(37.6)", snPDF(37.6), 4.04414480934867E-308, worst
teeOut "Largest snPDF relative error: " & Format(worst, "0.000000E-0")
limit = 0.00000000000011
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
teeOut str
teeOut "  approx " & Format(approx, "0.00000000000E-0") & _
       "  exact " & Format(exact, "0.00000000000E-0") & _
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
teeOut str
teeOut "  approx " & Format(approx, "0.00000000000E-0") & _
       "  exact " & Format(exact, "0.00000000000E-0") & _
       "  relErr " & Format(relErr, "0.000E-0")
End Sub

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Private Sub teeOut(Optional ByRef str As String = "")
' unit test support routine - John Trenholme - 3 Sep 2002

Debug.Print str  ' works only if in VB[A] editor environment
If ofi_m <> 0 Then Print #ofi_m, str
End Sub

#End If  ' UnitTest

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


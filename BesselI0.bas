Attribute VB_Name = "BesselI0Mod"
Attribute VB_Description = "Medium-accuracy approximation of the modified Bessel function of order 0. Devised and coded by John B. Trenholme."
'
'###############################################################################
'#
'# Visual Basic 6 & VBA code module "BesselI0.bas"
'#
'# The modified Bessel function of order zero.
'#
'# Exports the routines:
'#   Function besselI0
'#   Function besselI0Version
'#   Sub Test_BesselI0 (if UnitTest is True)
'#
'# Devised & coded by John Trenholme
'# Begun 2001-12-03
'#
'###############################################################################

Option Base 0
Option Compare Binary
Option Explicit
Option Private Module  ' Don't allow visibility outside this Project

Private Const Version_c As String = "2006-07-25"
Private Const m_c As String = "BesselI0"  ' module name

#Const UnitTest = True  ' set True to enable unit test code

Private Const EOL As String = vbNewLine  ' short form; works on both PC and Mac

'===============================================================================
Public Function besselI0(x As Double) As Double
Attribute besselI0.VB_Description = "Modified Bessel function of order 0, accurate to 3.9E-8 worst relative error for |x| <= 709.782712893384 (note I0(709.782712893384) = 2.6923992295663E+306)."
' Returns the modified Bessel function of order zero.
' |Relative error| <= 3.9E-8 for any value of x.
' Must have |x| <= 709.782712893384 to avoid overflow.
' Note that besselI0(709.782712893384) = 2.6923992295663E+306
' There is a slight discontinuity of slope near |x| = 4.07261970484112.

' See the Maple file "fitI0.mws"

Const p_c = "besselI0"  ' procedure name

Static calls_s As Double  ' number of times this routine has been called
calls_s = calls_s + 1#    ' stops adding at 9,007,199,254,740,992 calls

Dim t1 As Double, t2 As Double, u As Double, v As Double, w As Double

u = Abs(x)  ' function is even around x = 0, so only do positive part
If u > 709.782712893384 Then
  ' Note: when called from Excel, Err.Raise causes Excel's #VALUE! error
  Err.Raise 6&, m_c & "!" & p_c, _
    "Overflow error in " & m_c & "!" & p_c & " on call " & _
    Format$(calls_s, "#,##0") & EOL & EOL & _
    "|x| must be <= 709.782712893384 but:" & EOL & _
    "  x = " & x & EOL & _
    "Note that besselI0(709.782712893384) = 2.6923992295663E+306"
End If

' the join point is selected to match value and (approximately) slope
If u <= 4.07261970484112 Then
  v = x * x  ' use Pade form in x^2
  w = (1# + v * (0.2251516395 + v * (0.009693409864 + _
    v * 0.0001146169854))) / _
    (1# - v * (0.02484854838 - v * (0.0002808018655 - _
    v * 0.00000145965874)))
Else
  v = 1# / u  ' use asymptotic series as Pade form in 1/|x|
  ' VB expression-too-complex errors if not written out in parts this way
  t1 = v * (4.796896159 + v * (2.86881326 - v * 4.011967236))
  t1 = 0.3989422804 - v * (2.402141371 - t1)
  t2 = v * (12.72245399 + v * (5.948117657 - v * 11.23250366))
  t2 = 1# - v * (6.146281039 - t2)
  w = Exp(u) * Sqr(v) * t1 / t2
End If
besselI0 = w
End Function

'===============================================================================
Public Function besselI0Version() As String
Attribute besselI0Version.VB_Description = "The date of the latest revision to this module as a string in the format 'YYYY-MM-DD' such as 2004-06-18. It's a function so Excel etc. can use it."
' The date of the latest revision to this module as a string in the format
' 'YYYY-MM-DD' such as 2004-06-18. It's a function so Excel etc. can use it.
besselI0Version = Version_c
End Function

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
#If UnitTest Then

Public Sub Test_BesselI0()
Attribute Test_BesselI0.VB_Description = "Unit test routine. Test results go to file in the IDE/EXE directory & to Immediate window (if in IDE)."
' Main unit-test routine for this module.

' To run the test from VB6, enter this routine's name (above) in the Immediate
' window (if the Immediate window is not open, use View.. or Ctrl-G to open it).
' To run this routine from VBA, put the cursor somewhere in it and hit F5.
' To run this routine from somewhere in a code, call it.

' The output will be in the file 'Test_Bessel_I0.txt' on disk, and in the
' immediate window if in the VB[6,A] editor.

Dim nWarn As Long
nWarn = 0&

utFileOpen "Test_" & m_c & ".txt"

utTeeOut "########## Test of " & m_c & " routines at " & Now()
utTeeOut "Code version: " & Version_c
utTeeOut

' do binary search to find maximum possible argument before overflow
Dim d As Double, f As Double, x As Double
x = 705#
d = 10#
Do While d > 3E-16 * x
  d = 0.5 * d
  On Error Resume Next
  f = besselI0(x + d)
  If Err.Number = 0 Then x = x + d  ' small enough; nudge up
  On Error GoTo 0
Loop
utTeeOut "Largest possible argument before overflow = " & x
x = 709.782712893384
utTeeOut "Value of function at x = " & x & " is " & besselI0(x)
utTeeOut

' force an error
On Error Resume Next
x = 709.782712893385
Err.Clear
f = besselI0(x)
utErrorCheck "besselI0(" & x & ")", 6&, nWarn  ' did it overflow?
On Error GoTo 0
Dim limit As Double
Dim worst As Double
worst = 0#

' these argument values are specific to the approximation used
' they alternate between zeros & maxima of the error curve
utCompareRel "I0(0.0)", besselI0(0#), 1#, worst
utCompareRel "I0(0.6929382)", besselI0(0.6929382), 1.123691699, worst
utCompareRel "I0(1.134756)", besselI0(1.134756), 1.348771134, worst
utCompareRel "I0(-1.558315)", besselI0(-1.558315), 1.705681644, worst
utCompareRel "I0(1.970503)", besselI0(1.970503), 2.233306705, worst
utCompareRel "I0(2.369916)", besselI0(2.369916), 2.981058221, worst
utCompareRel "I0(2.749343)", besselI0(2.749343), 3.993840617, worst
utCompareRel "I0(-3.097791)", besselI0(-3.097791), 5.284944407, worst
utCompareRel "I0(3.410674)", besselI0(3.410674), 6.845628286, worst
utCompareRel "I0(3.675541)", besselI0(3.675541), 8.558744454, worst
utCompareRel "I0(3.889478)", besselI0(3.889478), 10.27562361, worst
utCompareRel "I0(-4.048445)", besselI0(-4.048445), 11.78527478, worst
utCompareRel "I0(4.0726197)", besselI0(4.07262), 12.03455148, worst  ' joint
utCompareRel "I0(4.190386)", besselI0(4.190386), 13.33040931, worst
utCompareRel "I0(4.390003)", besselI0(4.390003), 15.870654, worst
utCompareRel "I0(4.654062)", besselI0(4.654062), 20.02717304, worst
utCompareRel "I0(-5.045880)", besselI0(-5.04588), 28.38027712, worst
utCompareRel "I0(5.552404)", besselI0(5.552404), 44.76661981, worst
utCompareRel "I0(6.162154)", besselI0(6.162154), 77.97032084, worst
utCompareRel "I0(6.987602)", besselI0(6.987602), 166.6705361, worst
utCompareRel "I0(-8.059230)", besselI0(-8.05923), 451.9235213, worst
utCompareRel "I0(9.596594)", besselI0(9.596594), 1921.275799, worst
utCompareRel "I0(12.09783)", besselI0(12.09783), 20809.82354, worst
utCompareRel "I0(16.76299)", besselI0(16.76299), 1871328.843, worst
utCompareRel "I0(-24.28060)", besselI0(-24.2806), 2854269135#, worst
utCompareRel "I0(42.51715)", besselI0(42.51715), 1.790127002E+17, worst
utCompareRel "I0(107.1033)", besselI0(107.1033), 1.261508197E+45, worst
utCompareRel "I0(-709.7)", besselI0(-709.7), 2.478808727E+306, worst

utTeeOut
utTeeOut "Largest BesselI0 relative error: " & Format(worst, "0.000000E-0")
limit = 0.000000039
If Abs(worst) > limit Then
  utTeeOut "WARNING! That's too large - should be less than " & _
           Format(limit, "0.0000E-0")
  nWarn = nWarn + 1&
End If
utTeeOut

If nWarn = 0& Then
  utTeeOut "Success - all errors were within limits."
Else
  utTeeOut "FAILURE! - warning count: " & nWarn
End If

utTeeOut
utTeeOut "--- Test complete ---"

utFileClose
End Sub

#End If

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

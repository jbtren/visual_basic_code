Attribute VB_Name = "Bessel_I0"
Attribute VB_Description = "The modified Bessel function of order 0."
'
'###############################################################################
'#
'#  Visual Basic source file "Bessel_I0.bas"
'#
'#  Exports the routine:
'#    Function BesselI0
'#
'###############################################################################

Option Explicit

Private m_ofi As Integer  ' output file index used by unit test routine
'

'*******************************************************************************
Public Function BesselI0(x As Double) As Double
Attribute BesselI0.VB_Description = "Modified Bessel function of order 0 to 3.9E-8 worst relative error. Forces |x| <= 709.7 (note I0(709.7) = 2.48E306)."
' Returns the modified Bessel function of order zero.
' Relative error is 3.9E-8 or better for any value of x.
' The absolute value of x should be less than 709.7 to avoid overflow.
' There is a slight discontinuity of slope near |x| = 4.07261970484112.

' Version of 26 Jun 2002 - John Trenholme

Dim A As Double, B As Double, u As Double, v As Double, w As Double

u = Abs(x)
If u > 709.7 Then u = 709.7  ' avoids overflow in exp()

' the join point is selected to match value and (approximately) slope
If u <= 4.07261970484112 Then
  ' use Pade form in x^2
  v = x * x
  w = (1# + v * (0.2251516395 + v * (0.009693409864 + _
    v * 0.0001146169854))) / _
    (1# - v * (0.02484854838 - v * (0.0002808018655 - _
    v * 0.00000145965874)))
Else
  ' use asymptotic series as Pade form in 1/|x|
  v = 1# / u
  ' VB expression-too-complex errors if not written out in parts this way
  A = v * (4.796896159 + v * (2.86881326 - v * 4.011967236))
  A = 0.3989422804 - v * (2.402141371 - A)
  B = v * (12.72245399 + v * (5.948117657 - v * 11.23250366))
  B = 1# - v * (6.146281039 - B)
  w = Exp(u) * Sqr(v) * A / B
End If
BesselI0 = w
End Function

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'&
'& Unit test
'&
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

' set this to "True" to use unit test routines
' set it to "False" to avoid compiling unit test routines into code
#If True Then
'#If False Then

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Public Sub Test_Bessel_I0()
Attribute Test_Bessel_I0.VB_Description = "Unit test routine. Test results go to file in the IDE/EXE directory & to Immediate window (if in IDE)."
' Main unit test routine for this module.

' To run the test from VB, enter this routine's name (above) in the Immediate
' window (if the Immediate window is not open, use View.. or Ctrl-G to open it).
' To run this routine from VBA, put the cursor somewhere in it and hit F5.
' To run this routine from somewhere in a code, call it.

' The output will be in a file on disk in the IDE/EXE directory and in the
' immediate window (use Ctrl-G to open if not visible) if in the VB[A] editor.

' Version of 3 Sep 2002 - John Trenholme

Const ModuleName As String = "Bessel_I0"

Const limit As Double = 0.000000039
Dim worst As Double

' get path to current directory, and prepend to file name
Dim ofs As String
Dim path As String
' note: in Excel, save workbook at least once so path exists
'path = Excel.Workbooks(1).path  ' uncomment if this is VBA in Excel
path = App.path  ' uncomment if this is VB6 (either IDE or standalone)
If Right$(path, 1) <> "\" Then path = path & "\"  ' only C:\ etc. have "\"
ofs = path & "Test_" & ModuleName & ".txt"

m_ofi = FreeFile
On Error Resume Next
Open ofs For Output As #m_ofi  ' output file
If Err.Number <> 0 Then
  m_ofi = 0  ' file did not open - don't use it
  MsgBox "ERROR - unable to open output file:" & vbNewLine & vbNewLine & _
    """" & ofs & """" & vbNewLine & vbNewLine & _
    "No unit test output will be written to file.", _
    vbOKOnly Or vbExclamation, ModuleName & " Unit Test"
End If
On Error GoTo 0

teeOut "########## Test of " & ModuleName & " routines at " & Now()

' BesselI0 - argument values are specific to the approximation used
worst = 0#
compareRel "I0(0.0)", BesselI0(0#), 1#, worst
compareRel "I0(0.6929382)", BesselI0(0.6929382), 1.123691699, worst
compareRel "I0(1.134756)", BesselI0(1.134756), 1.348771134, worst
compareRel "I0(-1.558315)", BesselI0(-1.558315), 1.705681644, worst
compareRel "I0(1.970503)", BesselI0(1.970503), 2.233306705, worst
compareRel "I0(2.369916)", BesselI0(2.369916), 2.981058221, worst
compareRel "I0(2.749343)", BesselI0(2.749343), 3.993840617, worst
compareRel "I0(-3.097791)", BesselI0(-3.097791), 5.284944407, worst
compareRel "I0(3.410674)", BesselI0(3.410674), 6.845628286, worst
compareRel "I0(3.675541)", BesselI0(3.675541), 8.558744454, worst
compareRel "I0(3.889478)", BesselI0(3.889478), 10.27562361, worst
compareRel "I0(-4.048445)", BesselI0(-4.048445), 11.78527478, worst
compareRel "I0(4.0726197)", BesselI0(4.07262), 12.03455148, worst  ' joint
compareRel "I0(4.190386)", BesselI0(4.190386), 13.33040931, worst
compareRel "I0(4.390003)", BesselI0(4.390003), 15.870654, worst
compareRel "I0(4.654062)", BesselI0(4.654062), 20.02717304, worst
compareRel "I0(-5.045880)", BesselI0(-5.04588), 28.38027712, worst
compareRel "I0(5.552404)", BesselI0(5.552404), 44.76661981, worst
compareRel "I0(6.162154)", BesselI0(6.162154), 77.97032084, worst
compareRel "I0(6.987602)", BesselI0(6.987602), 166.6705361, worst
compareRel "I0(-8.059230)", BesselI0(-8.05923), 451.9235213, worst
compareRel "I0(9.596594)", BesselI0(9.596594), 1921.275799, worst
compareRel "I0(12.09783)", BesselI0(12.09783), 20809.82354, worst
compareRel "I0(16.76299)", BesselI0(16.76299), 1871328.843, worst
compareRel "I0(-24.28060)", BesselI0(-24.2806), 2854269135#, worst
compareRel "I0(42.51715)", BesselI0(42.51715), 1.790127002E+17, worst
compareRel "I0(107.1033)", BesselI0(107.1033), 1.261508197E+45, worst
compareRel "I0(-709.7)", BesselI0(-709.7), 2.478808727E+306, worst
teeOut "Largest BesselI0 relative error: " & Format(worst, "0.000000E-0")

teeOut ""
If Abs(worst) < limit Then
  teeOut "Success - error less than the limit of " & _
         Format(limit, "0.0000E-0")
Else
  teeOut "FAILURE! That's too large - should be less than " & _
         Format(limit, "0.0000E-0")
End If
teeOut "--- Test complete ---"
Close #m_ofi
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
teeOut str & "  approx " & Format(approx, "0.00000000000E-0") & _
       "  exact " & Format(exact, "0.00000000000E-0") & _
       "  relErr " & Format(relErr, "0.000E-0")
End Sub

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Private Sub teeOut(ByRef str As String)
' unit test support routine - John Trenholme - 3 Sep 2002

Debug.Print str  ' works only if in VB[A] editor environment
If m_ofi <> 0 Then Print #m_ofi, str
End Sub

#End If
'-------------------------------- end of file ----------------------------------

Attribute VB_Name = "AdaptInt1_"
Attribute VB_Description = "Module holding adaptive Simpson integrator routine that always integrates 'adaptIntFunc'. Devised & coded by John Trenholme."
'
'###############################################################################
'#
'# Visual Basic 6 & VBA code module "AdaptInt1.bas"
'#
'# Adaptive Simpson integration of any "reasonable" function.
'# This version always integrates 'adaptIntFunc'.
'#
'#  Exports the routines:
'#    Function adaptInt1
'#    Function adaptInt1Version
'#    Function adaptIntFunc
'#    Sub Test_AdaptInt1 (if UnitTest is True)
'#
'# Devised and coded by John Trenholme -  begun 2005-07-26
'# Based on Fortran code of September 1970
'#
'###############################################################################

Option Base 0
Option Compare Binary
Option Explicit
Option Private Module  ' Don't allow visibility outside this Project

Private Const Version_c As String = "2006-07-20"
Private Const m_c As String = "AdaptInt1"  ' module name

' WARNING! set False when running actual code - 'adaptIntFunc' redefined if True
#Const UnitTest = False  ' set True to enable unit test code
' #Const UnitTest = True

#If UnitTest Then
#Const VBA = True        ' set True in Excel (etc.) VBA ; False in VB6
Private ofi_m As Integer  ' output file index used by unit-test routines
Const Ripl_c As Double = 10#   ' number of spikes per unit change of x
Const Size_c As Double = 0.99  ' the closer this is to 1, the sharper the spikes
#End If

Private Const EOL As String = vbNewLine  ' short form; works on both PC and Mac

'-------------------------------------------------------------------------------
' Use the following return-code index values to make your code more readable.
Public Const aii_ResultCode As Long = 0&
Public Const aii_FunctionCalls = 1&
Public Const aii_MaxStackDepth = 2&
Public Const aii_SmallestIntervalStart = 3&
Public Const aii_SmallestIntervalSpan = 4&
Public Const aii_LastPointAccepted = 5&
Public Const aii_RoutineCalls = 6&

' Use the following result code values to make your code more readable.
' They describe the possible values returned in report(aii_ResultCode).
Public Const aie_NoError As Long = 0&
Public Const aie_TooManyCalls As Long = 1&
Public Const aie_IntervalUnderflow As Long = 2&

'===============================================================================
Public Function adaptInt1( _
  ByVal x1 As Double, _
  ByVal x2 As Double, _
  ByVal absErr As Double, _
  ByVal relErr As Double, _
  ByVal dxMax As Double, _
  ByVal maxCalls As Long, _
  ByRef report() As Double) _
As Double
Attribute adaptInt1.VB_Description = "Adaptive Simpson integrator routine that always integrates 'adaptIntFunc'. See comments in code for usage."
' Adaptive Simpson integration of a user-specified function.  Because VB6 & VBA
' can't pass functions by name, we always integrate 'adaptIntFunc'.
'
' Note: by using function objects that define a common interface, VB[6,A] can
' handle polymorphic functions, but the COM overhead on the calls makes
' calculations slower.
'
' If you want to use this version with different integrands, code 'adaptIntFunc'
' so that it examines a project-global index and switches to appropriate code.
'
' Inputs:
'    x1        Start of integration interval.
'    x2        End of integration interval.
'    absErr    Absolute error tolerance per sub-interval (try expected * 1E-10).
'    relErr    Relative error tolerance per sub-interval (usually use 0).
'    dxMax     Maximum sub-interval size which will be used (try |x2-x1| / 33).
'    maxCalls  Maximum number of function calls allowed (try 30,000).
'
' Output:
'    report()  Array with result code and diagnostic information
'              in its elements (see aii_xx variables above for index values):
'        report(aii_ResultCode)             Result (see aie_xx variables above).
'            aie_NoError            looks OK
'            aie_TooManyCalls       too many calls       <- integrand may be
'            aie_IntervalUnderflow  interval underflow   <- singular for these
'        report(aii_FunctionCalls)          Number of function calls used.
'        report(aii_MaxStackDepth)          Maximum stack depth reached.
'        report(aii_SmallestIntervalStart)  Smallest sub-interval start.
'        report(aii_SmallestIntervalSpan)   Smallest sub-interval span.
'        report(aii_LastPointAccepted)      End of last accepted interval.
'        report(aii_RoutineCalls)           Total calls of this routine.
'
' Return:
'    the function value returned is an estimate of the integral of func(x)
'    from x1 to x2. The estimate will be "good" (supplying a result that is
'    at least related to the user-supplied error tolerances absErr and relErr)
'    if 'report(aii_ResultCode)' is 'aie_NoError'.
'
'    The estimate will be of dubious accuracy if this isn't true; the exact
'    value of 'report(aii_ResultCode)' tells what the problem was (see
'    aie_xx variables above). The possibilities are that there were more
'    function calls than the user-specified maximum, or that the interval size
'    became too small to represent in a Double. Both these situations can arise
'    from singular integrands, so check for that if you get these errors.
'
' Usage example:
'   Dim report(aii_ResultCode To aii_RoutineCalls) As Double, res As Double
'   res = adaptInt(1.5, 2.75, 0.0000000001, 0#, 0.03, 30000, report())
'   If report(aii_ResultCode) <> aie_NoError Then
'     {there was an error - handle it}
'   End If
'
' This routine integrates a function by making a 3-point Simpson estimate, and
' then adding two new points between the initial points and estimating the
' change in the first estimate. If the change is too large, it puts the right
' half interval on a stack, and repeats the add-points-estimate-change
' procedure on the left half interval. Once an interval becomes small enough to
' satisfy the change-small-enough criterion, a new interval is popped off the
' stack and the process repeats. When the stack is empty, the job is done. This
' method concentrates function evaluations where they are needed, rather than
' just adding more points everywhere.
'
' Because the error tolerances apply to sub-intervals, rather than the overall
' integral, the error in the integral will be roughly equal to the specified
' tolerance when it is loose (say 0.001 to 0.000001) but can often become
' larger than the specified error when the specification is tight. This is
' because addition of sub-interval corrections is subject to roundoff. If the
' integrand is especially difficult (strong oscillations or isolated spikes)
' the error in the integral can easily be two or three orders of magnitude
' larger than the specification for relative-error requests of 1E-10 or 1E-12.
' See the unit-test routine for an example of integration of a spiky function.
'
' It is best to specify only the absolute error (setting the relative error to
' zero), unless the integrand will be roughly constant but at an unknown level.
' This is because control by relative error will cause the routine to spend much
' effort on subintervals in regions where a varying integrand is relatively
' small or zero, without increasing the overall accuracy.
'
' Note that the first error tolerance to be met will stop further integration
' of a subinterval, so setting one error tolerance to a finite value and the
' other to zero means the finite value will be the limit used.
'
' The number of function evaluations required rises steadily as the requested
' error decreases, and can be thousands or tens of thousands for 1E-12 error.
' Do not ask for the absolute error to be below about 1E-14 times the result,
' or for the relative error to be below about 1E-14, since that will lead
' to large roundoff errors when adding small-interval corrections, and thus
' extra work for no extra precision.
'
' The "dxMax" value can be used to require a minimum sampling density of
' the integration interval, before subdivision is attempted. This will help
' the routine see narrow features that might otherwise be missed. For example,
' strongly oscillating integrands should have "dxMax" set to less than an
' oscillation period. There will be at least 5 function evaluations no matter
' what "dxMax" is set to. If "dxMax" is zero, no density limit is imposed.
'
' Note that the integral from x2 to x1 is the negative of that from x1 to x2,
' in accordance with standard mathematical practice.

Dim del As Double
Dim Done As Boolean
Dim dx As Double
Dim dxLoc As Double
Dim dxMin As Double
Dim fA As Double
Dim fB As Double
Dim fC As Double
Dim fD As Double
Dim fE As Double
Dim jDeep As Long
Dim jStk As Long
Dim kalls As Long
Dim sApp As Double
Dim xA As Double
Dim xB As Double
Dim xE As Double
Dim xGood As Double

Dim stak() As Double

Static calls_s As Double  ' number of times this routine has been called
calls_s = calls_s + 1#    ' stops adding at 9,007,199,254,740,992 calls

' default dxMax value
If dxMax = 0# Then dxMax = x2 - x1

' impose minimal sanity
absErr = Abs(absErr)
dxMax = 0.25 * Abs(dxMax)  ' we keep 1/4 of the actual interval internally

' allocate initial array space for stack; depth 25 covers almost all cases
ReDim stak(0& To 2&, 0& To 24&)
' initialize active-interval arguments and function values, and call counter
xA = x1
fA = adaptIntFunc(xA)               ' beginning function value
xE = x2
fE = adaptIntFunc(xE)               ' end function value
fC = adaptIntFunc(0.5 * (xA + xE))  ' center function value
kalls = 3&
' evaluate initial Simpson's rule approximation for integral value (times 3)
sApp = 0.5 * (xE - xA) * (fA + 4# * fC + fE)
' initialize stack index, maximum-depth value and interval stuff
jStk = 0&
jDeep = jStk
dxLoc = xA                             ' start of smallest sub-interval
dxMin = 0.25 * (xE - xA)               ' smallest sub-interval so far
xGood = xA                             ' end of accepted-interval region
' start main iteration loop - active points are labeled A - B - C - D - E
Done = False
Do
  ' subdivide active interval
  dx = 0.25 * (xE - xA)                ' 1/4 of sub-interval length
  If Abs(dxMin) > Abs(dx) Then         ' keep track of shortest interval
    dxMin = dx                         ' note: really 4 times this
    dxLoc = xA
  End If
  fD = adaptIntFunc(xE - dx)           ' func. value in center of C-E interval
  xB = xA + dx                         ' xB used later to check for underflow
  fB = adaptIntFunc(xB)                ' func. value in center of A-C interval
  kalls = kalls + 2&
  ' calculate correction to integral after subdivision (times 3, for speed)
  del = dx * (4# * (fB + fD) - 6# * fC - fA - fE)
  sApp = sApp + del                    ' improve running estimate of integral
  ' test for possible termination conditions
  If kalls > maxCalls Then             ' fail if too many calls were made
    Done = True
    report(aii_ResultCode) = aie_TooManyCalls
  ElseIf xB = xA Then                  ' fail if sub-interval size underflowed
    Done = True
    report(aii_ResultCode) = aie_IntervalUnderflow
  ' done with this interval? If so, pop an interval, or quit if stack is empty
  ' done = ( met-relative-err Or met-absolute-err) And interval-short-enough
  ElseIf ((Abs(del) <= Abs(relErr * sApp)) Or _
          (Abs(del) <= absErr)) And _
          (Abs(dx) <= dxMax) Then
    xGood = xE                         ' set accepted-interval high-water mark
    If jStk <= 0& Then                 ' the stack is empty and we are done
      Done = True
      report(aii_ResultCode) = aie_NoError
    Else                               ' there is an interval on the stack
      xA = xE                          ' set start of interval to end of prev.
      fA = fE
      jStk = jStk - 1&                 ' pop rest of interval off of stack
      fC = stak(0&, jStk)
      xE = stak(1&, jStk)
      fE = stak(2&, jStk)
    End If
  ' not done - this sub-interval needs more work - try to subdivide it
  Else
    If jStk > UBound(stak, 2&) Then
      ' stack is full - add more slots - singular?
      ReDim Preserve stak(0& To 2&, 0& To UBound(stak, 2&) + 25&)
    End If
    ' push end half-interval onto stack
    stak(0&, jStk) = fD
    stak(1&, jStk) = xE
    stak(2&, jStk) = fE
    jStk = jStk + 1&
    ' update deepest-stack value (note 1 means 1 entry, etc.)
    If jDeep < jStk Then jDeep = jStk
    ' set argument & function values so beginning interval -> active interval
    ' this is the "tail-recursion elimination" move
    xE = 0.5 * (xA + xE)
    fE = fC
    fC = fB
  End If
Loop Until Done                        ' keep it up until we are done

' release dynamically-allocated stack memory
Erase stak

' set diagnostic values for user perusal
report(aii_FunctionCalls) = kalls
report(aii_MaxStackDepth) = jDeep
report(aii_SmallestIntervalStart) = dxLoc
report(aii_SmallestIntervalSpan) = 4# * dxMin  ' fix up smallest-interval span
report(aii_LastPointAccepted) = xGood   ' end of good region
report(aii_RoutineCalls) = calls_s
adaptInt1 = sApp / 3#                   ' we found 3 times the actual result
End Function

'===============================================================================
Public Function adaptInt1Version() As String
Attribute adaptInt1Version.VB_Description = "The date of the latest revision to this module as a string in the format 'YYYY-MM-DD' such as 2004-06-18. It's a function so Excel etc. can use it."
' The date of the latest revision to this module as a string in the format
' 'YYYY-MM-DD' such as 2004-06-18. It's a function so Excel etc. can use it.
adaptInt1Version = Version_c
End Function

' WARNING! set UnitTest False for actual code - 'adaptIntFunc' redefined if True
#If Not UnitTest Then

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function adaptIntFunc(ByVal x As Double) As Double
Attribute adaptIntFunc.VB_Description = "Function that is integrated by 'adaptInt1'. Must be user-coded anew for each different integrand."
' This is the function that is integrated. User must code it for each integrand.
'
' Get fraction of crud that fails at this fluence from expectation value of CDF
' of log-normal.
adaptIntFunc = logNormalCDF(x, grMeanInit_p, grStdDevInit_p) * _
  riPDFNv(x, grMeanBeam_p, grStdDevBeam_p, grShotCountBeam_p, _
  grVaryingFractionBeam_p)
End Function

#Else

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'&
'& Unit test
'&
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Public Function adaptIntFunc(ByVal x As Double) As Double
Attribute adaptIntFunc.VB_Description = "This unit-test integrand function has a number of narrow spikes. It is hard to integrate. The value of its integral from 1 to 2 is unity."
' This function has a number of narrow spikes. It is hard to integrate. The
' value of its integral from 1 to 2 is unity.
' Pi written as sum to avoid 15-digit cutoff when writing to & reading from file
Const Pi_c As Double = 3.1415926 + 5.35897932E-08
adaptIntFunc = Sqr(1# - Size_c ^ 2) / (1 + Size_c * Sin(Ripl_c * 2# * Pi_c * x))
End Function

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Public Sub Test_AdaptInt1()
Attribute Test_AdaptInt1.VB_Description = "Unit-test routine for this module. Sends results to a file and to the immediate window (if in IDE)."
' Main unit test routine for this module.

' To run the test from VB6, enter this routine's name (above) in the Immediate
' window (if the Immediate window is not open, use View.. or Ctrl-G to open it).
' To run this routine from VBA, put the cursor somewhere in it and hit F5.
' To run this routine from somewhere in a code, call it.

' The output will be in the file 'Test_AdaptInt1.txt' on disk, and in the
' immediate window if in the VB[6,A] editor.

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
#Else  ' this is VB6
  ' note: this is the project folder if in VB6 IDE; EXE folder if stand-alone
  path = App.path
#End If
If Right$(path, 1) <> "\" Then path = path & "\"  ' only C:\ etc. have "\"
ofs = path & "Test_" & m_c & ".txt"

ofi_m = FreeFile
On Error Resume Next
Open ofs For Output Access Write Lock Read Write As #ofi_m  ' output file
If Err.Number <> 0 Then
  On Error GoTo 0
  ofi_m = 0  ' file did not open - don't use it
  MsgBox "ERROR - unable to open unit-test output file:" & EOL & EOL & _
    """" & ofs & """" & EOL & EOL & _
    "No unit-test output will be written to file.", _
    vbOKOnly Or vbExclamation, m_c & " Unit Test"
End If
On Error GoTo 0

teeOut "######## Test of " & m_c & " routines at " & Now()
teeOut "Code version: " & Version_c
teeOut

Dim report(aii_ResultCode To aii_RoutineCalls) As Double, res As Double

Dim limit As Double
Dim nWarn As Long
Dim worst As Double
nWarn = 0&

teeOut "All cases integrate the function:"
teeOut "  Sqr(1 -  " & Size_c & "^2) / (1 + " & Size_c & " * Sin(" & Ripl_c & _
  " * 2# * Pi_c * x))"
teeOut "from x = 1 to x = 2. This integral is exactly equal to 1."
teeOut

' error 1E-5
worst = 0#
Dim x1 As Double, x2 As Double, absEr As Double, relEr As Double
Dim dxMax As Double, callMax As Long
x1 = 1#
x2 = 2#
relEr = 0#
dxMax = 0.25
callMax = 40000

absEr = 0.00001
res = adaptInt1(x1, x2, absEr, relEr, dxMax, callMax, report())
If report(aii_ResultCode) <> aie_NoError Then res = -report(aii_ResultCode)
compareAbs "Requested absErr per subinterval: " & absEr, res, 1#, worst
teeOut "    error tolerances:  absolute " & absEr & "  relative " & relEr
teeOut "    max interval: " & dxMax & "  max calls: " & callMax
teeOut "    function calls:     " & report(aii_FunctionCalls)
teeOut "    max stack depth:    " & report(aii_MaxStackDepth)
teeOut "    min interval start: " & report(aii_SmallestIntervalStart)
teeOut "    min interval end:   " & _
  report(aii_SmallestIntervalStart) + report(aii_SmallestIntervalSpan)
teeOut "    min interval span:  " & report(aii_SmallestIntervalSpan)
limit = 0.000008
If Abs(worst) > limit Then
  teeOut "WARNING! That's too large - should be less than " & _
         Format(limit, "0.0000E-0")
  nWarn = nWarn + 1&
End If
teeOut

' error 1E-7
worst = 0#
absEr = 0.0000001
res = adaptInt1(x1, x2, absEr, relEr, dxMax, callMax, report())
If report(aii_ResultCode) <> aie_NoError Then res = -report(aii_ResultCode)
compareAbs "Requested absErr per subinterval: " & absEr, res, 1#, worst
teeOut "    error tolerances:  absolute " & absEr & "  relative " & relEr
teeOut "    max interval: " & dxMax & "  max calls: " & callMax
teeOut "    function calls:     " & report(aii_FunctionCalls)
teeOut "    max stack depth:    " & report(aii_MaxStackDepth)
teeOut "    min interval start: " & report(aii_SmallestIntervalStart)
teeOut "    min interval end:   " & _
  report(aii_SmallestIntervalStart) + report(aii_SmallestIntervalSpan)
teeOut "    min interval span:  " & report(aii_SmallestIntervalSpan)
limit = 0.0000002
If Abs(worst) > limit Then
  teeOut "WARNING! That's too large - should be less than " & _
         Format(limit, "0.0000E-0")
  nWarn = nWarn + 1&
End If
teeOut

' error 1E-9
worst = 0#
absEr = 0.000000001
res = adaptInt1(x1, x2, absEr, relEr, dxMax, callMax, report())
If report(aii_ResultCode) <> aie_NoError Then res = -report(aii_ResultCode)
compareAbs "Requested absErr per subinterval: " & absEr, res, 1#, worst
teeOut "    error tolerances:  absolute " & absEr & "  relative " & relEr
teeOut "    max interval: " & dxMax & "  max calls: " & callMax
teeOut "    function calls:     " & report(aii_FunctionCalls)
teeOut "    max stack depth:    " & report(aii_MaxStackDepth)
teeOut "    min interval start: " & report(aii_SmallestIntervalStart)
teeOut "    min interval end:   " & _
  report(aii_SmallestIntervalStart) + report(aii_SmallestIntervalSpan)
teeOut "    min interval span:  " & report(aii_SmallestIntervalSpan)
limit = 0.000000005
If Abs(worst) > limit Then
  teeOut "WARNING! That's too large - should be less than " & _
         Format(limit, "0.0000E-0")
  nWarn = nWarn + 1&
End If
teeOut

' error 1E-11
worst = 0#
absEr = 0.00000000001
res = adaptInt1(x1, x2, absEr, relEr, dxMax, callMax, report())
If report(aii_ResultCode) <> aie_NoError Then res = -report(aii_ResultCode)
compareAbs "Requested absErr per subinterval: " & absEr, res, 1#, worst
teeOut "    error tolerances:  absolute " & absEr & "  relative " & relEr
teeOut "    max interval: " & dxMax & "  max calls: " & callMax
teeOut "    function calls:     " & report(aii_FunctionCalls)
teeOut "    max stack depth:    " & report(aii_MaxStackDepth)
teeOut "    min interval start: " & report(aii_SmallestIntervalStart)
teeOut "    min interval end:   " & _
  report(aii_SmallestIntervalStart) + report(aii_SmallestIntervalSpan)
teeOut "    min interval span:  " & report(aii_SmallestIntervalSpan)
limit = 0.0000000002
If Abs(worst) > limit Then
  teeOut "WARNING! That's too large - should be less than " & _
         Format(limit, "0.0000E-0")
  nWarn = nWarn + 1&
End If
teeOut

' error 1E-13
worst = 0#
absEr = 0.0000000000001
res = adaptInt1(x1, x2, absEr, relEr, dxMax, callMax, report())
If report(aii_ResultCode) <> aie_NoError Then res = -report(aii_ResultCode)
compareAbs "Requested absErr per subinterval: " & absEr, res, 1#, worst
teeOut "    error tolerances:  absolute " & absEr & "  relative " & relEr
teeOut "    max interval: " & dxMax & "  max calls: " & callMax
teeOut "    function calls:     " & report(aii_FunctionCalls)
teeOut "    max stack depth:    " & report(aii_MaxStackDepth)
teeOut "    min interval start: " & report(aii_SmallestIntervalStart)
teeOut "    min interval end:   " & _
  report(aii_SmallestIntervalStart) + report(aii_SmallestIntervalSpan)
teeOut "    min interval span:  " & report(aii_SmallestIntervalSpan)
limit = 0.000000000003
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
Private Sub teeOut(Optional ByRef str As String = "")
' Unit-test support routine. Sends 'str' to Immediate window (if in IDE) and to
' output file (if open).
' John Trenholme - 2006-07-20

Debug.Print str  ' works only if in VB[6,A] IDE editor environment
If ofi_m <> 0 Then Print #ofi_m, str
End Sub

#End If

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

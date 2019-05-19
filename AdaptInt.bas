Attribute VB_Name = "AdaptInt_"
Attribute VB_Description = "Module holding adaptive integrator routine that always integrates AdaptIntFunc, using a function index. Devised & coded by John Trenholme."
'
'###############################################################################
'# Visual Basic 6 source code file "AdaptInt.bas"
'#
'# Adaptive Simpson integration of any "reasonable" function.
'# This version always integrates AdaptIntFunc, using an index to select among
'# several functions in case there are more than one.
'#
'# Devised and coded by John Trenholme
'###############################################################################

Option Explicit

Public Const AdaptIntVersion As String = "2005-09-04"

' Use the following return code index values to make your code more readable.
Public Enum AdaptIntIndexEnum
  aiiResultCode = 0&
  aiiFunctionCalls = 1&
  aiiMaxStackDepth = 2&
  aiiSmallestIntervalStart = 3&
  aiiSmallestIntervalSpan = 4&
  aiiLastPointAccepted = 5&
  aiiRoutineCalls = 6&
End Enum

' Use the following result code values to make your code more readable.
' They describe the possible values returned in report(aiiResultCode).
Public Enum AdaptIntErrorEnum
  aieNoError = 0&
  aieTooManyCalls = 1&
  aieIntervalUnderflow = 2&
End Enum

'===============================================================================
Public Function AdaptInt( _
  ByVal funcIndex As Long, _
  ByVal x1 As Double, _
  ByVal x2 As Double, _
  ByVal absErr As Double, _
  ByVal relErr As Double, _
  ByVal dxMax As Double, _
  ByVal maxCalls As Long, _
  ByRef report() As Double) _
As Double
Attribute AdaptInt.VB_Description = "Adaptive integrator routine that always integrates AdaptIntFunc, using a function index. See comments in code for usage."
' Adaptive Simpson integration of a user-specified function.  Because VB6
' can't pass functions by name, we always integrate AdaptIntFunc, passing an
' index value to select among several functions in case there are more than one.
'
' Inputs:
'    funcIndex Index number of the function to be integrated.
'    x1        Beginning of integration interval.
'    x2        End of integration interval.
'    absErr    Absolute error tolerance per sub-interval (try expected * 1E-10).
'    relErr    Relative error tolerance per sub-interval (usually use 0).
'    dxMax     Maximum sub-interval size which will be used (try |x2-x1| / 33).
'    maxCalls  Maximum number of function calls allowed (try 30,000).
'
' Output:
'    report()  Array with result code and diagnostic information
'              in its elements (see AdaptIntIndexEnum above for index values):
'        report(aiiResultCode)             Result (see AdaptIntErrorEnum above).
'            aieNoError            looks OK
'            aieTooManyCalls       too many calls          <- integrand may be
'            aieIntervalUnderflow  interval size underflow <- singular for these
'        report(aiiFunctionCalls)          Number of function calls used.
'        report(aiiMaxStackDepth)          Maximum stack depth reached.
'        report(aiiSmallestIntervalStart)  Smallest sub-interval start.
'        report(aiiSmallestIntervalSpan)   Smallest sub-interval span.
'        report(aiiLastPointAccepted)      End of last accepted interval.
'        report(aiiRoutineCalls)           Total calls of this routine.
'
' Return:
'    the function value returned is an estimate of the integral of func(x)
'    from x1 to x2. The estimate will be "good" (to the user-supplied error
'    tolerances absErr and relErr) if 'report(aiiResultCode)' is 'aieNoError'.
'    The estimate will be of dubious accuracy if this isn't true; the exact
'    value of 'report(aiiResultCode)' tells what the problem was (see
'    AdaptIntErrorEnum above).
'
' Usage example:
'   Dim report(aiiResultCode To aiiRoutineCalls) As Double, res As Double
'   res = AdaptInt(j, 1.5, 2.75, 0.00000001, 0#, 0.1, 50000, report())
'   If report(aiiResultCode) <> aieNoError Then
'     {there was an error - handle it}
'   End If
'
' It is best to specify only the absolute error (setting the relative error to
' zero), unless the integrand will be roughly constant but at an unknown level.
' This is because control by relative error will cause the routine to spend much
' effort on subintervals in regions where a varying integrand is relatively small
' or zero, without increasing the overall accuracy.
'
' Note that the first error tolerance to be met will stop further integration
' of a subinterval, so setting one error tolerance to a finite value and the
' other to zero means the finite value will be the limit used.
'
' Do not ask for the absolute error to be below about 1E-14 times the result,
' or for the relative error to be below about 1E-14, since that will lead
' to underflow in the subinterval length and extra work for no extra precision.
'
' The "dxMax" value should be used to require a minimum sampling density of
' the integration interval, before subdivision is attempted. This will help
' the routine see narrow features that might otherwise be missed. For example,
' strongly oscillating integrands should have "dxMax" set to less than an
' oscillation period.
'
' Note that the integral from x2 to x1 is the negative of that from x1 to x2.

Dim del As Double
Dim done As Boolean
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

Static invoked As Long

' impose minimal sanity
absErr = Abs(absErr)
dxMax = Abs(dxMax)

invoked = invoked + 1&                 ' count total calls of this routine
' allocate array space for stack; depth 25 covers almost all cases
ReDim stak(0& To 2&, 0& To 24&)
' initialize active-interval arguments and function values, and call counter
xA = x1
fA = AdaptIntFunc(xA, funcIndex)       ' beginning function value
xE = x2
fE = AdaptIntFunc(xE, funcIndex)       ' end function value
fC = AdaptIntFunc(0.5 * (xA + xE), funcIndex)  ' center function value
kalls = 3&
' evaluate initial Simpson's rule approximation for integral value (times 3)
sApp = 0.5 * (xE - xA) * (fA + 4# * fC + fE)
' initialize stack index, maximum-depth value and interval stuff
jStk = 0&
jDeep = jStk
dxLoc = xA                             ' start of smallest sub-interval
dxMin = xE - xA                        ' smallest sub-interval so far
xGood = xA                             ' end of accepted-interval region
' start main iteration loop - active points are labeled A - B - C - D - E
done = False
Do
  ' subdivide active interval
  dx = 0.25 * (xE - xA)                ' 1/4 of sub-interval length
  If Abs(dxMin) > 2# * Abs(dx) Then    ' keep track of shortest interval
    dxMin = dx
    dxLoc = xA
  End If
  fD = AdaptIntFunc(xE - dx, funcIndex)  ' func. value in center of C-E interval
  xB = xA + dx                         ' xB used later to check for underflow
  fB = AdaptIntFunc(xB, funcIndex)     ' func. value in center of A-C interval
  kalls = kalls + 2&
  ' calculate correction to integral after subdivision (times 3, for speed)
  del = dx * (4# * (fB + fD) - 6# * fC - fA - fE)
  sApp = sApp + del                    ' improve running estimate of integral
  ' test for possible termination conditions
  If kalls > maxCalls Then             ' fail if too many calls were made
    done = True
    report(aiiResultCode) = aieTooManyCalls
  ElseIf xB = xA Then                  ' fail if sub-interval size underflowed
    done = True
    report(aiiResultCode) = aieIntervalUnderflow
  ' done with this interval? If so, pop an interval, or quit if stack is empty
  ElseIf ((Abs(del) <= Abs(relErr * sApp)) Or _
          (Abs(del) <= absErr)) And _
          (Abs(dx) <= dxMax) Then
    xGood = xE                         ' set accepted-interval high-water mark
    If jStk <= 0& Then                 ' the stack is empty and we are done
      done = True
      report(aiiResultCode) = aieNoError
    Else                               ' there is an interval on the stack
      xA = xE                          ' set beginning of interval
      fA = fE
      jStk = jStk - 1&                 ' pop rest of interval off of stack
      fC = stak(0&, jStk)
      xE = stak(1&, jStk)
      fE = stak(2&, jStk)
    End If
  ' not done - this sub-interval needs more work - try to subdivide it
  Else
    If jStk > UBound(stak, 2&) Then    ' stack is full (how is this possible?)
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
    xE = 0.5 * (xA + xE)
    fE = fC
    fC = fB
  End If
Loop Until done                        ' keep it up until we are done

' release dynamically-allocated stack memory
Erase stak

' set diagnostic values for user perusal
report(aiiFunctionCalls) = kalls
report(aiiMaxStackDepth) = jDeep
report(aiiSmallestIntervalStart) = dxLoc
report(aiiSmallestIntervalSpan) = 4# * dxMin  ' fix up smallest-interval span
report(aiiLastPointAccepted) = xGood   ' end of good region
report(aiiRoutineCalls) = invoked&
AdaptInt = sApp / 3#                   ' we found 3 times the actual result
End Function

'===============================================================================
Private Function AdaptIntFunc(ByVal x As Double, _
                              Optional ByVal j As Long = 0&) _
As Double
' This is the function that is integrated. If several different functions are
' to be integrated, identify which one by use of the index 'j' (as shown below).
'
' This constant is written as the sum of two parts to maintain full accuracy.
' VB[A] will otherwise truncate a digit when module is saved to file & loaded.
' Note that the last digit must be 1, not 2, for highest accuracy.
Const Pi As Double = 3.14159265 + 3.5897931E-09
Const TwoPi# = 2# * Pi
Dim ff#, pol#
'-- function(s) to be integrated (j% specifies which function for BASIC)
If j = 0& Then
  ff# = x# * jSubNu(TwoPi * q * x#, 0)
ElseIf j& = 1 Then
  ff# = x# * jSubNu(TwoPi * q * x#, 0) * (1# - x# * x# * (3# - x# * 2#))
ElseIf j& = 2 Then
  ff# = x# * jSubNu(TwoPi * q * x#, 0) * (1# + Cos(Pi * x#)) / 2#
ElseIf j& = 3 Then
  pol# = x# * (8.480424 - x# * 3.154144)
  pol# = 1# + x# * (0.1931894 + x# * (-6.51947 + pol#))
  ff# = x# * jSubNu(TwoPi * q * x#, 0) * pol#
Else
  Stop
End If
AdaptIntFunc = ff#
End Function

'----------------------------- end of file -----------------------------------

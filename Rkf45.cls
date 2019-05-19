VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Rkf45"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Runge-Kutta-Fehlberg Numerical Integration of Ordinary Differential Equations. Automatically adapts step size to meet requested error tolerance(s). Coded by John Trenholme."
'
'        8888888b.   888    d8P   8888888888   d8888   888888888
'        888   Y88b  888   d8P    888         d8P888   888
'        888    888  888  d8P     888        d8P 888   888
'        888   d88P  888d88K      8888888   d8P  888   8888888b.
'        8888888P"   8888888b     888      d88   888        "Y88b
'        888 T88b    888  Y88b    888      8888888888         888
'        888  T88b   888   Y88b   888            888   Y88b  d88P
'        888   T88b  888    Y88b  888            888    "Y8888P"
'
'###############################################################################
'#
'# Visual Basic 6 class file "Rkf45.cls"
'#
'# Runge-Kutta-Fehlberg Numerical Integration of Ordinary Differential Equations
'# Automatically adapts step size to meet requested error tolerance(s)
'#
'# Designed and coded by John Trenholme
'#
'###############################################################################
'
Option Explicit

Private Const version_m As String = "2006-03-23"

'*******************************************************************************
' Parameters used in derivative calculation
' Set these in 'setParams', use them in 'derivs'
Private TE_m As Boolean
Private sinSqrTheta_m As Double
'*******************************************************************************

' Derived quantity to be reported (if any)
Private reported_m As Double

' Class Properties
Private errNum_m As Integer  ' numeric error code
Private errVal_m As Double   ' value associated with error
Private h_m As Double        ' step size
Private nDown_m As Long      ' step-size-decrease count
Private nEvals_m As Long     ' number of derivative evaluations
Private nUp_m As Long        ' step-size-increase count
Private x_m As Double        ' present location

'These quantities need to be retained between calls to "integrate"
Private ea_m() As Double     ' local copy of absolute error array
Private er_m() As Double     ' local copy of relative error array
Private j1_m As Long         ' low array subscript on this problem
Private j2_m As Long         ' high array subscript on this problem
Private jHi_m As Long        ' high subscript on allocated arrays
Private jLo_m As Long        ' low subscript on allocated arrays
Private ff_m() As Double     ' temporary function storage (values not retained)
Private d0_m() As Double     ' derivative storage 0-6
Private d1_m() As Double
Private d2_m() As Double
Private d3_m() As Double
Private d4_m() As Double
Private d5_m() As Double

' Quantities used during one call to "integrate"
Private maxEvals_m As Long

' Symbolic constants for error codes; Public for caller use if desired
' Note that Enum values in [square brackets] can contain spaces
' Note that Enum values can be assigned to, and compared with, Integers & Longs
Public Enum Rkf45errors
  [Rkf45 No error] = 0
  [Rkf45 Tiny initial step] = 1
  [Rkf45 Too many derivative eval's] = 2
  [Rkf45 Tiny adjusted step] = 3
  [Rkf45 Didn't call init] = 4
End Enum

Private Const HugeLong_c As Long = 2147483647  ' largest possible positive Long

'===============================================================================
Private Sub Class_Initialize()
Attribute Class_Initialize.VB_Description = "Sets a few defaults. Do actual object initialization by calling 'init'."
' Set impossible default index values to force allocation of the
' dynamically-allocated arrays on the first call to "init".
jLo_m = -HugeLong_c
jHi_m = HugeLong_c
errNum_m = [Rkf45 Didn't call init]  ' error: not initialized
End Sub

'===============================================================================
Private Sub Class_Terminate()
Attribute Class_Terminate.VB_Description = "Releases dynamically-allocated array memory. Be sure to Set object = Nothing when done with it."
' Release dynamically-allocated array memory
Erase ea_m, er_m, ff_m, d0_m, d1_m, d2_m, d3_m, d4_m, d5_m
End Sub

'===============================================================================
Public Property Get DecreaseCount() As Double
Attribute DecreaseCount.VB_Description = "Number of times step size was decreased in last integration(s) (since 'init')."
' Number of times step size was decreased in last integration(s) (since init)
DecreaseCount = nDown_m
End Property

'===============================================================================
Public Property Get Deriv(ByVal index As Long) As Double
Attribute Deriv.VB_Description = "Derivative of dependent variable 'index' with respect to independent variable."
' Derivative of dependent variable 'index' with respect to independent variable
Deriv = d0_m(index)
End Property

'===============================================================================
Public Sub derivs( _
  ByVal x As Double, _
  ByRef f() As Double, _
  ByRef d() As Double)
Attribute derivs.VB_Description = "User-coded routine returns vector of derivatives 'd()', given independent variable 'x' and vector of dependent variables 'f()'."
' Given the independent variable 'x' and input function vector 'f()', returns
' the derivative vector 'd()' of the same length, where d(j) is the derivative
' of f(j) with respect to x. Must be recoded for each specific problem.
' Note: don't modify f().

' Const values must be written as sums to maintain 17-digit accuracy in file
Const piSqr4_c = 39.478417 + 6.04357434E-09       ' 4 * Pi^2
Dim temp As Double

' common factor
temp = refInd(x)
temp = piSqr4_c * (sinSqrTheta_m - temp * temp)
If TE_m Then
  ' we are handling one complex second-order equation as four first-order ones
  d(1&) = f(2&)         ' da/dx = p
  d(2&) = temp * f(1&)  ' dp/dx = 4*Pi^2*(sin(theta)^2-n(x)^2)*a
  d(3&) = f(4&)         ' db/dx = p
  d(4&) = temp * f(3&)  ' dp/dx = 4*Pi^2*(sin(theta)^2-n(x)^2)*b
Else
  Err.Raise 5, "Rkf.derivs", "In Rkf.derivs: can't do TM wave"
End If
End Sub

'===============================================================================
Public Property Get ErrorCode() As Integer
Attribute ErrorCode.VB_Description = "Numeric code corresponding to error state (see Enum 'Rkf45errors' for symbolic forms and values)."
' Numeric error codes (see Enum Rkf45errors):
'   0 if no error
'   1 if initial step is non-zero but near roundoff (probably caller error)
'   2 if maximum derivative evaluation count exceeded during integration
'   3 if step becomes "too small" during integration
'     this may indicate a singularity, or "stiff" equations
'   4 object not yet initialized via "init"
ErrorCode = errNum_m
End Property

'===============================================================================
Public Property Get ErrorText() As String
Attribute ErrorText.VB_Description = "Text description of error state. May have a related numeric value embedded."
' Text description of error state. May have a related numeric value embedded.
Select Case errNum_m
  Case [Rkf45 No error]
    ErrorText = "rkf45: no error - first accepted step size = " & errVal_m
  Case [Rkf45 Tiny initial step]
    ErrorText = "rkf45: initial step too small  h = " & errVal_m
  Case [Rkf45 Too many derivative eval's]
    ErrorText = "rkf45: derivative eval. count exceeded  nEvals = " & errVal_m
  Case [Rkf45 Tiny adjusted step]
    ErrorText = "rkf45: step made too small (singularity?)  h = " & errVal_m
  Case [Rkf45 Didn't call init]
    ErrorText = "rkf45: not yet initialized - call ""init"""
  Case Else
    ErrorText = "rkf45: oops - undefined error number " & errNum_m
End Select
End Property

'===============================================================================
Public Property Get EvalCount() As Double
Attribute EvalCount.VB_Description = "Number of derivative evaluations used in last integration(s) (since 'init')."
' Number of derivative evaluations used in last integration(s) (since init)
EvalCount = nEvals_m
End Property

'===============================================================================
Public Property Get IncreaseCount() As Double
Attribute IncreaseCount.VB_Description = "Number of times step size was increased in last integration(s) (since 'init')."
' Number of times step size was increased in last integration(s) (since init)
IncreaseCount = nUp_m
End Property

'===============================================================================
Public Function init( _
  ByVal xInit As Double, _
  ByRef funcInit() As Double, _
  ByVal stepInit As Double, _
  ByRef errAbs() As Double, _
  ByRef errRel() As Double, _
  ByVal maxEvals As Long)
Attribute init.VB_Description = "Call 'setParams', then call this, then call 'integrate' to integrate the ODE(s)."
' Call this first, then call "integrate" to integrate the ODE(s)

' The arguments are:
'   xInit       = start value of independent variable
'   funcInit()  = dependent variable(s) at start point
'   stepInit    = suggested initial step size - set zero to use a default value
'                 be cautious - if too large, error estimate may be wrong
'   errAbs()    = absolute errors - set to < 0.01 and > 1e-15 of "average"
'   errRel()    = relative errors for each variable - set < 0.01 and > 1e-15
'   maxEvals    = maximum derivative evaluations allowed (try 10000)

' Initialize the integrator
x_m = xInit
h_m = stepInit  ' will be checked and perhaps adjusted in "integrate"
maxEvals_m = maxEvals

' Check error-array bounds for conformance to funcInit()
j1_m = LBound(funcInit)
j2_m = UBound(funcInit)
If (j1_m > LBound(errRel)) Or (j2_m < UBound(errRel)) Then
  Err.Raise 9, "Rkf45.init", _
    "In Rkf45.init: errRel Dim values don't match funcInit"
End If
If (j1_m > LBound(errAbs)) Or (j2_m < UBound(errAbs)) Then
  Err.Raise 9, "Rkf45.init", _
    "In Rkf45.init: errAbs Dim values don't match funcInit"
End If

' New or different-sized problem -> (re)allocate array memory
If (j1_m <> jLo_m) Or (j2_m <> jHi_m) Then
  jLo_m = j1_m
  jHi_m = j2_m
  ReDim ff_m(jLo_m To jHi_m), ea_m(jLo_m To jHi_m), er_m(jLo_m To jHi_m), _
    d0_m(jLo_m To jHi_m), d1_m(jLo_m To jHi_m), d2_m(jLo_m To jHi_m), _
    d3_m(jLo_m To jHi_m), d4_m(jLo_m To jHi_m), d5_m(jLo_m To jHi_m), _
    d6_m(jLo_m To jHi_m)
End If

' Make sure all error values are positive & not tiny - silently fix if so
Dim allTiny As Boolean
allTiny = True
Const Tiny_c As Double = 0.000000000000001
Dim j As Long
For j = jLo_m To jHi_m
  ea_m(j) = Abs(errAbs(j))
  er_m(j) = Abs(errRel(j))
  If (ea_m(j) >= Tiny_c) Or (er_m(j) >= Tiny_c) Then allTiny = False
Next j
If allTiny Then
  For j = jLo_m To jHi_m
    ea_m(j) = Tiny_c
    er_m(j) = Tiny_c
  Next j
End If

' Evaluate initial derivative values
derivs x_m, funcInit, d0_m()
nEvals_m = 1&

' Indicate there is no error
errNum_m = [Rkf45 No error]
End Function

'===============================================================================
Public Sub integrate( _
  ByRef funcs() As Double, _
  ByVal xGoal As Double)
Attribute integrate.VB_Description = "Runge-Kutta-Fehlberg adaptive-stepsize integration of first-order ordinary differential equation(s). Requires user-coded routine 'derivs' to calculate derivatives. Call 'setParams' and 'init' first, then call this to integrate the ODE(s)."
' Call "init" first, then call this to integrate the ODE(s)

' Adaptive Runge-Kutta-Fehlberg 4-5 ODE integrator method of Watts & Shampine.
' See, for example, their article "The Art of Writing a Runge-Kutta Code, II"
' in Appl. Math. Comp., 5 (1979) pp. 93-121.

' For the coefficients, see the ACM TOMS Algorithm 504 (Subroutine GERK etc.)

' Carries out integration of one or more first-order differential equations,
' adjusting the step size to maintain user-specified error tolerances on
' absolute or relative error, whichever is greatest. Note that setting an
' error to zero will turn it off, so don't set abs and rel errors both to zero.

' The arguments are:
'   funcs()  = dependent variable(s) at present location - will be updated
'   xGoal    = desired end value of 'x' - tries to integrate to this point

' In addition to the arguments, you must supply a derivative-evaluation routine
'
' Sub derivs( _
'   ByVal x As Double, _
'   ByRef funcs() As Double, _
'   ByRef dervs() As Double)
'
' Given a value of the independent variable 'x' and values of the dependent
' variables funcs(), the routine should evaluate and return the derivatives
' dervs(), where dervs(j) is the derivative of funcs(j) with respect to x.

' To continue integrating after return, just set 'xGoal' to a new value &
' call again. This is useful if results are desired at specified intervals
' for plots or other usage.

' A calling sequence to go from x1 to x2 in one call might be:
'    Dim rk As Rkf45
'    Set rk = New Rkf45
'    ' Array bounds can be anything you want - code will use your values
'    Dim funx(1& To N), erRl(1& To N), erAb(1& To N)  ' for example
'    funx(1&) = 12.54  ' set initial values of dependent variables
'       ... etc ...
'    For j = 1& To N  ' set all errors the same (or set errors individually)
'      erAb(j) = errAbs
'      erRl(j) = errRel
'    Next j
'    rk.setParams 1.1, 2.2, ...
'    rk.init xIni, funx, 0.001 * (xEnd - xIni), erAb, erRl, 10000
'    rk.integrate funx(), xEnd
'    If rk.ErrorCode() <> [Rkf45 No error] Then
'      Err.Raise 5, "MyCode", rk.ErrorText()
'    End If
'    ' Final values are now in 'funx()'

' A calling sequence to go from xIni to xEnd in 'i' equal steps might be:
'    ... same as above through rk.setParams call ...
'    dx = (xEnd - xIni) / (i - 1)
'    rk.init xIni, funx, 0.01 * dx, erAb, erRl, 1000
'    For j = 2 To i  ' j is index of point at end of step; j = i is end point
'      xNxt = xEnd - (i - j) * dx  ' end point for this part of integration
'      rk.integrate funx(), xNxt
'      If rk.ErrorCode() <> [Rkf45 No error] Then
'        Err.Raise 5, "MyCode", rk.ErrorText()
'      End If
'    Next j
'    ' End-of-step values are now in 'funx()'

' Algorithm tuning constants - adjust at your own risk!
' It's not a good idea to make Shrink * Grow = 1, or too close to it
Const Shrink As Double = 0.55, _
      Grow As Double = 1.6, _
      Tiny As Double = 0.01

' Check to be sure "init" was called
If (jLo_m = -HugeLong_c) And (jHi_m = HugeLong_c) Then
  errNum_m = [Rkf45 Didn't call init]
  Exit Sub
End If

' Initialize step-size-change counters
nDown_m = 0
nUp_m = 0

' Local variables
Dim first As Boolean
Dim ej As Double, g As Double, hSave As Double, steps As Double, _
    t1 As Double, t2 As Double, tol As Double, xTest As Double
Dim j As Integer, jDown As Integer, nFail As Integer, nTiny As Integer

'--- code ---
' If initial step is zero, use a nominal value
If h_m = 0# Then h_m = 0.001 * (xGoal - x_m)
' Find how many steps of size 'h_m' it will take to reach the end
steps = (xGoal - x_m) / h_m
' If the step points away from the end, reverse the step (caller sign error)
If steps < 0# Then
  h_m = -h_m
  steps = -steps
End If
' Save the 'real' step size to be returned for possible restart
hSave = h_m
' Set for exactly one step to end if less than one step
If steps < 1# Then
  h_m = xGoal - x_m
  steps = 1#
End If
' Claim there is no error, and the related value is the starting step size
errNum_m = [Rkf45 No error]
errVal_m = h_m
' If 'h_m' is near roundoff on entry, bail out (probably caller error)
xTest = x_m + h_m / 16#
If x_m = xTest Then  ' test relies on compiler not optimizing it away
  errNum_m = [Rkf45 Tiny initial step]
  errVal_m = h_m
  GoTo Done
End If
' Step toward the end until we reach it or die trying
first = True  ' tells us when we have the first successful step
Do
  ' if we're at the end, return to caller
  ' use 0.5 here to allow for small-step roundoff
  If steps < 0.5 Then Exit Do
  ' if within one or two steps of end, set step size to land exactly on end
  ' if nearly 1 step set to exactly 1;  if approx. 2 steps set exactly 2
  If steps < 1.1 Then
    h_m = xGoal - x_m
  ElseIf steps < 1.9 Then
    h_m = 0.5 * (xGoal - x_m)
  End If
  ' count of number of times step was reduced for this move (failures)
  jDown = 0
  ' make trial steps - adjust step downward until error is within tolerances
  Do
    ' do Runge-Kutta-Fehlberg 4-5 move of size 'h_m' with grouping in temps
    g = h_m / 4#  ' sub-step size for R-K
    For j = j1_m To j2_m
      ff_m(j) = funcs(j) + g * d0_m(j)
    Next j
    derivs x_m + g, ff_m(), d1_m()

    g = h_m * 3# / 32#
    For j = j1_m To j2_m
      ff_m(j) = funcs(j) + g * (d0_m(j) + 3# * d1_m(j))
    Next j
    derivs x_m + h_m * 3# / 8#, ff_m(), d2_m()

    g = h_m / 2197#
    For j = j1_m To j2_m
      t1 = 7296# * d2_m(j) - 7200# * d1_m(j)
      ff_m(j) = funcs(j) + g * (1932# * d0_m(j) + t1)
    Next j
    derivs x_m + h_m * 12# / 13#, ff_m(), d3_m()

    g = h_m / 4104#
    For j = j1_m To j2_m
      t1 = 29440# * d2_m(j) - 32832# * d1_m(j)
      t2 = 8341# * d0_m(j) - 845# * d3_m(j)
      ff_m(j) = funcs(j) + g * (t1 + t2)
    Next j
    derivs x_m + h_m, ff_m(), d4_m()

    g = h_m / 20520#
    For j = j1_m To j2_m
      t1 = 41040# * d1_m(j) - 28352# * d2_m(j)
      t2 = 9295# * d3_m(j) - 5643# * d4_m(j)
      ff_m(j) = funcs(j) + g * (t1 + t2 - 6080# * d0_m(j))
    Next j
    derivs x_m + h_m / 2#, ff_m(), d5_m()

    g = h_m / 7618050#
    For j = j1_m To j2_m
      t1 = 3855735# * d3_m(j) - 1371249# * d4_m(j)
      t1 = 902880# * d0_m(j) + t1
      t2 = 3953664# * d2_m(j) + 277020# * d5_m(j)
      ff_m(j) = funcs(j) + g * (t1 + t2)
    Next j

    ' add the 5 derivative evaluations to the call counter
    nEvals_m = nEvals_m + 5

    ' find estimated errors for this trial step; count those above tolerance
    ' and those that are well below tolerance (for step-size growth)
    ' tolerance values for error are largest of relative and absolute errors
    ' relative error is with respect to average of first and last values so that
    ' passage through zero near a step endpoint doesn't cause unnecessary work
    nFail = 0
    nTiny = 0
    For j = j1_m To j2_m
      t1 = 21970# * d3_m(j) - 15048# * d4_m(j)
      t1 = t1 - 2090# * d0_m(j)
      t2 = 22528# * d2_m(j) - 27360# * d5_m(j)
      ej = Abs((t1 + t2) * h_m / 752400#)  ' estimated error for j'th component
      ' find error we have to beat - largest (yes!) of relative & absolute
      tol = er_m(j) * 0.5 * (Abs(funcs(j)) + Abs(ff_m(j)))  ' relative error
      If tol < ea_m(j) Then tol = ea_m(j)  ' use absolute error if larger
      If ej > tol Then nFail = nFail + 1   ' fail if error above tolerance
      If ej < Tiny * tol Then nTiny = nTiny + 1  ' test if way below tolerance
    Next j
    ' if no errors were above tolerance, go on to the next step
    If nFail = 0 Then Exit Do
    ' if there were too many calls, give up (leaving inaccurate step unmade)
    If nEvals_m >= maxEvals_m Then
      errNum_m = [Rkf45 Too many derivative eval's]
      errVal_m = nEvals_m
      GoTo Done
    End If
    ' otherwise add to shrink counts, shrink the step and try again
    jDown = jDown + 1
    nDown_m = nDown_m + 1
    ' shrink step by increasing amounts as failures add up (1/2, 1/3, 1/4, ...)
    h_m = h_m * Shrink / (jDown + 1)
    hSave = h_m                        ' save changed value of step size
    ' in case step gets near roundoff, bail out (singularity, stiff, ...)
    xTest = x_m + h_m / 16#
    If x_m = xTest Then  ' test relies on compiler not optimizing it away
      errNum_m = [Rkf45 Tiny adjusted step]
      errVal_m = h_m
      GoTo Done
    End If
  Loop
  ' step was accepted, so update x, f and d
  x_m = x_m + h_m
  For j = j1_m To j2_m  ' array assignment won't work here
    funcs(j) = ff_m(j)
  Next j
  derivs x_m, ff_m(), d0_m()
  nEvals_m = nEvals_m + 1
  ' if this is the first successful step, save the step size for no-error return
  If first Then
    first = False
    errVal_m = h_m
  End If
  ' if errors were all 'tiny', cautiously increase the step size
  If nTiny > j2_m - j1_m Then
    h_m = h_m * Grow
    hSave = h_m                        ' save changed value of step size
    nUp_m = nUp_m + 1
  End If
  ' find how many steps of size 'h_m' it will take to reach the end
  steps = (xGoal - x_m) / h_m
  ' if there were too many calls, give up (after making accurate step)
  If nEvals_m >= maxEvals_m Then
    errNum_m = [Rkf45 Too many derivative eval's]
    errVal_m = nEvals_m
    GoTo Done
  End If
Loop
' Single exit point for routine
Done:
h_m = hSave                     ' restore adjusted step for next call
End Sub

'===============================================================================
Public Property Get Reported() As Double
Attribute Reported.VB_Description = "Evaluate any quantity you want as 'side information' somewhere else (perhaps in 'derivs') and save it in 'reported_m'. Then get it here."
' Evaluate any quantity you want as "side information" somewhere else (perhaps
' in "derivs") and save it in reported_m. Then get it here.
Err.Raise 5, "Rkf45.Reported()", _
  "In Rkf45.Reported(): no reportable item defined."
Reported = reported_m
End Property

'===============================================================================
Public Sub setParams( _
  ByVal TE As Boolean, _
  ByVal thetaRad As Double)
Attribute setParams.VB_Description = "User-coded routine to set 'side' parameters used in derivative evaluation. Call this, then 'init', then 'integrate' to integrate the ODE(s)."
' User-coded routine to set "side" parameters used in derivative evaluation.
' Must be recoded for each problem.
TE_m = TE
sinSqrTheta_m = Sin(thetaRad)
sinSqrTheta_m = sinSqrTheta_m * sinSqrTheta_m
End Sub

'===============================================================================
Public Property Get StepNow() As Double
Attribute StepNow.VB_Description = "Size of step that was last used (not counting adjustments to hit end exactly). Useful when next setting the initial step value."
' Size of step that was last used (not counting adjustments to hit end exactly)
StepNow = h_m
End Property

'===============================================================================
Public Property Get Version() As Double
Attribute Version.VB_Description = "Version of this code, as a date in a string with the format 'YYYY-MM-DD'."
' Version of this code, as a date in a string with the format "YYYY-MM-DD"
Version = version_m
End Property

'===============================================================================
Public Property Get XNow() As Double
Attribute XNow.VB_Description = "Value of the independent variable after the integration. Should be equal to 'xGoal' argument to 'integrate', unless there was an error."
' Value of the independent variable after the integration. Should be equal to
' 'xGoal' argument to 'integrate', unless there was an error.
XNow = x_m
End Property

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

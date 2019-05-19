Attribute VB_Name = "RungeKuttaVec"
Attribute VB_Description = "Integrates a system of ODE's using a fixed-step Runge-Kutta method. Derivative evaluation is coded inline at point of call. Devised & coded by John Trenholme."
'
'###############################################################################
'#
'# Visual Basic (VB6 or VBA) Module file "RungeKuttaVec.bas"
'#
'# Simple fixed-step fourth-order Runge-Kutta integration of vector of first-
'# order ODE's. Exports a single-function-call routine that is used inside a
'# caller's loop (see usage information below). This lets the caller calculate
'# the values of the derivative vector inline, rather than having to write a
'# subroutine. Uses an internal state machine implementation.
'#
'# To integrate a single ODE, use the Module "RungeKuttaFun.bas". As an
'# alternative, just dimension the arrays (see usage below) to length one by,
'# for example, (1& To 1&).
'#
'# Devised & coded by John Trenholme - Started 2008-08-27
'#
'# The Module exports the following routines:
'#   Function RungeKuttaN
'#   Sub RungeKuttaNReset
'#   Function RungeKuttaVecVersion
'#
'###############################################################################

Option Base 0
Option Compare Binary
Option Explicit

Private Const Version_c As String = "2008-09-01"  ' latest revision date
Private Const M_c As String = "RungeKuttaVec!"     ' module name + separator
Private Const EOL As String = vbNewLine           ' shorthand

' User-friendly names of the state indices
' Because state_m (below) will initialize to 0, the initial state MUST equal 0
Private Const StateFirstCall_c As Long = 0&  ' initial state
Private Const State1_c As Long = 1&
Private Const State2_c As Long = 2&
Private Const State3_c As Long = 3&
Private Const State4_c As Long = 4&

' Module-global variables (retained between calls; initialize as 0)
Private h_m As Double     ' size of one step
Private j1_m As Long      ' lower bound of y()
Private j2_m As Long      ' upper bound of y()
Private k_m As Long       ' offset between arrays (if any)
Private k1_m As Long      ' lower bound of d()
Private k2_m As Long      ' upper bound of y()
Private s1_m() As Double  ' first estimate of delta-y vector
Private s2_m() As Double  ' second estimate of delta-y vector
Private s3_m() As Double  ' third estimate of delta-y vector
Private state_m As Long   ' index of the next state to be executed
Private step_m As Long    ' number of steps that have been calculated
Private x2_m As Double    ' desired end point of integration
Private y_m() As Double   ' value of y vector at start of step

'===============================================================================
Public Function RungeKuttaN( _
  ByRef x As Double, _
  ByRef y() As Double, _
  ByRef d() As Double, _
  ByVal xEnd As Double, _
  ByVal stepCount As Long) _
As Boolean
Attribute RungeKuttaN.VB_Description = "Main routine for carrying out integration of ODE system. See comments for usage."
' Integrates a vector of first-order ODE's from a starting point to a specified
' end point. A fixed step size is used, so you must be careful that the step
' size is small enough to give the accuracy you want (some experimentation on
' your part may be necessary to get this right). Usually, a few hundred steps
' per major change or oscillation gives absolute accuracy in the 1E-10 range.
' Note that fourth-order Runge-Kutta will oscillate wildly if the step size
' is too large, so err on the side of smaller steps.
'
' Most Runge-Kutta solver routines require the caller to code a subroutine that
' calculates the derivative vector, given the independent variable and the
' vector of function values. This state-machine implementation allows the caller
' to code the derivative evaluation inline, at the point of call. This makes
' it easier to maintain code, and also allows one routine to integrate an
' arbitrary number of ODE systems.
'
' Variables:
'   x          in-out  independent variable
'   y()        in-out  vector of function values
'   d()        in      vector of derivative values (functions of x and y())
'   xEnd       in      desired end point (can be in either direction from x)
'   stepCount  in      number of fixed-length steps to use in integration
'
' Usage:
'   Note: use any names you want for the variables - these are just examples
'   Note: any lower and upper bounds values may be used for the arrays, as long
'         as the number of functions equals the number of derivatives
'
'   Dim x As Double  ' independent variable
'   Dim y(1& To 2&) As Double, d(1& To 2&) As Double  ' functions & derivatives
'   x = 1.5  ' set initial position (must be a variable)
'   y(1&) = 0.25  ' set initial values of functions (must be variables)
'   y(2&) = -0.5
'   xEnd = 12.75  ' desired end point (can be a constant in the call)
'   nStep = 100&  ' number of steps to use (can be a constant in the call)
'   RungeKuttaNReset  ' not usually needed, but a wise precaution
'   ' loop to carry out the integration of the ODE's
'   Do
'     ' write code here to evaluate the vector of derivatives
'     ' as an alternative, replace 'd in the function call with a Function that
'     ' returns the derivative vector and leave this Loop empty
'     d(1&) = function1 of x, y()
'     d(2&) = function2 of x, y()
'   Loop Until RungeKuttaN(x, y, d, xEnd, nStep)
'   ' upon exit from the loop, x and y() have the desired end-point values
'
' The trick here is that each time RungeKuttaN is called inside the Loop, it
' uses the present values of x, y and d to advance x and y to their values at
' the end of the step. Then when the next pass around the Loop is made, the
' derivative gets a new value calculated from the new x and y values, and the
' next step can be made. RungeKuttaN returns False as long as the end point has
' not been reached, so the Loop will repeat. When x has reached the end value,
' RungeKuttaN returns True and control exits the loop, with the final values
' of x and y available for use.

Const R_c As String = "RungeKuttaN"

Dim retVal As Boolean
retVal = False  ' the most likely return value
Dim j As Long
Select Case state_m
  Case StateFirstCall_c, State1_c  ' maybe first call; first point in a step
    If state_m = StateFirstCall_c Then  ' first call; initialize
      If x = xEnd Then  ' probably a caller error; handle gracefully & silently
        RungeKuttaNReset
        RungeKuttaN = True
        Exit Function
      Else
        If stepCount < 1& Then  ' caller specified an impossible step count
          Err.Raise 5&, M_c & R_c, _
            "Argument ERROR in routine " & M_c & R_c & "[initialize]" & EOL & _
            "Step count must be > 0 but got " & stepCount & EOL & EOL & _
            "Cannot proceed. Sorry!"
        End If
        step_m = stepCount
        h_m = (xEnd - x) / stepCount  ' step size
        x2_m = xEnd
        ' check array bounds - must be same size, but bases can differ
        j1_m = LBound(y)
        j2_m = UBound(y)
        k1_m = LBound(d)
        k2_m = UBound(d)
        If (j2_m - j1_m) <> (k2_m - k1_m) Then
          Err.Raise 5&, M_c & R_c, _
            "Argument ERROR in routine " & M_c & R_c & "[initialize]" & EOL & _
            "Size of y(" & j1_m & " To " & j2_m & ") not same as" & EOL & _
            "size of d(" & k1_m & " To " & k2_m & ")" & EOL & EOL & _
            "Cannot proceed. Sorry!"
        End If
        k_m = k1_m - j1_m  ' offset between arrays (if any)
        ' allocate local function & derivative space - conform to y()
        ReDim y_m(j1_m To j2_m)
        ReDim s1_m(j1_m To j2_m)
        ReDim s2_m(j1_m To j2_m)
        ReDim s3_m(j1_m To j2_m)
      End If
    End If
    ' fall through to state 1 code

    ' first point in step - enter here with x, y, d(x, y)
    For j = j1_m To j2_m  ' save start-of-step values for use in later steps
      y_m(j) = y(j)
    Next j
    x = x + 0.5 * h_m  ' move to next x
    For j = j1_m To j2_m
      s1_m(j) = h_m * d(j + k_m)     ' estimate of change in y() for whole step
      y(j) = 0.5 * s1_m(j) + y_m(j)  ' update y()
    Next j
    state_m = State2_c  ' set next state

  Case State2_c  ' enter here with x+h/2, y+s1/2, d(x+h/2, y+s1/2)
    For j = j1_m To j2_m
      s2_m(j) = h_m * d(j + k_m)
      y(j) = 0.5 * s2_m(j) + y_m(j)
    Next j
    state_m = State3_c  ' set next state

  Case State3_c  ' enter here with x+h/2, y+s2/2, d(x+h/2, y+s2/2)
    x = x + 0.5 * h_m  ' move to end of step
    For j = j1_m To j2_m
      s3_m(j) = h_m * d(j + k_m)
      y(j) = s3_m(j) + y_m(j)
    Next j
    ' set x more accurately for start of next step, or windup (avoid roundoff)
    step_m = step_m - 1&  ' count down
    x = x2_m - step_m * h_m
    state_m = State4_c  ' set next state

  Case State4_c  ' enter here with x+h, y+s3, d(x+h, y+s3)
    Dim s4j As Double
    Const Inv3 As Double = 1# / 3#, Inv6 As Double = 1# / 6#
    For j = j1_m To j2_m  ' use weighted average to update y()
      s4j = h_m * d(j + k_m)  ' no need to save final change estimate
      y(j) = (s1_m(j) + s4j) * Inv6 + (s2_m(j) + s3_m(j)) * Inv3 + y_m(j)
    Next j
    If step_m > 0& Then  ' there are more steps to be done
      state_m = State1_c  ' set next state (beginning of new step)
    Else  ' this is the end point
      RungeKuttaNReset
      retVal = True
    End If

  Case Else
    Dim wrongState As Long
    wrongState = state_m
    RungeKuttaNReset
    Err.Raise 51&, M_c & R_c, _
      "Internal logic ERROR in routine " & M_c & R_c & " [Case Else]" & EOL & _
      "Tried to go to nonexistent state " & wrongState & EOL & EOL & _
      "Cannot proceed. Sorry!"
End Select
RungeKuttaN = retVal
End Function

'===============================================================================
Public Sub RungeKuttaNReset()
Attribute RungeKuttaNReset.VB_Description = "Reset the integration routine to its initial state. Call this if the main routine halted abnormally. Does not hurt, and may help, if called just before call to main routine."
' Initialize for a new run. Not needed if all has gone well, or on the first
' call, but it's a wise precaution to call this just before a new integration
' in case the internal state of 'RungeKuttaN' is somehow screwed up.
state_m = StateFirstCall_c
Erase y_m, s1_m, s2_m, s3_m
End Sub

'===============================================================================
Public Function RungeKuttaVecVersion()
Attribute RungeKuttaVecVersion.VB_Description = "Supplies the date of the latest revision to this code, as a string in the format ""yyyy-mm-dd"""
' Returns the date of the latest revision to this code, as a string in the
' format "yyyy-mm-dd"
RungeKuttaVecVersion = Version_c
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


Attribute VB_Name = "RungeKuttaFun"
Attribute VB_Description = "Integrates a single ODE using a fixed-step Runge-Kutta method. Derivative evaluation is coded inline at point of call. Devised & coded by John Trenholme."
'
'###############################################################################
'#
'# Visual Basic (VB6 or VBA) Module file "RungeKuttaFun.bas"
'#
'# Simple fixed-step fourth-order Runge-Kutta integration of a single first-
'# order ODE. Exports a single-function-call routine that is used inside a
'# caller's loop (see usage information below). This lets the caller calculate
'# the value of the derivative inline, rather than having to write a
'# subroutine. Uses an internal state machine implementation.
'#
'# To integrate a system of 2 or more ODE's, use the Module "RungeKuttaVec.bas".
'#
'# Devised & coded by John Trenholme - Started 2008-08-27
'#
'# The Module exports the following routines:
'#   Function RungeKutta1
'#   Sub RungeKutta1Reset
'#   Function RungeKutta1Results
'#   Function RungeKuttaFunVersion
'#
'###############################################################################

Option Base 0
Option Compare Binary
Option Explicit

Private Const Version_c As String = "2008-09-01"  ' latest revision date
Private Const M_c As String = "RungeKuttaFun!"     ' module name + separator
Private Const EOL As String = vbNewLine           ' shorthand

' User-friendly names of the state indices
' Because state_m (below) will initialize to 0, the initial state MUST equal 0
Private Const StateFirstCall_c As Long = 0&  ' initial state
Private Const State1_c As Long = 1&
Private Const State2_c As Long = 2&
Private Const State3_c As Long = 3&
Private Const State4_c As Long = 4&

' Module-global variables (retained between calls; initialize as 0)
Private h_m As Double         ' size of one step
Private s1_m As Double        ' first estimate of delta-y
Private s2_m As Double        ' second estimate of delta-y
Private s3_m As Double        ' third estimate of delta-y
Private state_m As Long       ' index of the next state to be executed
Private step_m As Long        ' number of steps that have been calculated
Private steps_m As Long       ' number of steps that were requested
Private values_m() As Double  ' x and y values at each step
Private x2_m As Double        ' desired end point of integration
Private y_m As Double         ' value of y at start of step

'===============================================================================
Public Function RungeKutta1( _
  ByRef x As Double, _
  ByRef y As Double, _
  ByVal d As Double, _
  ByVal xEnd As Double, _
  ByVal stepCount As Long) _
As Boolean
Attribute RungeKutta1.VB_Description = "Main routine for carrying out integration of ODE. See comments for usage."
' Integrates a single first-order ODE from a starting point to a specified
' end point. A fixed step size is used, so you must be careful that the step
' size is small enough to give the accuracy you want (some experimentation on
' your part may be necessary to get this right). Usually, a few hundred steps
' per major change or oscillation gives absolute accuracy in the 1E-10 range.
' Note that fourth-order Runge-Kutta will oscillate wildly if the step size
' is too large, so err on the side of smaller steps.
'
' Most Runge-Kutta solver routines require the caller to code a subroutine that
' calculates the derivative, given the independent variable and the function
' value. This state-machine implementation allows the caller to code the
' derivative evaluation inline, at the point of call. This makes it easier to
' maintain code, and also allows one routine to integrate an arbitrary number
' of ODE's.
'
' Variables:
'   x          in-out  independent variable
'   y          in-out  function value
'   d          in      derivative value (a function of x and y)
'   xEnd       in      desired end point (can be in either direction from x)
'   stepCount  in      number of fixed-length steps to use in integration
'
' Usage:
'   Note: use any names you want for the variables - these are just examples
'
'   Dim x As Double  ' independent variable
'   Dim y As Double, d As Double  ' function & derivative
'   x = 1.5  ' set initial position (must be a variable)
'   y = 0.25  ' set initial value of function (must be a variable)
'   xEnd = 12.75  ' desired end point (can be a constant in the call)
'   nStep = 100&  ' number of steps to use (can be a constant in the call)
'   RungeKutta1Reset  ' not usually needed, but a wise precaution
'   ' loop to carry out the integration of the ODE's
'   Do
'     ' write code here to evaluate the derivative
'     ' as an alternative, replace 'd in the function call with a Function that
'     ' returns the derivative and leave this Loop empty
'     d = function of x, y
'   Loop Until RungeKutta1(x, y, d, xEnd, nStep)
'   ' upon exit from the loop, x and y have the desired end-point values
'
' The trick here is that each time RungeKutta1 is called inside the Loop, it
' uses the present values of x, y and d to advance x and y to their values at
' the end of the step. Then when the next pass around the Loop is made, the
' derivative gets a new value calculated from the new x and y values, and the
' next step can be made. RungeKutta1 returns False as long as the end point has
' not been reached, so the Loop will repeat. When x has reached the end value,
' RungeKutta1 returns True and control exits the loop, with the final values
' of x and y available for use.

Const R_c As String = "RungeKutta1"

Dim retVal As Boolean
retVal = False  ' the most likely return value
Select Case state_m
  Case StateFirstCall_c, State1_c  ' maybe first call; first point in a step
    If state_m = StateFirstCall_c Then  ' first call; initialize
      If x = xEnd Then
        ' this is either a request to initialize or a caller error
        RungeKutta1Reset
        Erase values_m
        RungeKutta1 = True
        Exit Function
      Else
        If stepCount < 1& Then  ' caller specified an impossible step count
          Err.Raise 5&, M_c & R_c, _
            "Argument ERROR in routine " & M_c & R_c & "[initialize]" & EOL & _
            "Step count must be > 0 but got " & stepCount & EOL & EOL & _
            "Cannot proceed. Sorry!"
        End If
        step_m = 0&
        steps_m = stepCount
        h_m = (xEnd - x) / stepCount  ' step size
        x2_m = xEnd
        ReDim values_m(0& To 1&, 0& To stepCount)  ' space for results
        values_m(0&, 0&) = x  ' initial point
        values_m(1&, 0&) = y
      End If
    End If
    ' fall through to state 1 code

    ' first point in step - enter here with x, y, d(x, y)
    y_m = y               ' save first point in this step
    x = x + 0.5 * h_m     ' move to next x
    s1_m = h_m * d        ' estimate change in y for whole step
    y = 0.5 * s1_m + y_m  ' update y
    state_m = State2_c    ' set next state

  Case State2_c  ' enter here with x+h/2, y+s1/2, d(x+h/2, y+s1/2)
    s2_m = h_m * d
    y = 0.5 * s2_m + y_m
    state_m = State3_c    ' set next state

  Case State3_c  ' enter here with x+h/2, y+s2/2, d(x+h/2, y+s2/2)
    x = x + 0.5 * h_m     ' move to end of step
    s3_m = h_m * d
    y = s3_m + y_m
    ' set x more accurately for start of next step, or windup (avoid roundoff)
    step_m = step_m + 1&  ' count up
    x = x2_m - (steps_m - step_m) * h_m
    state_m = State4_c  ' set next state

  Case State4_c  ' enter here with x+h, y+s3, d(x+h, y+s3)
    Dim s4 As Double
    Const Inv3 As Double = 1# / 3#, Inv6 As Double = 1# / 6#
    ' use weighted average to update y()
    s4 = h_m * d  ' no need to save final change estimate between calls
    y = (s1_m + s4) * Inv6 + (s2_m + s3_m) * Inv3 + y_m
    values_m(0&, step_m) = x
    values_m(1&, step_m) = y
    If step_m < steps_m Then  ' there are more steps to be done
      state_m = State1_c  ' set next state (beginning of new step)
    Else  ' this is the end point
      RungeKutta1Reset
      retVal = True
    End If

  Case Else
    Dim wrongState As Long
    wrongState = state_m
    RungeKutta1Reset
    Erase values_m
    Err.Raise 51&, M_c & R_c, _
      "Internal logic ERROR in routine " & M_c & R_c & " [Case Else]" & EOL & _
      "Tried to go to nonexistent state " & wrongState & EOL & EOL & _
      "Cannot proceed. Sorry!"
End Select
RungeKutta1 = retVal
End Function

'===============================================================================
Public Sub RungeKutta1Reset()
Attribute RungeKutta1Reset.VB_Description = "Reset the integration routine to its initial state. Call this if the main routine halted abnormally. Does not hurt, and may help, if called just before call to main routine."
' Initialize for a new run. Not needed if all has gone well, or on the first
' call, but it's a wise precaution to call this just before a new integration
' in case the internal state of 'RungeKutta1' is somehow screwed up.
state_m = StateFirstCall_c
End Sub

'===============================================================================
Function RungeKutta1Results()
Attribute RungeKutta1Results.VB_Description = "Return a 2D array containing the results of the integration of the ODE. Indices (0=X, 1=Y), (0=initial, UBound=final)"
' Return an array containing the results of the integration of the ODE. The X
' values are in the array at index values (0, 0 To stepCount) and the Y values
' are in the array at index values (1, 0 To stepCount). The value 0 in the
' second index is the location of the initial values, and subsequent indices
' have the calculated values from step 1 to step 'stepCount'.
RungeKutta1Results = values_m
End Function

'===============================================================================
Public Function RungeKuttaFunVersion()
Attribute RungeKuttaFunVersion.VB_Description = "Supplies the date of the latest revision to this code, as a string in the format ""yyyy-mm-dd"""
' Return the date of the latest revision to this code, as a string in the
' format "yyyy-mm-dd"
RungeKuttaFunVersion = Version_c
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


Attribute VB_Name = "ZeroBrentMod"
Attribute VB_Description = "Zero-crossing finder using Brent's method, implemented as a single-expression-call state machine. Devised and coded by John Trenholme."
'
'###############################################################################
'#
'# Visual Basic 6 & VBA code Module file "ZeroBrentMod.bas"
'#
'# State-machine implementation of a zero-crossing finder using Brent's method.
'# See: Brent, Richard - 1971
'#   An Algorithm with Guaranteed Convergence for Finding a Zero of a Function
'# See: Brent, Richard - 1973
'#   Algorithms for Minimization Without Derivatives (book)
'#   Chapter 4  ISBN: 0-13-022335-2 (PDF version on Brent's web pages)
'# See: http://maths-people.anu.edu.au/~brent/
'# See: "Numerical Recipes in X" (where X is some computer language)
'# by W. Press et. al., Cambridge University Press, section 9.3
'#
'# Exports the routines:
'#   Function zeroBrent
'#   Function zeroBrentBestF
'#   Function zeroBrentBestX
'#   Function zeroBrentBracketWidth
'#   Function zeroBrentEvals
'#   Function zeroBrentHistory
'#   Function zeroBrentHistoryCodes
'#   Sub zeroBrentInit
'#   Function zeroBrentOtherF
'#   Function zeroBrentOtherX
'#   Function ZeroBrentVersion
'#   Function zeroBrentWhy
'#   Function zeroBrentWhyTexts
'#
'# Requires the module "UnitTestSupport.bas" if unit test code enabled
'#
'# Devised and coded by John Trenholme - begun 2006-09-05
'#
'###############################################################################

Option Base 0
Option Compare Binary
Option Explicit
'Option Private Module  ' Don't allow visibility outside this Project (if VBA)

Private Const Version_c As String = "2013-10-01"
Private Const M_c As String = "ZeroBrentMod[" & Version_c & "]."  ' module name

#Const UnitTest = True  ' set True to enable unit test code
' #Const UnitTest = False  ' set False to eliminate unit test code

' Text strings describing state of problem; they go into 'why_m'
' the leading number allows Val(Tn_c) to give a numeric error
Private Const T1_c As String = _
  "1 Found an exact zero at arg = %1"
Private Const T2_c As String = _
  "2 Bracket dropped below adjusted size of %1"
Private Const T3_c As String = _
  "3 Evaluations reached limit of %1"
Private Const T4_c As String = _
  "4 ERROR in call - expression values at initial points have the same sign"
Private Const T5_c As String = _
  "5 ERROR in call - initial-interval ends have the same location (zero length)"
Private Const T6_c As String = _
  "6 ERROR in internal logic - attempted transition to non-existent state %1"
Private Const T7_c As String = _
  "7 WARNING: 'zeroBrent' routine not finished yet"
Private Const T8_c As String = _
  "8 NOTE: 'zeroBrent' routine not yet called"
Private Const T9_c As String = _
  "9 ERROR: 'zeroBrentInit' routine not called before 'zeroBrent'"

Private Const EOL As String = vbNewLine  ' short form; works on both PC and Mac
' The largest possible Double
Private Const MaxDouble_c As Double = 1.79769313486231E+308 + 5.7E+293
' The smallest number that can cause a change when added to 1#
Private Const EpsDouble_c As Double = 2.22044604925031E-16  ' 2.0^(-52)

' User-friendly names of the state indices
' Because state_m (below) will initialize to 0, the initial state MUST equal 0
Private Const StateFirstPoint_c As Long = 0&  ' initial state
Private Const StateSecondPoint_c As Long = 1&
Private Const StateIterate_c As Long = 2&

' Default value for maximum number of expression evaluations
' This value will do almost any problem
Private Const EvalMaxDefault_c As Long = 200&

' Module-global variables (retained between calls; initialize as 0)
Private didOld_m As String   ' action taken before present function evaluation
Private errAbs_m As Double   ' crossing must be in bracket of this size to quit
Private evalMax_m As Long    ' maximum number of expression evaluations allowed
Private evals_m As Long      ' number of expression evaluations so far
Private fa_m As Double       ' expression at third point
Private fb_m As Double       ' expression at 'best' end of bracketing interval
Private fc_m As Double       ' expression at 'worst' end of  bracketing interval
Private forceState_m As Boolean  ' if True, force an illegal state error
Private hist_m As String     ' eval. sequence: B|b=bisect L=linear Q=quadratic
Private initialized_m As Boolean  ' has 'zeroBrentInit' been called first?
Private inits_m As Double    ' number of times zeroBrent has been initialized
Private nEntries_m As Long   ' number of characters in history string
Private state_m As Long      ' index of the next state to be executed
Private step_m As Double     ' step from xa_m to new point
Private stepOld_m As Double  ' previous value of step_m
Private why_m As String      ' reason for exit from routine
Private x2_m As Double       ' the other end of the initial search interval
Private xa_m As Double       ' variable at third point (often previous xb_m)
Private xb_m As Double       ' variable at 'best' end of bracketing interval
Private xc_m As Double       ' variable at 'worst' end of bracketing interval

'===============================================================================
Public Function zeroBrent( _
  ByRef x As Double, _
  ByVal f As Double) _
As Boolean
Attribute zeroBrent.VB_Description = "Finds a value of 'x' near a zero crossing of any expression 'f' that depends on 'x' (and perhaps other variables) within the region from 'x1' to 'x2'. Should be in a loop such as 'zeroBrentInit x, 1.2, 3.5: Do: Loop While zeroBrent(x, Cos(x) - Sqr(x))'. See usage in code listing."
' Allows the caller to localize the value of 'x' that gives a zero crossing of
' any expression 'f' that depends on an independent variable 'x' (and perhaps
' other fixed variables). Different expressions may be user-coded and used in
' the same code. If you want the value of 'x' where f(x) = a instead of zero,
' just get the zero crossing of f(x) - a.
'
' To be more precise, this routine finds a small interval where the expression
'  f(x) has different signs at the ends, or a single point where the expression
' value is zero. Thus a pole will be seen as a zero crossing, which is
' technically correct but can lead to confusing results.
'
' This looks for a zero crossing between initial variable values of 'x1' and
' 'x2', as specified to 'zeroBrentInit'. The signs of the expression values must
' be different at these points, or the value at one or both must be zero, or you
' will get an error. The routine systematically shortens the bracketing interval
' in which a zero crossing must lie, keeping the end signs different. It stops
' if an exact zero is found, or if the bracketing 'x' interval drops below the
' caller's value 'errAbs', or if the caller's evaluation-count limit 'evalMax'
' is exceeded, or if the size of the bracketing interval becomes so small that
' there are no, or few, available floating-point values between the end values.
'
' Note that 'x1' and 'x2' do not need to be in increasing numerical order on
' input. They just have to be at the ends of a finite interval.
'
' If multiple zero crossings exist in the initial interval, the routine picks
' arbitrarily from among points where zero is crossed in the same direction as
' a straight line between the initial points. If there is a continuous region
' of zero values, any point in the region may be picked.
'
' Returns True if still working (continue in caller's loop), or False if done
' (time to exit loop). See below for how this return value is used.
'
' Raises error 5 if both initial points have the same sign and are non-zero.
' Raises error 17 if 'zeroBrentInit' has not been called before 'zeroBrent'.
' Raises error 51 if there is an internal move to an undefined state.
'
' The return value from this routine should be used as the test value at the
' end of a "Do ... Loop" control structure. Usage is as follows:
'
'   Dim x as Double  ' will be passed by reference, so must be a variable
'   ...
'   ' initial values of search interval ends, absolute error in x, & call max
'   ' warning: do not change x by yourself until the loop exits
'   zeroBrentInit x, 1.2, 3.5, 0.0001, 60&  ' note: x = x1 = 1.2 after this call
'   Do
'   Loop While zeroBrent(x, Cos(x) - Sqr(x) + 1#)
'
' In this example 'x' is the independent variable (any name can be used; we use
' 'x' to make the example concrete). First, call 'zeroBrentInit' with the
' independent variable, two values to set the interval ends (1.2 and 3.5 in this
' example), the x-difference error tolerance, (0.0001 in this example; optional,
' and defaults to 0# - meaning to reduce the bracket to the minimum possible
' value consistent with roundoff), and the call-count maximum (also optional,
' and defaults to 200, which should solve even very hard problems). The first
' argument to 'zeroBrent' is the independent variable (again), and the second
' argument is an expression ("Cos(x) - Sqr(x) + 1#" in this example) that is the
' user-defined quantity whose zero crossing is desired. Note that the evaluation
' count will not be checked until 2 evaluations have taken place, so there will
' always be 1 or 2 evaluations). No code is needed in the body of the loop, but
' can exist if useful (see below).
'
' When the loop exits, 'x' and 'zeroBrentBestX()' hold the zero-crossing-bracket
' end point where the absolute value of the expression was least. In most cases,
' this will be the point closest to the zero crossing, and frequently will be
' much closer than the requested error. Other values of interest can be retrieved
' by calls to 'zeroBrentBestF()', 'zeroBrentOtherX()' etc. when done. If the
' value returned by 'zeroBrentBestF()' is zero, 'x' has the exact location of
' the zero crossing, as do 'zeroBrentBestX()' and  'zeroBrentOtherX()'; in
' addition, 'zeroBrentBracketWidth()' will be zero. If f(x) is not zero, the
' signs at 'x' (and 'zeroBrentBestX()') and 'zeroBrentOtherX()' are
' different and the zero crossing lies between these points. Recall that a
' discontinuous expression that has different signs at the point of
' discontinuity will be reported to have a zero crossing there. Thus, poles
' will be classified as zero crossings, the same as roots.
'
' This method works because the expression to be evaluated can be written in
' just one place if the routine is called repeatedly in a loop. The routine
' changes the value of 'x' after evaluating the expression (this can be done
' because 'x' is passed by reference), and returns the new value to get
' evaluation at a new location on the next pass through the loop. The routine
' keeps track of where it is in the algorithm using an internal static state
' variable, and switches its return value from True to False when the job is
' done, causing the user's loop to exit.
'
' The expression may depend on other parameters in addition to 'x', but only
' one parameter may be varied at a time.
'
' If the expression is too complicated to be coded 'in-line' as a single
' expression (perhaps because of conditional statements or For loops) it can be
' calculated inside the loop (since the updated value of 'x' is available
' anywhere inside the loop) and its value can be passed to 'zeroBrent':
'
'   Dim s as Double, t as Double
'   ...
'   'set initial value of 's' to start of search interval
'   zeroBrentInit s, 1#, 2#, 0.0001
'   Do
'     t = <result of some process, depending on 's', coded inside this loop>
'   Loop While zeroBrent(s, t)
'
' Alternatively, the user can code their own Function, such as 'MyFunc', and
' use an evaluation of that routine as the expression passed to 'zeroBrent':
'
'   Dim u as Double  ' will be passed by reference, so must be a variable
'   ...
'   ' set initial value of 'u' to start of search interval
'   zeroBrentInit u, 1#, 2#, 0.0001, 50&
'   Do
'   Loop While zeroBrent(u, MyFunc(a, b, u, z))
'
' Note that in this example 'MyFunc' depends on several variables, not just 'u'.
'
Const ID_c As String = M_c & "zeroBrent"

If Not initialized_m Then  ' VB Error 17 = "Can't perform requested operation"
  why_m = T9_c  ' "zeroBrentInit not called first"
  state_m = StateFirstPoint_c  ' back to square zero
  hist_m = "{Not initialized}"
  Err.Raise 17&, ID_c, _
    "Problem in routine " & ID_c & " {entry} after" & EOL & _
    "Initializations: " & inits_m & EOL & _
    "'zeroBrentInit' not called before 'zeroBrent'" & EOL & _
    "Cannot proceed. Sorry!"
  ' this is here just in case someone steps past Err.Raise
  zeroBrent = False
  Exit Function
Else
  why_m = T7_c  ' "not finished"; for premature calls to 'zeroBrentWhy'
End If

' Code begins - count initializations & evaluations so far
If state_m = StateFirstPoint_c Then  ' this is the first call of a new problem
  historyAdd "1"  ' in case of immediate zero
  evals_m = 1&    ' initialize evaluation count on first call
Else  ' we are working on an existing problem
  evals_m = evals_m + 1&
  If evals_m = 2& Then historyAdd "2"  ' in case of immediate zero
End If

If f = 0# Then  ' we have an exact zero; finish up immediately
  historyAdd "Z"
  ' make both ends of bracketing interval equal to zero's location
  xb_m = x
  fb_m = f
  xc_m = x
  fc_m = f
  ' save reason why we quit
  why_m = Replace(T1_c, "%1", CStr(x))
  initialized_m = False  ' claim that we are not initialized
  ' set to exit from caller's loop
  zeroBrent = False

Else  ' carry out actions for the state we are in now
  If forceState_m Then
    state_m = -42&  ' to debug bad-state error handling
    forceState_m = False
  End If

  Select Case state_m  ' jump to code for the present state

    Case StateFirstPoint_c
      ' get variable & expression value at one end of initial bracket interval
      xc_m = x
      fc_m = f
      ' set end of initial interval
      ' note: if the caller perversely starts with 'x1' = 'x2' & the expression
      ' is non-zero at both points, there will be an error exit after the second
      ' point is evaluated, since the 1st & 2nd signs will be the same
      xb_m = x2_m
      ' insert a huge fictitious expression value for 'zeroBrentOtherF' calls
      ' make it have the same sign as fc_m, so there is no "root"
      fb_m = Sgn(fc_m) * MaxDouble_c
      ' get caller's (or default) absolute error tolerance, force non-negative
      errAbs_m = 0.5 * errAbs_m ' Brent test (rarely) allows 2X the user's value
      ' keep a minimum error value based on the initial end-point values
      ' avoids extra work if root is near zero compared to max of 'x1' & 'x2'
      Dim roundoff As Double
      If Abs(x) >= Abs(x2_m) Then
        roundoff = 0.5 * EpsDouble_c * Abs(x)
      Else
        roundoff = 0.5 * EpsDouble_c * Abs(x2_m)
      End If
      If errAbs_m < roundoff Then
        errAbs_m = roundoff
      End If
      If evalMax_m < 3& Then  ' enforce some measure of sanity
        evalMax_m = 3&
      End If
      ' set up for next call at other end of initial interval
      x = xb_m
      state_m = StateSecondPoint_c
      ' set to repeat caller's loop
      zeroBrent = True

    Case StateSecondPoint_c
      ' get expression value at other end of initial bracket interval
      fb_m = f
      ' test for both initial-point signs same (note fb_m <> 0 & fc_m <> 0 here)
      If Sgn(fb_m) = Sgn(fc_m) Then
        state_m = StateFirstPoint_c  ' back to square zero
        initialized_m = False  ' claim that we are not initialized
        ' the caller may have used "On Error Resume Next", so set return values
        sortPoints
        ' we can't decide what to do now, so someone else must fix the problem
        ' VB error 5 is "Invalid procedure call or argument"
        Dim ts As String
        If xb_m = xc_m Then
          historyAdd "I"
          ts = "Initial-interval ends have the same location (zero length):"
          why_m = T5_c
        Else
          historyAdd "S"
          ts = "Initial expression values have same sign:"
          why_m = T4_c
        End If
        Err.Raise 5&, ID_c, _
          "Problem in routine " & ID_c & " {secondPoint} after" & EOL & _
          "Initializations: " & inits_m & _
          "  evaluations this time: " & evals_m & EOL & _
          ts & EOL & _
          "f(" & xc_m & ") = " & fc_m & EOL & _
          "f(" & xb_m & ") = " & fb_m & EOL & _
          "Cannot proceed. Sorry!"
        ' this is here just in case someone steps past Err.Raise
        zeroBrent = False
        Exit Function
      Else ' signs are different - set up for next call
        ' set values for entry to main loop - crossing is between 'b' & 'c'
        xa_m = xc_m  ' copy 'c' into 'a' - will be reversed as first action
        fa_m = fc_m
        fc_m = fb_m  ' force linear interpolation as the first action
        ' common code - checks termination conditions; False if met
        zeroBrent = findNewBracket(x)
      End If

    Case StateIterate_c  ' state inside the main loop of the algorithm
      ' save argument & expression values
      xb_m = x
      fb_m = f
      ' common code - checks termination conditions; False if met
      zeroBrent = findNewBracket(x)

    Case Else  ' we have been sent to a state that does not exist - abort
      Dim badState As Long
      badState = state_m
      state_m = StateFirstPoint_c  ' back to square zero
      hist_m = zeroBrentHistory() & "{Bad state " & badState & "}"
      ' save reason why we quit
      why_m = Replace(T6_c, "%1", CStr(badState))
      initialized_m = False  ' claim that we are not initialized
      ' the caller may have used "On Error Resume Next", so set return values
      sortPoints
      initialized_m = False  ' claim that we are not initialized
      forceStateError False  ' in case this was a forced error
      ' abandon all hope
      Err.Raise 51&, ID_c, _
        "Problem in routine " & ID_c & " {Case Else} after" & EOL & _
        "Initializations: " & inits_m & _
        "  evaluations: " & evals_m & EOL & EOL & _
        "Tried to go to non-existent state " & badState & EOL & EOL & _
        "This is a programming logic error. Cannot proceed. Sorry!"
      ' this is here just in case someone steps past Err.Raise
      zeroBrent = False
      Exit Function
  End Select
End If
End Function

'===============================================================================
Public Function zeroBrentBestF() As Double
Attribute zeroBrentBestF.VB_Description = "The bracket-end expression value of least absolute value."
' The bracket-end expression value of least absolute value.
zeroBrentBestF = fb_m
End Function

'===============================================================================
Public Function zeroBrentBestX() As Double
Attribute zeroBrentBestX.VB_Description = "The bracket-end variable value where expression is of least absolute value. Probably closest to zero crossing."
' The bracket-end variable value where expression is of least absolute value.
' This is probably the closest point to the zero crossing.
zeroBrentBestX = xb_m
End Function

'===============================================================================
Public Function zeroBrentBracketWidth() As Double
Attribute zeroBrentBracketWidth.VB_Description = "The absolute width of the zero-bracketing interval."
' The absolute width of the bracketing interval after the previous call.
zeroBrentBracketWidth = Abs(xb_m - xc_m)
End Function

'===============================================================================
Public Function zeroBrentEvals() As Long
Attribute zeroBrentEvals.VB_Description = "The number of expression evaluations that have been carried out."
' The number of expression evaluations that have been carried out.
zeroBrentEvals = evals_m
End Function

'===============================================================================
Public Function zeroBrentHistory() As String
Attribute zeroBrentHistory.VB_Description = "A coded history of the actions carried out by the algorithm. The codes are explained by 'zeroBrentHistoryCodes'."
' Reports the sequence of code actions. The code definitions are returned by
' 'zeroBrentHistoryCodes'. Note that the final code indicates the reason for
' exit from "zeroBrent".
Dim qMark As Long
qMark = InStr(hist_m, "?")  ' text may contain "?" in unused positions at end
If qMark > 0& Then
  zeroBrentHistory = Left$(hist_m, qMark - 1&)  ' if so, trim them off
Else
  zeroBrentHistory = hist_m
End If
If 0& = Len(hist_m) Then hist_m = "<empty string>"
End Function

'===============================================================================
Public Function zeroBrentHistoryCodes() As String
Attribute zeroBrentHistoryCodes.VB_Description = "A multi-line text description of the codes returned by 'zeroBrentHistory'."
' Returns a multi-line text string describing the history codes returned by
' 'zeroBrentHistory'.
zeroBrentHistoryCodes = _
"1 = first point" & EOL & _
"2 = second point" & EOL & _
"b = bisection point: interpolation not accepted" & EOL & _
"B = bisection point: small 2-ago bracket, or increasing values" & EOL & _
"I = error halt because initial interval length is zero" & EOL & _
"L = linear-interpolation point" & EOL & _
"N = return because function count limit exceeded" & EOL & _
"S = error halt because initial signs are the same" & EOL & _
"T = return because error tolerance (as adjusted) was met" & EOL & _
"Q = inverse-quadratic-interpolation point" & EOL & _
"Z = return because function exactly zero at evaluation point" & EOL & _
"- = previous action's position adjusted for minimum allowed spacing"
End Function

'===============================================================================
Public Sub zeroBrentInit( _
  ByRef x As Double, _
  ByVal x1 As Double, _
  ByVal x2 As Double, _
  Optional ByVal errAbs As Double = 0#, _
  Optional ByVal evalMax As Long = EvalMaxDefault_c)
Attribute zeroBrentInit.VB_Description = "Resets the state machine to the initial state, and sets parameters for the subsequent calls to 'zeroBrent'. Sets x = x1."
' Resets the state machine to the initial state, and sets parameters for the
' subsequent calls to 'zeroBrent'. Sets x = x1. For example:
'
' Dim x As Double
' zeroBrentInit x, 1#, 2#, 0.00000001, 20&
' Do
' Loop While zeroBrent(x, Sin(x - Sqr(2.26)))
'
' Example result: x = 1.50332963783714  Sin(x - Sqr(2.26) = -1.5077E-13
' Exact zero near x = 1.50332963783729
state_m = StateFirstPoint_c  ' start in first-point state
x = x1                       ' first evaluation at one end of search interval
x2_m = x2                    ' other end of search interval
errAbs_m = errAbs            ' absolute error tolerance (0 -> go to roundoff)
evalMax_m = evalMax          ' maximum evaluations allowed
inits_m = inits_m + 1&       ' increment total initialization count
initialized_m = True         ' claim that we are initialized
hist_m = String$(32&, "?")   ' 'empty' history
why_m = T8_c                 ' for premature calls to 'zeroBrentWhy'
End Sub

'===============================================================================
Public Function zeroBrentOtherF() As Double
Attribute zeroBrentOtherF.VB_Description = "The bracket-end expression value of greater or equal absolute value. Huge if only one evaluation so far."
' The bracket-end expression value of greater or equal absolute value. Huge if
' only one evaluation so far.
zeroBrentOtherF = fc_m
End Function

'===============================================================================
Public Function zeroBrentOtherX() As Double
Attribute zeroBrentOtherX.VB_Description = "The bracket-end variable value where expression is of greater or equal absolute value."
' The bracket-end variable value where expression is of greater or equal
' absolute value.
zeroBrentOtherX = xc_m
End Function

'===============================================================================
Public Function ZeroBrentVersion(Optional ByVal trigger As Variant) As String
Attribute ZeroBrentVersion.VB_Description = "Date of latest revision to this module as a string in the format 'YYYY-MM-DD' such as 2009-06-18. It's a Function so Excel etc. can use it."
' The date of the latest revision to this module as a string in the format
' 'YYYY-MM-DD' such as 2019-06-18. It's a function so Excel etc. can use it.
ZeroBrentVersion = Version_c
End Function

'===============================================================================
Public Function zeroBrentWhy() As String
Attribute zeroBrentWhy.VB_Description = "When done, the reason why the routine terminated, in text form. A numeric code is the first item in the string; Val(zeroBrentWhy()) yields the code. Possible code and text values are supplied by 'zeroBrentWhyTexts'."
' When done, the reason why the routine terminated in text form. A numeric code
' is the first item in the string; Val(zeroBrentWhy()) yields the code.
' The possible code text values are returned by "zeroBrentWhyTexts".
' The codes are defined in Const values at the top of this Module.

' if VB is cold-starting, 'why_m' is uninitialized, so set it
If why_m = "" Then why_m = T8_c  ' "'zeroBrent' routine not yet called"
zeroBrentWhy = why_m
End Function

'===============================================================================
Public Function zeroBrentWhyTexts() As String
Attribute zeroBrentWhyTexts.VB_Description = "A multi-line text description of the code and text returned by 'zeroBrentWhy'."
' All the error texts that may be returned by 'zeroBrentWhy', one per line.
zeroBrentWhyTexts = _
Replace(T1_c, "%1", "x.xxx") & EOL & _
Replace(T2_c, "%1", "x.xxx") & EOL & _
Replace(T3_c, "%1", "NN") & EOL & _
T4_c & EOL & _
T5_c & EOL & _
Replace(T6_c, "%1", "NN") & EOL & _
T7_c & EOL & _
T8_c & EOL & _
T9_c
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Function findNewBracket( _
  ByRef x As Double) _
As Boolean
' This common code is performed after the 2nd point, and after all later points.
' Insert new point, maintaining bracket. Sort points. Find how far to step
' from best-so-far x to new x, and carry out that step. Check exit conditions,
' and return True to continue working, or return False to terminate process.
' Update the value of 'x' in preparation for the next iteration.

' fix bracket & prevent quadratic interpolation if root not between 'b' & 'c'
' this test is also True if this is the initial pass through the "loop"
If Sgn(fb_m) = Sgn(fc_m) Then
  xc_m = xa_m  ' this forces linear interpolation
  fc_m = fa_m
  step_m = xb_m - xa_m
  stepOld_m = step_m
End If

' if 'b' is worse than 'c', swap them & set 'a' = old 'b'
' this forces a linear interpolation attempt
sortPoints

' the three points are: 'b' is best so far, 'a' is the previous value of 'b',
' and 'c' is on the other side of the zero crossing from 'b'
' Debug.Assert Sgn(fb_m) * Sgn(fc_m) = -1#

' the magnitudes obey |fb_m| <= |fc_m|
' Debug.Assert Abs(fb_m) <= Abs(fc_m)

' calculate the present (signed) bracket size
Dim dx As Double
dx = xc_m - xb_m

' get the error tolerance needed to allow for floating-point roundoff
Dim tol As Double
tol = 2# * EpsDouble_c * Abs(xb_m)
' use maximum of user request and roundoff limit
If tol < errAbs_m Then tol = errAbs_m

' set up for possible exit
x = xb_m
' test for exit conditions

' see if we have reduced the bracket size to the caller's limit or less
' if converging rapidly, 'b' is very close to the zero crossing
' thus in most cases we are within much less than 'dx' of the crossing
If Abs(dx) <= 2# * tol Then
  historyAdd "T"
    ' adjust tolerance for internal 0.5 factor
  tol = 2# * tol
  ' save reason why we quit
  Dim valStr As String
  If (tol >= 0.00000001) And (tol <= 1000000000#) Then
    valStr = CStr(tol)  ' fixed-point format
  Else
    valStr = Format$(tol, "0.0000E-0")  ' floating-point format
  End If
  why_m = Replace(T2_c, "%1", CStr(valStr))
  initialized_m = False   ' claim that we are not initialized
  findNewBracket = False  ' set to terminate the caller's loop
  Exit Function
' see if we have met or exceeded the caller's maximum evaluation count
ElseIf evals_m >= evalMax_m Then
  historyAdd "N"
  ' save reason why we quit
  why_m = Replace(T3_c, "%1", CStr(evalMax_m))
  initialized_m = False   ' claim that we are not initialized
  findNewBracket = False  ' set to terminate the caller's loop
  Exit Function
Else  ' no termination criteria were met, so continue on
  state_m = StateIterate_c  ' set to iterate the "loop"
  findNewBracket = True  ' set to redo the caller's loop
End If

' set the new value of x, in preparation for evaluation of f(x)
' these local quantities don't need to be saved between calls
Dim denom As Double, did As String, linear As Boolean, numer As Double
Dim rAC As Double, rBA As Double, rBC As Double, t1 As Double, t2 As Double
If (Abs(stepOld_m) >= tol) And (Abs(fa_m) > Abs(fb_m)) Then
  ' bracket before last above tolerance & values decreasing - interpolate
  rBA = fb_m / fa_m  ' we have -1 < rBA < 1
  linear = False
  If xa_m = xc_m Then  ' equality is special flag forcing linear interpolation
    numer = dx * rBA
    denom = 1# - rBA
    linear = True
  Else  ' get coefficients for inverse quadratic interpolation
    rAC = fa_m / fc_m
    rBC = fb_m / fc_m
    numer = rBA * (dx * rAC * (rAC - rBC) - (xb_m - xa_m) * (rBC - 1#))
    denom = (rBA - 1#) * (rAC - 1#) * (rBC - 1#)
    linear = False
  End If
  ' swap sign so numer >= 0 but numer/denom unchanged; used in tests below
  If numer < 0# Then
    numer = -numer
  Else
    denom = -denom
  End If
  ' check whether interpolation result is acceptable
  ' test 1: must step less than half as far as two steps ago
  '         failure to do so indicates poor convergence
  ' test 2: must be less than 3/4 of way from 'b' to 'c'
  '         if not, vertex of inverse parabola is too close to 'c'
  t1 = 0.5 * Abs(stepOld_m * denom)
  ' note that 'step_m' & 'denom' both switch signs if 'x' is reversed - no Abs()
  t2 = 0.75 * dx * denom - 0.5 * tol * Abs(denom)
  ' save present value of bracket size as 'previous' value
  stepOld_m = step_m
  If (numer < t1) And (numer < t2) Then  ' accepted - do interpolation
    step_m = numer / denom
    If linear Then
      historyAdd "L"
    Else
      historyAdd "Q"
    End If
  Else
    ' interpolation step not accepted - bisect
    step_m = 0.5 * dx
    stepOld_m = step_m
    historyAdd "b"
  End If
Else
  ' bracket before last below tolerance or values non-decreasing - bisect
  step_m = 0.5 * dx
  stepOld_m = step_m
  historyAdd "B"
End If
' make 'a' have previous 'b' for use in 3-point interpolation
xa_m = xb_m
fa_m = fb_m
' step to new point 'b' but don't take a "small" step
If Abs(step_m) > tol Then
  xb_m = xb_m + step_m
Else  ' would step by less than 'tol' so use that step size instead
  xb_m = xb_m + Sgn(dx) * tol
  historyAdd "-"
End If
' remember what we did here
didOld_m = did
' set up for next function call
x = xb_m
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub forceStateError(ByVal forceState As Boolean)
' This is used during debugging to force an illegal-state error (if True)
' It will be toggled back to False on first use
forceState_m = forceState
End Sub

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Function historyAdd(ByVal addOn As String)
' Add a history character without doing much concatenation.
' initialize on first call of a new problem
If addOn = "1" Then nEntries_m = 0&
' step to next slot, making more room if necessary
nEntries_m = nEntries_m + 1&
If nEntries_m > Len(hist_m) Then  ' we need more room
  hist_m = hist_m & String$(32&, "?")
End If
' insert the history character
Mid$(hist_m, nEntries_m, 1&) = Left$(addOn, 1&)
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub sortPoints()
' Put points in order, so that point of least absolute function value is in 'b'.
' The fact that this sets 'xc_m' to 'xa_m' forces linear interpolation.
If Abs(fb_m) > Abs(fc_m) Then  ' point 'c' is best; swap them
  xa_m = xb_m
  xb_m = xc_m
  xc_m = xa_m
  fa_m = fb_m
  fb_m = fc_m
  fc_m = fa_m
End If
End Sub

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
#If UnitTest Then

Public Sub zeroBrent_Unit_Tests()
Attribute zeroBrent_Unit_Tests.VB_Description = "Unit test routine. Test results go to file in the IDE/EXE/Workbook directory & to Immediate window (if in IDE)."
' Main unit-test routine for this module.

' To run the test from VB6, enter this routine's name (above) in the Immediate
' window (if the Immediate window is not open, use View.. or Ctrl-G to open it).
' To run this routine from VBA, put the cursor somewhere in it and hit F5.
' To run this routine from somewhere in a code, call it.

' The output will be in the file 'ZeroBrentMod[yyyy-mm-dd].UnitTests.txt' on
' disk, and in the immediate window if in the VB[6|A] editor (may not fit).

Dim nWarn As Long
nWarn = 0&

utFileOpen M_c & "UnitTests.txt"

utTeeOut "########## Test of " & Left$(M_c, Len(M_c) - 1&) & " routines at " & _
  Format$(Now(), "yyyy-mm-dd hh:mm:ss")
utTeeOut

utTeeOut "=== Possible 'zeroBrentWhy()' text strings:"
utTeeOut zeroBrentWhyTexts()
utTeeOut

utTeeOut "=== Single-letter codes returned by 'zeroBrentHistory()':"
utTeeOut zeroBrentHistoryCodes()
utTeeOut

Dim x As Double  ' will be passed by reference, so must be a variable
utTeeOut "=== Most tests use the initial interval from 1 to 2"
utTeeOut "=== Most tests use default errAbs = 0 & evalMax = 200 arguments"
utTeeOut

' test immediate-zero logic
Dim worst As Double
worst = 0#

' zero at first point
zeroBrentInit x, 1#, 2#
utTeeOut "Zero at first point: zeroBrentInit x, 1#, 2#"
Do
Loop While zeroBrent(x, Sin(x - 1#))
utCompareAbs "zeroBrent(x, Sin(x - 1#)) x", x, 1#, worst
utTeeOut "Why: " & zeroBrentWhy() & "  evals " & zeroBrentEvals()
utTeeOut "History: " & zeroBrentHistory()
utCompareAbs "zeroBrent(x, Sin(x - 1#)) evals", _
  zeroBrentEvals(), 1#, worst
utTeeOut

' zero at second point
zeroBrentInit x, 1#, 2#
utTeeOut "Zero at second point: zeroBrentInit x, 1#, 2#"
Do
Loop While zeroBrent(x, Sin(2# - x))
utCompareAbs "zeroBrent(x, Sin(2# - x)) x", x, 2#, worst
utTeeOut "Why: " & zeroBrentWhy() & "  evals " & zeroBrentEvals()
utTeeOut "History: " & zeroBrentHistory()
utCompareAbs "zeroBrent(x, Sin(2# - x)) evals", _
  zeroBrentEvals(), 2#, worst
utTeeOut

' zero at third point
zeroBrentInit x, 1#, 2#
utTeeOut "Zero at third point: zeroBrentInit x, 1#, 2#"
Do
Loop While zeroBrent(x, Sin(x - 1.5))
utCompareAbs "zeroBrent(x, Sin(x - 1.5)) x", x, 1.5, worst
utTeeOut "Why: " & zeroBrentWhy() & "  evals " & zeroBrentEvals()
utTeeOut "History: " & zeroBrentHistory()
utCompareAbs "zeroBrent(x, Sin(x - 1.5)) evals", _
  zeroBrentEvals(), 3#, worst
utTeeOut
utTeeOut "Immediate-zero test summary:"
utCheckLimit worst, 0#, nWarn
utTeeOut

utTeeOut "=== Test error handling"
utTeeOut

' this will cause an error, because there is no sign change
On Error Resume Next
Err.Clear
zeroBrentInit x, 1#, 2#
utTeeOut "No sign change in interval: zeroBrentInit x, 1#, 2#"
Do
Loop While zeroBrent(x, Sin(x))
utErrorCheck "zeroBrent(x, Sin(x))", 5&, nWarn
On Error GoTo 0
utTeeOut "Why: " & zeroBrentWhy() & "  evals " & zeroBrentEvals()
utTeeOut "History: " & zeroBrentHistory()
utTeeOut

' this will cause an error, because initial interval = 0 (no sign change)
On Error Resume Next
Err.Clear
zeroBrentInit x, Sqr(2#), Sqr(2#)
utTeeOut "Zero-length initial interval: zeroBrentInit x, Sqr(2#), Sqr(2#)"
Do
Loop While zeroBrent(x, Sin(x))
utErrorCheck "zeroBrent(x, Sin(x))", 5&, nWarn
On Error GoTo 0
utTeeOut "Why: " & zeroBrentWhy() & "  evals " & zeroBrentEvals()
utTeeOut "History: " & zeroBrentHistory()
utTeeOut

' this will cause an error, because we do not call 'zeroBrentInit'
On Error Resume Next
Err.Clear
utTeeOut "No zeroBrentInit call"
Do
Loop While zeroBrent(x, Sin(x))
utErrorCheck "zeroBrent(x, Sin(x))", 17&, nWarn
On Error GoTo 0
utTeeOut "Why: " & zeroBrentWhy() & "  evals " & zeroBrentEvals()
utTeeOut "History: " & zeroBrentHistory()
utTeeOut

' this will cause a illegal-state error
On Error Resume Next
Err.Clear
zeroBrentInit x, 1#, 2#
utTeeOut "Forced illegal-state error: zeroBrentInit x, 1#, 2#"
forceStateError True
Do
Loop While zeroBrent(x, Sin(x - 1.5))
utErrorCheck "zeroBrent(x, Sin(x - 1.5))", 51&, nWarn
On Error GoTo 0
utTeeOut "Why: " & zeroBrentWhy() & "  evals " & zeroBrentEvals()
utTeeOut "History: " & zeroBrentHistory()
utTeeOut

utTeeOut "=== Function with many zero crossings, and no exact zero value"
utTeeOut

' many zeros, and an expression so steep that an exact zero does not appear
zeroBrentInit x, 1#, 2#
utTeeOut "Many zeros, no exact zero: zeroBrentInit x, 1#, 2#"
Do
Loop While zeroBrent(x, Cos(1000# * x))
utTeeOut "zeroBrent(x, Cos(1000# * x)) return value is pi * " & _
  1000# * x / (4# * Atn(1#))
utTeeOut "Why: " & zeroBrentWhy() & "  evals " & zeroBrentEvals()
utTeeOut "History: " & zeroBrentHistory()
utTeeOut "Best X: " & zeroBrentBestX()
utTeeOut "Best F: " & zeroBrentBestF()
utTeeOut "Bracket: " & zeroBrentBracketWidth()
utTeeOut "Other X: " & zeroBrentOtherX()
utTeeOut "Other F: " & zeroBrentOtherF()
utTeeOut "Check function error:"
utCompareAbs "zeroBrent(x, Cos(1000# * x)) f(x)", zeroBrentBestF(), 0#, worst
utCheckLimit worst, 0.0000000000003, nWarn
utCompareLess "zeroBrent(x, Cos(1000# * x)) evals", _
  zeroBrentEvals(), EvalMaxDefault_c, nWarn

' location of all subsequent roots
Const R2_c As Double = 2.26
Dim r As Double
r = Sqr(R2_c)
utTeeOut
utTeeOut "=== All subsequent roots are located at r = Sqr(" & R2_c & ") = " & r
utTeeOut

' zero inside interval; expression increases, negative error tolerance
zeroBrentInit x, 1#, 2#, -0.001
utTeeOut "Negative error tolerance: zeroBrentInit x, 1#, 2#, -0.001"
Do
Loop While zeroBrent(x, Sin(x - r))
utCompareAbs "zeroBrent(x, Sin(x - r)) x", x, r, worst
utTeeOut "Why: " & zeroBrentWhy() & "  evals " & zeroBrentEvals()
utTeeOut "History: " & zeroBrentHistory()
utTeeOut

Const erAbs As Double = 0#
utTeeOut "=== Absolute error tolerance 'erAbs' from now on: " & _
  Format$(erAbs, "0.000E-0")
utTeeOut

' zero inside interval; expression increases
zeroBrentInit x, 1#, 2#, erAbs
utTeeOut "Simple, smooth, increasing function: zeroBrentInit x, 1#, 2#, erAbs"
Do
Loop While zeroBrent(x, Sin(x - r))
utCompareAbs "zeroBrent(x, Sin(x - r)) x", x, r, worst
utTeeOut "Why: " & zeroBrentWhy() & "  evals " & zeroBrentEvals()
utTeeOut "History: " & zeroBrentHistory()
utTeeOut

' zero inside interval; expression decreases
zeroBrentInit x, 1#, 2#, erAbs
utTeeOut "Simple, smooth, decreasing function: zeroBrentInit x, 1#, 2#, erAbs"
Do
Loop While zeroBrent(x, Sin(r - x))
utCompareAbs "zeroBrent(x, Sin(r - x)) x", x, r, worst
utTeeOut "Why: " & zeroBrentWhy() & "  evals " & zeroBrentEvals()
utTeeOut "History: " & zeroBrentHistory()
utTeeOut

' zero inside interval; expression increases, ends switched
utTeeOut "Initial interval reversed: zeroBrentInit x, 2#, 1#, erAbs"
zeroBrentInit x, 2#, 1#, erAbs
Do
Loop While zeroBrent(x, Sin(x - r))
utCompareAbs "zeroBrent(x, Sin(x - r)) x", x, r, worst
utTeeOut "Why: " & zeroBrentWhy() & "  evals " & zeroBrentEvals()
utTeeOut "History: " & zeroBrentHistory()
utTeeOut

' zero inside interval; expression increases; evaluation count limit = -999
zeroBrentInit x, 1#, 2#, erAbs, -999&
utTeeOut "Evaluation limit -999: zeroBrentInit x, 1#, 2#, erAbs, -999&"
Do
Loop While zeroBrent(x, Sin(x - r))
utCompareAbs "zeroBrent(x, Sin(x - r)) x", x, r, worst
utTeeOut "Why: " & zeroBrentWhy() & "  evals " & zeroBrentEvals()
utTeeOut "History: " & zeroBrentHistory()
utCheckLimit worst, 0.5, nWarn
utCompareLessEqual "zeroBrent(1#, Sin(x - r)) evals", _
  zeroBrentEvals(), EvalMaxDefault_c, nWarn
utTeeOut

' zero inside interval; expression increases; evaluation count restricted
zeroBrentInit x, 1#, 2#, erAbs, 32&
utTeeOut "Calls restricted, hard function: zeroBrentInit x, 1#, 2#, erAbs, 32&"
Do
Loop While zeroBrent(x, (x - r) ^ 19#)
utCompareAbs "zeroBrent(x, (x - r)^19) x", x, r, worst
utTeeOut "Why: " & zeroBrentWhy() & "  evals " & zeroBrentEvals()
utTeeOut "History: " & zeroBrentHistory()
utCheckLimit worst, 0.00021, nWarn
utCompareLessEqual "zeroBrent(x, Sin(x - r)) evals", zeroBrentEvals(), _
  32&, nWarn
utTeeOut

' zero inside interval; expression increases; hard function
zeroBrentInit x, 1#, 2#, erAbs, 150&
utTeeOut "Hard function: zeroBrentInit x, 1#, 2#, erAbs, 150&"
Do
Loop While zeroBrent(x, (x - r) ^ 19#)
utCompareAbs "zeroBrent(x, (x - r)^19) x", x, r, worst
utTeeOut "Why: " & zeroBrentWhy() & "  evals " & zeroBrentEvals()
utTeeOut "History: " & zeroBrentHistory()
utCheckLimit worst, 0.0000000007, nWarn
utCompareLessEqual "zeroBrent(x, Sin(x - r)) evals", _
  zeroBrentEvals(), 150&, nWarn
utTeeOut

' evilness parameters
Const A As Double = 0.000000000000001
Const B As Double = A * A
Const C As Double = 2# * A
Const D As Double = 1E-99

utTeeOut "Cases below all use the evil expression:" & EOL & _
  "  f(x) = " & Format$(C, "0E-0") & _
  " * (x - r) / ((x - r) * (x - r) + " & Format$(B, "0E-0") & ") + " & _
  Format$(D, "0E-0")
utTeeOut
utTeeOut "Error tolerance adjusted steadily tighter:"
utTeeOut

' increasing accuracy
Dim e As Double, ei As Double
ei = 1#

Dim j As Long
Dim ec As Double  ' comparison value for error
Dim bc As Double  ' comparison value for bracket
For j = 1& To 17&
  If j < 17& Then
    e = 1# / ei
    ec = e
    bc = e
  Else
    e = 0#
    ec = 2.23E-16
    bc = 9E-16
  End If
  zeroBrentInit x, 1#, 2#, e
  utTeeOut "zeroBrentInit x, 1#, 2#, " & Format$(e, "0E-0")
  Do
  Loop While zeroBrent(x, C * A * (x - r) / ((x - r) * (x - r) + B) + D)
  utCompareAbs "zeroBrent(x, f(x)) x", _
    x, r, worst
  utTeeOut "Why: " & zeroBrentWhy() & "  evals " & zeroBrentEvals()
  utTeeOut "History: " & zeroBrentHistory()
  utCheckLimit worst, ec, nWarn
  utCompareLessEqual _
    "zeroBrent(x, f(x)) bracket", zeroBrentBracketWidth(), bc, nWarn
  utTeeOut
  ei = ei * 10#  ' this is exact, while e = e / 10# is not
Next j

utTeeOut "=== Step-by-step test"
utTeeOut
utTeeOut "Call zeroBrent(x, " & Format$(C, "0E-0") & " * (x - r) / ((x - r)" & _
  " * (x - r) + " & Format$(B, "0E-0") & ") + " & Format$(D, "0E-0") & ")"
zeroBrentInit x, 1#, 2#
utTeeOut "Step by step test: zeroBrentInit x, 1#, 2#"
For j = 1& To 4&
  Call zeroBrent(x, C * A * (x - r) / ((x - r) * (x - r) + B) + D)
  utTeeOut "  zeroBrentEvals() = " & zeroBrentEvals()
  utTeeOut "    zeroBrentBestF() = " & zeroBrentBestF()
  utTeeOut "    zeroBrentBestX() = " & zeroBrentBestX()
  utTeeOut "    zeroBrentBracketWidth() = " & zeroBrentBracketWidth()
  utTeeOut "    zeroBrentOtherF() = " & zeroBrentOtherF()
  utTeeOut "    zeroBrentOtherX() = " & zeroBrentOtherX()
  utTeeOut "    zeroBrentWhy() = " & zeroBrentWhy()
  If j = 1& Then
    Dim r0 As Double, r1 As Double, r2 As Double, r3 As Double, r4 As Double, _
      r5 As Double, r6 As String
    r0 = zeroBrentEvals()
    r1 = zeroBrentBestF()
    r2 = zeroBrentBestX()
    r3 = zeroBrentBracketWidth()
    r4 = zeroBrentOtherF()
    r5 = zeroBrentOtherX()
    r6 = zeroBrentWhy()
  End If
Next j
utTeeOut "History: " & zeroBrentHistory()
utTeeOut

utTeeOut "=== Call 'zeroBrentInit', test start values against previous ones"
utTeeOut
zeroBrentInit x, 1#, 2#
utTeeOut "zeroBrentInit x, 1#, 2#"
Call zeroBrent(x, C * A * (x - r) / ((x - r) * (x - r) + B) + D)
utCompareAbs "  zeroBrentEvals()", zeroBrentEvals(), r0, worst
utCompareAbs "  zeroBrentBestF()", zeroBrentBestF(), r1, worst
utCompareAbs "  zeroBrentBestX()", zeroBrentBestX(), r2, worst
utCompareAbs "  zeroBrentBracketWidth()", zeroBrentBracketWidth(), r3, worst
utCompareAbs "  zeroBrentOtherF()", zeroBrentOtherF(), r4, worst
utCompareAbs "  zeroBrentOtherX()", zeroBrentOtherX(), r5, worst
utCheckLimit worst, 0#, nWarn
utCompareEqualString "  zeroBrentWhy() = previous start value", _
  zeroBrentWhy(), r6, nWarn
utTeeOut "History: " & zeroBrentHistory()
utTeeOut

If nWarn = 0& Then
  utTeeOut "Unit test success - no warnings were raised."
Else
  utTeeOut "Unit test FAILURE! - warning count: " & nWarn
End If
utTeeOut

' demo of how to trace function arguments and results; evaluate f(x) in loop
Dim f As Double
zeroBrentInit x, 1#, 2#
utTeeOut "Demo of x & f(x) trace printing in loop: zeroBrentInit x, 1#, 2#"
utTeeOut "Function is: f = Sin(x - r)"
Do
  f = Sin(x - r)
  utTeeOut Left$("x = " & x & Space$(21), 22) & "  f(x) = " & f
Loop While zeroBrent(x, f)
utTeeOut "Why: " & zeroBrentWhy() & "  evals " & zeroBrentEvals()
utTeeOut "History: " & zeroBrentHistory()
utTeeOut "Best X: " & zeroBrentBestX()
utTeeOut "Best F: " & zeroBrentBestF()
utTeeOut "Bracket: " & zeroBrentBracketWidth()
utTeeOut "Other X: " & zeroBrentOtherX()
utTeeOut "Other F: " & zeroBrentOtherF()
utTeeOut

utTeeOut "--- Unit Tests Completed ---"

utFileClose

' force an unhandled error that will put up a notification box
zeroBrentInit x, 1# + Sqr(0.02), 2# - Sqr(0.02)
Do
Loop While zeroBrent(x, Sin(x))

End Sub

#End If  ' UnitTest

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


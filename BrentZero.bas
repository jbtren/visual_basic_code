Attribute VB_Name = "BrentZeroMod"
Attribute VB_Description = "Localizes a zero of a one-argument function using Brent's method. Revised and coded by John Trenholme."
'
'###############################################################################
'#
'# Visual Basic 6 & VBA code module 'BrentZero.bas'
'#
'# Zero-crossing finder using Brent's method. Uses function 'brentZeroFunction'.
'#
'# Exports the routines:
'#   Function brentZero
'#   Function brentZeroBestF
'#   Function brentZeroBestX
'#   Function brentZeroBracketWidth
'#   Function brentZeroEvals
'#   Function brentZeroFunction
'#   Function brentZeroFunctionGet
'#   Sub brentZeroFunctionSet
'#   Function brentZeroHistory
'#   Function brentZeroHistoryCodes
'#   Function brentZeroOtherF
'#   Function brentZeroOtherX
'#   Function brentZeroVersion
'#   Function brentZeroWhy
'#   Function brentZeroWhyTexts
'#
'# Requires the module "UnitTestSupport.bas" if unit test code enabled
'#
'# Devised and coded by John Trenholme -  begun 2006-09-06
'#
'###############################################################################

Option Base 0
Option Compare Binary
Option Explicit
Option Private Module  ' Don't allow visibility outside this Project (if VBA)

Private Const Version_c As String = "2008-07-03"
Private Const Mod_c As String = "BrentZero"  ' module name

#Const UnitTest = True  ' set True to enable unit test code
' #Const UnitTest = False  ' set False to eliminate unit test code

' Text strings describing state of problem; they go into 'why_m'
Private Const T1_c As String = _
  "1 Found acceptably small function value at %1"
Private Const T2_c As String = _
  "2 Argument bracket dropped below adjusted size of %1"
Private Const T3_c As String = _
  "3 WARNING - evaluation count reached limit of %1"
Private Const T4_c As String = _
  "4 ERROR in call - function values at initial points have the same sign"
Private Const T5_c As String = _
  "5 ERROR in function specification - function %1 not defined"

Private Const EOL As String = vbNewLine  ' short form; works on both PC and Mac

' The smallest number that can cause a change when added to 1#
Private Const EpsDouble_c As Double = 2.22044604925031E-16  ' 2.0^(-52)

' Default value for maximum number of function evaluations
' This value will do almost any problem
Private Const EvalMaxDefault_c As Long = 200&

' Module-global variables (retained between calls; initialize as 0)
Private evals_m As Long     ' number of function evaluations so far
Private fa_m As Double      ' function at third point (often previous fb_m)
Private fb_m As Double      ' function at 'best' end of bracketing interval
Private fc_m As Double      ' function at 'worst' end of  bracketing interval
Private function_m As Long  ' index of the function to be used
Private hist_m As String    ' eval. sequence: B=bisect L=linear Q=quadratic
Private why_m As String     ' reason for exit from routine
Private xa_m As Double      ' variable at third point (often previous xb_m)
Private xb_m As Double      ' variable at 'best' end of bracketing interval
Private xc_m As Double      ' variable at 'worst' end of bracketing interval

'===============================================================================
Public Function brentZero( _
  ByVal x1 As Double, _
  ByVal x2 As Double, _
  Optional ByVal xErrAbs As Double = 0#, _
  Optional ByVal xErrRel As Double = 0#, _
  Optional ByVal fErrAbs As Double = 0#, _
  Optional ByVal evalMax As Long = EvalMaxDefault_c) _
As Double
Attribute brentZero.VB_Description = "Returns a value of 'x' at or near a zero crossing of a function that depends on the single variable 'x'. User must code his function(s) in ""brentZeroFunction"". See comments in code for usage."
' Allows the caller to localize a value of the real scalar variable 'x' that
' gives a zero crossing of the real scalar function 'brentZeroFunction(x)'.
'
' Looks for a zero crossing between initial variable values of 'x1' and 'x2'.
' Systematically shortens the bracketing interval in which a zero crossing must
' lie, keeping end signs different. Exits when absolute value of bracketing
' interval drops to or below caller's value 'xErrAbs', or when relative size
' (bracket / X) drops to or below caller's value 'xErrRel', or when absolute
' function value drops to or below caller's value 'fErrAbs', or when caller's
' evaluation-count limit 'evalMax' is exceeded.
'
' Most of the time, the output bracket will be around half the requested size
' for "good" functions, and between 0.5 and 1.0 times the requested size for
' "evil" functions. The routine may exit with a bracket larger than the
' requested size when the bracketing interval becomes so small that there are
' only a few available floating-point values between the end values.
'
' The first exit test to be met will cause exit, so any test can be made
' inoperative by setting the test value to zero (or to a very large number, in
' the case of the evaluation-count limit). Calling this routine with only two
' arguments, so that the remaining arguments take on their default values,
' will almost always return a very accurate result near the roundoff limit.
'
' Different functions can be zeroed by writing them in a Select Case statement
' inside brentZeroFunction(x), and then calling 'brentZeroFunctionSet()'
' before calling 'brentZero', to specify which one will be used. If you are
' only using one function, put it in case 0. This is the default, so you don't
' need to call 'brentZeroFunctionSet(0&)'.
'
' A zero crossing is considered to exist if f(x) = 0 and values on both sides
' have different signs, or if f(x) has different signs at values separated by
' the minimum difference allowed by finite-precision floating-point arithmetic.
' Note that a discontinuous function that has different signs on opposite sides
' of the point of discontinuity will be reported to have a zero crossing there.
' Thus, poles will be classified as zero crossings the same as roots.
'
' The points 'x1' and 'x2' do not need to be in increasing numerical order on
' input. They must, of course, be distinct. The function values must have
' different signs at these points, or the value at one (or both) must be below
' the function value tolerance 'fErrAbs'. Either situation will cause an
' immediate exit.
'
' If multiple zero crossings exist in the initial interval, the routine picks
' arbitrarily from among points where zero is crossed in the same direction as
' a straight line between the initial points. If there is a continuous region
' of zero values, some arbitrary point in the region will be picked.
'
' Returns the value of 'x' at the zero-crossing-bracket end point where the
' absolute value of the function was least. In most cases, this will be the
' point closest to the zero crossing. For reasonably smooth functions, the
' returned location will usually be much closer to the zero crossing than the
' other end of the bracketing interval.
'
' Once this routine has run, other values of interest can be retrieved by calls
' to the routines:
'
'   brentZeroBestF         The function value at the returned 'x' value
'   brentZeroBestX         The returned 'x' value (in case you want it again)
'   brentZeroBracketWidth  The width of the interval known to contain a zero
'   brentZeroEvals         The number of function evaluations used
'   brentZeroFunctionGet   The index number of the function being used
'   brentZeroHistory       Record of actions - see 'brentZeroHistoryCodes'
'   brentZeroOtherF        Function value at 'brentZeroOtherX'
'   brentZeroOtherX        Value of 'x' at other end of bracket from best 'x'
'   brentZeroWhy           Reason for exit from routine, as text
'
' Raises error 5 ("Invalid procedure call or argument") if both initial points
' have the same sign and are greater than 'fErrAbs'.
'
' Raises error 17 ("Can't perform requested operation") if the requested
' function is undefined. This is a user coding or function specification error.
'
' For more information about Brent's method, see:
'
' "Algorithms for Minimization without Derivatives" by Richard P. Brent
' Prentice-Hall (1973), Chapter 4  ISBN: 0-13-022335-2
'
' "Numerical Recipes in X" (where X is some computer language) by W. Press et.
' al., Cambridge University Press, section 9.3
'
'-------------------------------------------------------------------------------

Const R_c As String = "brentZero"  ' name of this routine
Const ID_c As String = Mod_c & "|" & R_c

On Error GoTo ErrorPlus  ' set to report problems more clearly

Static calls_s As Double  ' number of times this routine has been called
calls_s = calls_s + 1#    ' stops adding at 9,007,199,254,740,992 calls

evals_m = 0&  ' initialize evaluation-call counter

' get caller's (or default) argument and function error tolerances
Dim xErAbs As Double
xErAbs = xErrAbs
If xErAbs < 0# Then
  xErAbs = 0#  ' prevent infinite loop
End If
Dim xErRel As Double
xErRel = xErrRel
If xErRel < 1.5 * EpsDouble_c Then
  xErRel = 1.5 * EpsDouble_c ' keep at or above roundoff limit
End If
Dim fErAbs As Double
fErAbs = fErrAbs
If fErAbs < 0# Then
  fErAbs = 0#  ' prevent infinite loop
End If

' get caller's (or default) maximum evaluation count; fix silly values
If evalMax < 2& Then
  evalMax = EvalMaxDefault_c
End If

' get variable & function value at one end of initial bracket interval
historyAdd "1"
xc_m = x1
fc_m = brentZeroFunction(xc_m)
evals_m = evals_m + 1&

If Abs(fc_m) <= fErAbs Then  ' we have an acceptably small value; quit
  historyAdd "V"
  ' make both ends of bracketing interval equal to zero's location
  xb_m = xc_m
  fb_m = fc_m
  why_m = Replace(T1_c, "%1", CStr(xc_m))  ' save reason why we quit
  ' exit from routine
  GoTo Done
End If

' get variable & function value at other end of initial bracket interval
historyAdd "2"
xb_m = x2
fb_m = brentZeroFunction(xb_m)
evals_m = evals_m + 1&

' test for both initial-point signs same (note fc_m <> 0 here)
If Sgn(fb_m) = Sgn(fc_m) Then
  historyAdd "S"
  why_m = T4_c  ' save reason why we quit
  ' the caller may have used "On Error Resume Next", so set return values
  sortPoints
  ' initial points the same?
  Dim s As String
  If xb_m = xc_m Then
    s = "The initial points are the same"
  Else
    s = "May be no zero crossing between them"
  End If
  ' we can't decide what to do now, so someone else must fix the problem
  ' VB error 5 is "Invalid procedure call or argument"
  Err.Raise 5&, ID_c, _
    "Initial function values have same sign:" & EOL & _
    "f(" & xb_m & ") = " & fb_m & EOL & _
    "f(" & xc_m & ") = " & fc_m & EOL & _
    s
End If

' put 'c' into 'a'; will be reversed as the first action
xa_m = xc_m
fa_m = fc_m
' this will force a linear interpolation attempt as the first action
fc_m = fb_m

' ---------- main iteration loop ----------
Dim denom As Double    ' denominator in interpolation step
Dim dx As Double       ' size of present bracket
Dim dx1 As Double      ' value of most recent step
Dim dx2 As Double      ' value of previous step
Dim linear As Boolean  ' True if we did linear interpolation
Dim numer As Double    ' numerator in interpolation step
Dim rBA As Double      ' fb_m / fa_m
Dim rAC As Double      ' fa_m / fc_m
Dim rBC As Double      ' fb_m / fc_m
Dim t1 As Double       ' test value 1
Dim t2 As Double       ' test value 2
Dim tol As Double      ' termination tolerance on bracket size
Do
  ' see if latest function value was small enough
  If Abs(fb_m) <= fErAbs Then  ' we have an acceptably small value; quit
    historyAdd "V"
    ' make both ends of bracketing interval equal to zero's location
    xc_m = xb_m
    fc_m = fb_m
    ' save reason why we quit
    why_m = Replace(T1_c, "%1", CStr(xb_m))
    ' exit from routine
    GoTo Done
  End If
  ' fix bracket & prevent quadratic interpolation if root not between 'b' & 'c'
  ' except that when Abs(fb_m) = Abs(fc_m), bisection will be forced
  ' this test is also True if this is the initial pass through the loop
  If Sgn(fb_m) = Sgn(fc_m) Then
    xc_m = xa_m  ' this forces linear interpolation
    fc_m = fa_m
    dx1 = xb_m - xc_m
    dx2 = dx1
  End If

  ' if 'b' is worse than 'c', swap them & set 'a' = old 'b'
  ' this forces a linear interpolation attempt
  sortPoints
  
  ' at this point, the three points are: 'b' is best so far, 'a' is the
  ' previous value of 'b', and 'c' is on the other side of the zero crossing
  ' from 'b' - the magnitudes obey |fb_m| <= |fc_m|

  ' too many evaluations?
  If evals_m >= evalMax Then
    historyAdd "N"
    ' save reason why we quit
    why_m = Replace(T3_c, "%1", CStr(evalMax))
    ' exit from routine
    GoTo Done
  End If

  ' get present (signed) bracket size
  dx = xc_m - xb_m

  ' get the relative X error tolerance
  tol = xErRel * 0.5 * (Abs(xb_m) + Abs(xc_m))
  ' use maximum of relative X error and absolute X error
  If tol < xErAbs Then tol = xErAbs

  ' if converging rapidly, 'b' is very close to the zero crossing
  ' thus in most cases we are within much less than 'dx' of the crossing
  If Abs(dx) <= tol Then
    historyAdd "X"
    ' save reason why we quit
    Dim valStr As String
    If (tol >= 0.00000001) And (tol <= 1000000000#) Then
      valStr = CStr(tol)  ' fixed-point format
    Else
      valStr = Format$(tol, "0.0000E-0")  ' floating-point format
    End If
    why_m = Replace(T2_c, "%1", valStr)
    ' exit from routine
    GoTo Done
  End If

  ' set the new value of x, in preparation for evaluation of f(x)
  If (Abs(dx2) >= tol) And (Abs(fa_m) > Abs(fb_m)) Then
    ' bracket before last above tolerance & values decreasing - interpolate
    rBA = fb_m / fa_m  ' we have -1 < rBA < 1
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
    t1 = 0.5 * Abs(dx2 * denom)
    ' note that 'dx' & 'denom' both switch signs if 'x' is reversed - no Abs()
    t2 = 0.75 * dx * denom - 0.5 * tol * Abs(denom)
    ' save present value of bracket size as 'previous' value
    dx2 = dx1
    If (numer < t1) And (numer < t2) Then  ' accepted - do interpolation
      dx1 = numer / denom
      If linear Then
        historyAdd "L"
      Else
        historyAdd "Q"
      End If
    Else
      ' interpolation step not accepted - bisect
      dx1 = 0.5 * dx
      dx2 = dx1
      historyAdd "b"
    End If
  Else
    ' bracket before last below tolerance or values non-decreasing - bisect
    dx1 = 0.5 * dx
    dx2 = dx1
    historyAdd "B"
  End If

  ' make 'a' have previous 'b' for use in 3-point interpolation
  xa_m = xb_m
  fa_m = fb_m
  ' step to new point 'b' but don't take a "small" step
  If Abs(dx1) >= 0.5 * tol Then
    xb_m = xb_m + dx1
  Else  ' would step by less than tol/2 so use that step size instead
    xb_m = xb_m + Sgn(dx) * 0.5 * tol
    historyAdd "-"
  End If
  ' evaluate the function at the new point
  fb_m = brentZeroFunction(xb_m)
  evals_m = evals_m + 1&
Loop

Done:  ' single exit point, for breakpointing
brentZero = xb_m
Exit Function

ErrorPlus:  ' re-raise any error, but with added location & call info for user

Dim errNum As Long, errDsc As String, errSrc As String  ' save Err properties
errNum = Err.Number
errDsc = Err.Description
errSrc = Err.Source
On Error GoTo 0  ' avoid recursion; clears the Err object
' add location & call information (gives call stack if caller does this too)
errDsc = errDsc & EOL & _
  "routine " & ID_c & " call " & calls_s & " evals " & evals_m
Err.Raise errNum, errSrc, errDsc
End Function

'===============================================================================
Public Function brentZeroBestF() As Double
Attribute brentZeroBestF.VB_Description = "The bracket-end function value of least absolute value."
' The bracket-end function value of least absolute value.
brentZeroBestF = fb_m
End Function

'===============================================================================
Public Function brentZeroBestX() As Double
Attribute brentZeroBestX.VB_Description = "The bracket-end variable value where the function is of least absolute value. Same as return value from ""brentZero""."
' The bracket-end variable value where function is of least absolute value.
' This is the same as the return value of 'brentZero'.
brentZeroBestX = xb_m
End Function

'===============================================================================
Public Function brentZeroBracketWidth() As Double
Attribute brentZeroBracketWidth.VB_Description = "The absolute width of the bracketing interval that encloses the zero crossing."
' The absolute width of the bracketing interval after the previous call.
brentZeroBracketWidth = Abs(xb_m - xc_m)
End Function

'===============================================================================
Public Function brentZeroEvals() As Long
Attribute brentZeroEvals.VB_Description = "The number of function evaluations that were used to localize the zero crossing."
' The number of function evaluations that have been carried out.
brentZeroEvals = evals_m
End Function

'===============================================================================
Public Function brentZeroFunction( _
  ByVal x As Double, _
  Optional ByVal calledByBrent As Boolean = False) _
As Double
' The function whose zero crossing is to be found. If several functions must
' have their zero crossings found, select them with 'brentZeroFunctionSet()'.
Const R_c As String = "brentZeroFunction"  ' name of this routine
Const ID_c As String = Mod_c & "|" & R_c

On Error GoTo ErrorPlus  ' set to report problems more clearly

Static calls_s As Double  ' number of times this routine has been called
calls_s = calls_s + 1#    ' stops adding at 9,007,199,254,740,992 calls

#If UnitTest Then
Dim r As Double
r = Sqr(2.26)
#End If
Dim f As Double
Select Case function_m
  '
  '---------------------------- user-coded function(s) here --------------------
  '
  ' Case 0 is the default function. If you need only one function, put it here.
  Case 0&:
    f = Cos(x)
  '
  '---------------------------- end user-coded function(s) ---------------------
  '
#If UnitTest Then
  '---------- functions below here are used in unit tests ----------
  Case -1000&:
    f = Sin(x - 1#) + 1E-300
  Case -1001&:
    f = Sin(x - 2) + 1E-300
  Case -1002&:
    f = Sin(x - 1.5) + 1E-300
  Case -1003&:
    f = Sin(x) + 1E-300
  Case -1004&:
    f = Cos(1000# * x) + 1E-300
  Case -1005&:
    f = Sin(x - r) + 1E-300
  Case -1006&:
    f = Sin(r - x) + 1E-300
  Case -1007&:
    f = (x - r) ^ 19 + 1E-300
  Case -1008&:
    Const A As Double = 0.000000000000001
    Const B As Double = A * A
    Const C As Double = 2# * A
    Const D As Double = 1E-99
    f = C * A * (x - r) / ((x - r) * (x - r) + B) + D + 1E-300
#End If
  Case Else:  ' there is no such function
    If calledByBrent Then
      ' save reason why we quit in Brent routine variables
      historyAdd "F"
      why_m = Replace(T5_c, "%1", CStr(function_m))
    End If
    ' VB error 17 is "Can't perform requested operation"
    Err.Raise 17&, ID_c, _
      "Function number " & function_m & " requested, but isn't defined."
End Select
brentZeroFunction = f
Exit Function

ErrorPlus:  ' re-raise any error, but with added location & call info for user

Dim errNum As Long, errDsc As String, errSrc As String  ' save Err properties
errNum = Err.Number
errDsc = Err.Description
errSrc = Err.Source
On Error GoTo 0  ' avoid recursion; clears the Err object
' add location & call information (gives call stack if caller does this too)
errDsc = errDsc & EOL & _
  "routine " & ID_c & " call " & calls_s
Err.Raise errNum, errSrc, errDsc
End Function

'===============================================================================
Public Function brentZeroFunctionGet() As Long
Attribute brentZeroFunctionGet.VB_Description = "The index of the function that was set using ""brentZeroFunctionSelect."" Defaults to zero."
' The index of the function being used (set by 'brentZeroFunctionSet').
brentZeroFunctionGet = function_m
End Function

'===============================================================================
Public Sub brentZeroFunctionSet(ByVal functionIndex As Long)
Attribute brentZeroFunctionSet.VB_Description = "Select the user-coded function to be used, if there are more than one. If not called, function zero is used."
' Call this before calling 'brentZero' to set the function index. If not called,
' function 0 will be used by default.
function_m = functionIndex
End Sub

'===============================================================================
Public Function brentZeroHistory() As String
Attribute brentZeroHistory.VB_Description = "A coded history of the actions carried out by the algorithm. The codes are explained by ""brentZeroHistoryCodes""."
' Reports the sequence of code actions. The code definitions are returned by
' 'brentZeroHistoryCodes' Note that the final code indicates the reason for
' exit from 'brentZero'.
Dim qMark As Long
qMark = InStr(hist_m, "?")  ' text may contain "?" in unused positions at end
If qMark > 0& Then
  brentZeroHistory = Left$(hist_m, qMark - 1&)  ' if so, trim them off
Else
  brentZeroHistory = hist_m
End If
End Function

'===============================================================================
Public Function brentZeroHistoryCodes() As String
Attribute brentZeroHistoryCodes.VB_Description = "A multi-line text description of the codes returned by ""brentZeroHistory""."
' Returns a multi-line text string describing the history codes that are
' returned by 'brentZeroHistory'.
brentZeroHistoryCodes = _
"1 = first point" & EOL & _
"2 = second point" & EOL & _
"b = bisection point: interpolated point out of bounds" & EOL & _
"B = bisection point: 2-ago bracket small or values increasing" & EOL & _
"F = tried to evaluate an undefined function" & EOL & _
"L = linear-interpolation point" & EOL & _
"N = return because function count limit exceeded" & EOL & _
"Q = inverse-quadratic-interpolation point" & EOL & _
"S = error halt because initial signs are the same" & EOL & _
"X = return because X error tolerance (as adjusted) was met" & EOL & _
"V = return because function value error tolerance was met" & EOL & _
"- = previous action's position adjusted for minimum allowed spacing"
End Function

'===============================================================================
Public Function brentZeroOtherF() As Double
Attribute brentZeroOtherF.VB_Description = "The bracket-end function value of greater or equal absolute value."
' The bracket-end function value of greater or equal absolute value.
brentZeroOtherF = fc_m
End Function

'===============================================================================
Public Function brentZeroOtherX() As Double
Attribute brentZeroOtherX.VB_Description = "The bracket-end variable value where the function value is of greater or equal absolute value."
' The bracket-end variable value where function is of greater or equal
' absolute value.
brentZeroOtherX = xc_m
End Function

'===============================================================================
Public Function brentZeroVersion() As String
Attribute brentZeroVersion.VB_Description = "The date of the latest revision to this module as a string in the format 'YYYY-MM-DD' such as 2009-06-18. It's a Function so Excel etc. can use it."
' The date of the latest revision to this module as a string in the format
' 'YYYY-MM-DD' such as 2009-06-18. It's a function so Excel etc. can use it.
brentZeroVersion = Version_c
End Function

'===============================================================================
Public Function brentZeroWhy() As String
Attribute brentZeroWhy.VB_Description = "The reason why the routine terminated, in text form. A numeric code is the first item in the string; Val(zeroBrentWhy()) yields the code. The possible values are supplied by ""brentZeroWhyTexts""."
' When done, the reason why the routine terminated in text form. A numeric code
' is the first item in the string; Val(brentZeroWhy()) yields the code.
' The possible code text values are returned by 'brentZeroWhyTexts'.
' The codes are defined in Const values at the top of this Module.
brentZeroWhy = why_m
End Function

'===============================================================================
Public Function brentZeroWhyTexts() As String
Attribute brentZeroWhyTexts.VB_Description = "A multi-line text description of the text returned by ""brentZeroWhy""."
' All the error texts that may be returned by 'brentZeroWhy', one per line.
brentZeroWhyTexts = _
Replace(T1_c, "%1", "x.xxx") & EOL & _
Replace(T2_c, "%1", "x.xxx") & EOL & _
Replace(T3_c, "%1", "NN") & EOL & _
T4_c & EOL & _
Replace(T5_c, "%1", "NN")
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Function historyAdd(ByVal addOn As String)
Attribute historyAdd.VB_Description = "Internal routine to add an action code to the history text string."
' Add a history character without doing much expensive concatenation.
Static nEntries As Long
If addOn = "1" Then  ' this is the first call; initialize the string
  hist_m = String$(32&, "?")
  nEntries = 0&
End If
nEntries = nEntries + 1&
If nEntries > Len(hist_m) Then  ' we need more room
  hist_m = hist_m & String$(32&, "?")
End If
Mid$(hist_m, nEntries, 1&) = Left$(addOn, 1&)
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub sortPoints()
Attribute sortPoints.VB_Description = "Internal routine to put points in order."
' The algorithm requires that |fb_m| <= |fc_m|, but that may not be the case at
' this point. If not, put the points in order, so that point of least absolute
' function value is in 'b'. As a side effect, set 'xc_m' to 'xa_m' and force
' linear interpolation.
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

Public Sub Test_BrentZero()
Attribute Test_BrentZero.VB_Description = "Unit test routine. Test results go to file in the IDE/EXE/Workbook directory & to Immediate window (if in IDE)."
' Main unit-test routine for this module.

' To run the test from VB6, enter this routine's name (above) in the Immediate
' window (if the Immediate window is not open, use View.. or Ctrl-G to open it).
' To run this routine from VBA, put the cursor somewhere in it and hit F5.
' To run this routine from somewhere in a code, call it.

' The output will be in the file 'Test_BrentZero.txt' on disk, and in the
' immediate window if in the VB[6|A] editor.

Dim nWarn As Long
nWarn = 0&
Dim worst As Double
worst = 0#

utFileOpen "Test_" & Mod_c & ".txt"

utTeeOut "########## Test of " & Mod_c & " routines at " & Now()
utTeeOut Mod_c & " code version: " & Version_c
utTeeOut
utTeeOut "Single-letter history codes:"
utTeeOut brentZeroHistoryCodes()
utTeeOut
utTeeOut "Possible 'Why' text strings:"
utTeeOut brentZeroWhyTexts()
utTeeOut

Dim x As Double  ' holds return values

' small at first point
brentZeroFunctionSet -1000&
x = brentZero(1#, 2#, 0#, 0#, 1E-299)
utCompareAbs "brentZero(1#, 2#, 0#, 0#, 1E-299) Sin(x-1) x", x, 1#, worst
utTeeOut "Why: " & brentZeroWhy() & "  evals " & brentZeroEvals()
utTeeOut "History: " & brentZeroHistory()
utCompareAbs "brentZero(1#, 2#, 0#, 0#, 1E-299) evals", brentZeroEvals(), _
  1#, worst
utTeeOut

' small at second point
brentZeroFunctionSet -1001&
x = brentZero(1#, 2#, 0#, 0#, 1E-299)
utCompareAbs "brentZero(1#, 2#, 0#, 0#, 1E-299) Sin(x-2) x", x, 2#, worst
utTeeOut "Why: " & brentZeroWhy() & "  evals " & brentZeroEvals()
utTeeOut "History: " & brentZeroHistory()
utCompareAbs "brentZero(1#, 2#, 0#, 0#, 1E-299) evals", brentZeroEvals(), _
  2#, worst
utTeeOut

' small at third point
brentZeroFunctionSet -1002&
x = brentZero(1#, 2#, 0#, 0#, 1E-299)
utCompareAbs "brentZero(1#, 2#, 0#, 0#, 1E-299) Sin(x-1.5) x", x, 1.5, worst
utTeeOut "Why: " & brentZeroWhy() & "  evals " & brentZeroEvals()
utTeeOut "History: " & brentZeroHistory()
utCompareAbs "brentZero(1#, 2#, 0#, 0#, 1E-299) evals", brentZeroEvals(), _
  3#, worst
utTeeOut

' this will cause an error, because there is no such function
On Error Resume Next
Err.Clear
brentZeroFunctionSet 9999&
x = brentZero(1#, 2#)
utErrorCheck "brentZero(1#, 2#) F=9999", 17&, nWarn
On Error GoTo 0
utTeeOut "Why: " & brentZeroWhy() & "  evals " & brentZeroEvals()
utTeeOut "History: " & brentZeroHistory()
utTeeOut

' this will cause an error, because there is no sign change
On Error Resume Next
Err.Clear
brentZeroFunctionSet -1003&
x = brentZero(1#, 2#)
utErrorCheck "brentZero(1#, 2#) Sin(x)", 5&, nWarn
On Error GoTo 0
utTeeOut "Why: " & brentZeroWhy() & "  evals " & brentZeroEvals()
utTeeOut "History: " & brentZeroHistory()
utTeeOut

' this will cause an error, because initial interval = 0 (no sign change)
On Error Resume Next
Err.Clear
x = brentZero(Sqr(2.24), Sqr(2.24))
utErrorCheck "brentZero(Sqr(2.24), Sqr(2.24)) Sin(x)", 5&, nWarn
On Error GoTo 0
utTeeOut "Why: " & brentZeroWhy() & "  evals " & brentZeroEvals()
utTeeOut "History: " & brentZeroHistory()
utTeeOut

' many zeros, and a function so steep that an exact zero does not appear
brentZeroFunctionSet -1004&
x = brentZero(1#, 2#)
utTeeOut "brentZero(1#, 2#) Cos(1000*x) return value is pi * " & _
  x / (4# * Atn(1#))
utTeeOut "return value = " & x
utTeeOut "Why: " & brentZeroWhy() & "  evals " & brentZeroEvals()
utTeeOut "History: " & brentZeroHistory()
utTeeOut "Best F = " & brentZeroBestF()
utTeeOut "Best X = " & brentZeroBestX()
utTeeOut "Bracket width = " & brentZeroBracketWidth()
utTeeOut "Other F = " & brentZeroOtherF()
utTeeOut "Other X = " & brentZeroOtherX()
utTeeOut "Check function error:"
utCompareAbs "brentZero(1#, 2#) Cos(1000*x) f(x)", brentZeroBestF(), 0#, worst
utCheckLimit worst, 0.00000000000006, nWarn
utCompareLessEqual "brentZero(1#, 2#) Cos(1000*x) evals", brentZeroEvals(), _
  17&, nWarn
utTeeOut

' location of all subsequent roots
Const R2_c As Double = 2.26  ' define only once
Dim r As Double
r = Sqr(R2_c)
utTeeOut "==== All subsequent roots are located at r = Sqr(" & R2_c & ") = " & r
utTeeOut

' zero inside interval; function increases, negative absolute error tolerance
brentZeroFunctionSet -1005&
x = brentZero(1#, 2#, -0.001)
utCompareAbs "brentZero(1#, 2#, -0.001) Sin(x-r) x", x, r, worst
utTeeOut "Why: " & brentZeroWhy() & "  evals " & brentZeroEvals()
utTeeOut "History: " & brentZeroHistory()
utCompareLessEqual "brentZero(1#, 2#, -0.001) Sin(x-r) evals", _
  brentZeroEvals(), 7&, nWarn
utTeeOut

Const erAbs As Double = 0#
utTeeOut "==== Absolute error tolerance is now erAbs = " & _
  Format$(erAbs, "0.000E-0")
utTeeOut

' zero inside interval; function increases
x = brentZero(1#, 2#, erAbs)
utCompareAbs "brentZero(1#, 2#, erAbs) Sin(x-r) x", x, r, worst
utTeeOut "Why: " & brentZeroWhy() & "  evals " & brentZeroEvals()
utTeeOut "History: " & brentZeroHistory()
utCompareLessEqual "brentZero(1#, 2#, erAbs) Sin(x-r) evals", _
  brentZeroEvals(), 7&, nWarn
utTeeOut

' zero inside interval; function decreases
brentZeroFunctionSet -1006&
x = brentZero(1#, 2#, erAbs)
utCompareAbs "brentZero(1#, 2#, erAbs) Sin(r-x) x", x, r, worst
utTeeOut "Why: " & brentZeroWhy() & "  evals " & brentZeroEvals()
utTeeOut "History: " & brentZeroHistory()
utCompareLessEqual "brentZero(1#, 2#, erAbs) Sin(r-x) evals", _
  brentZeroEvals(), 7&, nWarn
utTeeOut

' zero inside interval; function increases, ends switched
brentZeroFunctionSet -1005&
x = brentZero(2#, 1#, erAbs)
utCompareAbs "brentZero(2#, 1#, erAbs) Sin(x-r) x", x, r, worst
utTeeOut "Why: " & brentZeroWhy() & "  evals " & brentZeroEvals()
utTeeOut "History: " & brentZeroHistory()
utCompareLessEqual "brentZero(2#, 1#, erAbs) Sin(x-r) evals", _
  brentZeroEvals(), 7&, nWarn
utTeeOut

' very nonlinear near root, min near underflow; evaluation count restricted
Const MaxEval_c As Long = 32&
brentZeroFunctionSet -1007&
x = brentZero(1#, 2#, erAbs, 0#, 0#, MaxEval_c)
utCompareAbs "brentZero(1#, 2#, erAbs, 0, 0, " & MaxEval_c & ") (x-r)^19 x", _
  x, r, worst
utTeeOut "Why: " & brentZeroWhy() & "  evals " & brentZeroEvals()
utTeeOut "History: " & brentZeroHistory()
utCheckLimit worst, 0.00021, nWarn
utCompareLessEqual "brentZero(1#, 2#, erAbs, 0#, 0#, " & MaxEval_c & _
  ") (x-r)^19 evals", brentZeroEvals(), MaxEval_c, nWarn
utTeeOut

' very nonlinear near root, min near underflow; evaluation count limit = -999
x = brentZero(1#, 2#, erAbs, 0#, 0#, -999&)
utCompareAbs "brentZero(1#, 2#, erAbs, 0, 0, -999) (x-r)^19 x", x, r, worst
utTeeOut "Why: " & brentZeroWhy() & "  evals " & brentZeroEvals()
utTeeOut "History: " & brentZeroHistory()
utCheckLimit worst, 0.5, nWarn
utCompareLessEqual "brentZero(1#, 2#, erAbs, 0#, 0#, -999&) (x-r)^19 evals", _
  brentZeroEvals(), EvalMaxDefault_c, nWarn
utTeeOut

' very nonlinear near root, min near underflow; evaluation count limit = 145
x = brentZero(1#, 2#, erAbs, 0#, 0#, 145&)
utCompareAbs "brentZero(1#, 2#, erAbs, 0, 0, 145) (x-r)^19 x", x, r, worst
utTeeOut "Why: " & brentZeroWhy() & "  evals " & brentZeroEvals()
utTeeOut "History: " & brentZeroHistory()
utCheckLimit worst, 1E-16, nWarn
utCompareLessEqual "brentZero(1#, 2#, erAbs, 0#, 0#, 145) (x-r)^19 evals", _
  brentZeroEvals(), 145&, nWarn
utTeeOut

' evilness parameters
Const A As Double = 0.000000000000001
Const B As Double = A * A
Const C As Double = 2# * A
Const D As Double = 1E-99

utTeeOut "==== Cases below all use the evil function:" & EOL & _
  "  f(x) = " & Format$(C, "0E-0") & _
  " * (x - r) / ((x - r) * (x - r) + " & Format$(B, "0E-0") & ") + " & _
  Format$(D, "0E-0")
utTeeOut

utTeeOut "Error tolerance adjusted steadily tighter."
utTeeOut

' increasing accuracy
Dim e As Double
e = 1#
Dim nWant As Long

Dim j As Long
Dim ec As Double  ' comparison value for error
brentZeroFunctionSet -1008&
For j = 1& To 17&
  ' for very low desired error, the routine will adjust the requested
  ' error value up above the roundoff limit, so we allow for that by
  ' increasing the warning level
  If j >= 16& Then
    ec = 9E-16
  Else
    ec = e
  End If
  x = brentZero(1#, 2#, e)
  utCompareAbs "brentZero(1#, 2#, " & Format$(e, "0E-0") & ") fEvil x", _
    x, r, worst
  utTeeOut "Why: " & brentZeroWhy() & "  evals " & brentZeroEvals()
  utTeeOut "History: " & brentZeroHistory()
  utCheckLimit worst, ec, nWarn
  utCompareLessEqual _
    "brentZero(1#, 2#, " & Format$(e, "0E-0") & ") Fevil bracket", _
    brentZeroBracketWidth(), ec, nWarn
  nWant = Array(2, 7, 10, 15, 19, 23, 28, 32, 37, 41, 46, 50, 54, 59, 63, _
    66, 67)(j - 1&)
  utCompareLessEqual "brentZero(1#, 2#, " & Format$(e, "0E-0") & _
    ") fEvil evals", brentZeroEvals(), nWant, nWarn
  utTeeOut
  e = e / 10#
Next j

' write many requested & returned bracket sizes to file for plotting
Dim fn As Integer, bw As Double
fn = FreeFile
' note: with the "./" prefix, file will go to Project or EXE folder under VB6,
' and to "My Documents" root under VBA
Open "./WantGot.txt" For Output Access Write Lock Read Write As #fn
Print #fn, "# Requested and returned bracket values - good & evil functions"
For ec = 0# To -16.01 Step -0.05
  e = 10# ^ ec
  brentZeroFunctionSet -1005&  ' Sin(x - r)
  x = brentZero(2#, 1#, e)
  bw = brentZeroBracketWidth()
  brentZeroFunctionSet -1008&  ' fEvil
  x = brentZero(2#, 1#, e)
  Print #fn, e; ", "; bw; ", "; brentZeroBracketWidth()
Next ec
Close #fn

brentZeroFunctionSet -1005&  ' Sin(x - r)
' check behavior when unexpectadly large bracket is returned for a simple
' function
utTeeOut "Show shifting from full-error bracket to half-error bracket:"
e = 0.000294134
x = brentZero(1#, 2#, e)
utTeeOut "Function = Sin(x - r)  xErrAbs = " & e & _
  " bracket = " & brentZeroBracketWidth()
utTeeOut "Why: " & brentZeroWhy() & "  evals " & brentZeroEvals()
utTeeOut "History: " & brentZeroHistory()
e = 0.000294133
x = brentZero(1#, 2#, e)
utTeeOut "Function = Sin(x - r)  xErrAbs = " & e & _
  " bracket = " & brentZeroBracketWidth()
utTeeOut "Why: " & brentZeroWhy() & "  evals " & brentZeroEvals()
utTeeOut "History: " & brentZeroHistory()
utTeeOut

utTeeOut "---------------------------------------------------------------------"
If nWarn = 0& Then
  utTeeOut "Unit test success - no warnings were raised."
Else
  utTeeOut "Unit test FAILURE! - warning count: " & nWarn
End If
utTeeOut "---------------------------------------------------------------------"
utTeeOut
utTeeOut "--- Test complete ---"
utFileClose
End Sub

#End If  ' UnitTest

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


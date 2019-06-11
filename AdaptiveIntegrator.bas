Attribute VB_Name = "AdaptiveIntegrator"
'    _      _           _   _         ___     _                     _.
'   /_\  __| |__ _ _ __| |_(_)_ _____|_ _|_ _| |_ ___ __ _ _ _ __ _| |_ ___ _ _.
'  / _ \/ _` / _` | '_ \  _| \ V / -_)| || ' \  _/ -_) _` | '_/ _` |  _/ _ \ '_|
' /_/ \_\__,_\__,_| .__/\__|_|\_/\___|___|_||_\__\___\__, |_| \__,_|\__\___/_|
'                 |_|                                |___/
'         __       __           ______                 __         __
'    __  / /___   / /   ___    /_  __/____ ___  ___   / /  ___   / /__ _  ___.
'   / /_/ // _ \ / _ \ / _ \    / /  / __// -_)/ _ \ / _ \/ _ \ / //  ' \/ -_)
'   \____/ \___//_//_//_//_/   /_/  /_/   \__//_//_//_//_/\___//_//_/_/_/\__/
'
'###############################################################################
'#
'# Visual Basic (VBA or VB6) Module "AdaptiveIntegrator"
'# Saved in text file "AdaptiveIntegrator.bas"
'#
'# Faster than Gaussian integration...  more powerful than the trapezoidal
'# rule...  able to leap tall function peaks with a single bound... it's
'#
'#                *** Adaptive Integrator ***
'#
'# Our motto:  "We integrate anything!"
'#
'# The routine "aiIntegrate" in this module carries out adaptive Simpson
'# integration of a user-specified real univariate function. The output is a
'# numeric estimate of the quantity
'#
'#                  / xEnd
'#                 |
'#      integral = | Func(x) dx
'#                 |
'#                / xBeg
'#
'# where the finite lower limit "xBeg", the finite upper limit "xEnd", and the
'# function to be integrated "Func(x)" are supplied by the user (see usage
'# below). The user may also optionally supply control parameters:
'#    a relative-error tolerance per sub-segment          [default 3E-14]
'#    a max sub-segment size, or (if negative) min count  [default -100]
'#    the maximum number of function calls allowed        [default 100,000]
'# The default values of these quantities will usually supply a good result.
'#
'# In addition to the return of the approximation to the integral, this
'# function also saves diagnostic information that can be retrieved after it
'# finishes. See the getAiXxx Functions for details.
'#
'# Because Visual Basic can't pass functions by name, this module always
'# integrates the user-coded function "aiIntegrand(x [, N])". You can do
'# different integrals by writing code in a specific "If ... ElseIf..." block
'# inside "aiIntegrand" and then changing the optional which-function Long index
'# N (default 0&) to match.
'#
'# Devised and coded by John Trenholme - started Sept 1970 (at NRL)
'#
'# This Module exports the following routines:
'#
'# Function AdaptiveIntegratorVersion
'# Function aiIntegrand
'# Function aiIntegrate
'# Function getAiCallCount
'# Function getAiIntegral
'# Function getAiInvoked
'# Function getAiLastGoodSegEnd
'# Function getAiLeastSegBegin
'# Function getAiLeastSegSize
'# Function getAiParam
'# Function getAiResultCode
'# Function getAiResultText
'# Function getAiStackDepth
'# Sub setAiInvoked
'# Sub setAiParam
'#
'# Sub adaptIntUnitTest  if UnitTest_ is True
'#
'###############################################################################

Option Base 0          ' array base value when not specified - the default
Option Compare Binary  ' string comparison based on Asc(char) - the default
Option Explicit        ' forces explicit variable declaration - changes default
' Option Private Module  ' no effect in VB6; visible-this-Project-only in VBA

' Module-global Const values (convention: start with Upper-case; suffix "_")
Private Const Version_ As String = "2012-11-28"  ' update manually on each edit
Private Const File_ As String = "AdaptiveIntegrator[" & Version_ & "]."

Private Const BadArg_ As Long = 5&  ' = "Invalid procedure call or argument"
Private Const BadSub_ As Long = 9&  ' = "Subscript out of range"
Private Const EOL_ As String = vbNewLine  ' shorthand for end-of-line char pair
Private Const Pi_ As Double = 3.1415926 + 5.358979324E-08  ' good to 53 bits

' Set this True to include unit-test code, False to exclude it.
#Const UnitTest_ = True
' #Const UnitTest_ = False

'############################## Exported Constants ############################
Public Const AiNoIntegration As Long = 0&  ' default value at startup
Public Const AiResultCodeMin As Long = 950&  ' set this to an unused value
Public Const AiSuccess As Long = AiResultCodeMin
Public Const AiTooManyCalls As Long = AiSuccess + 1&
Public Const AiSegmentTooSmall As Long = AiTooManyCalls + 1&
Public Const AiResultCodeMax As Long = AiSegmentTooSmall

'############################## Module-Global Variables #######################
' Module-global variables (convention: suffix "_m")
' Retained between calls; initialize as 0, "" or False (etc.)

Private functionCalls_m As Long
Private integral_m As Double
Private invoked_m As Long
Private relErr_m As Double
Private resultCode_m As Long
Private lastGoodSegmentEnd_m As Double
Private leastSegmentBegin_m As Double
Private leastSegmentSize_m As Double
Private stackDepth_m As Long

' note: the "side parameters" for 'aiIntegrand' are in an array of Variants
' this allows the passing of many different Types, at a small speed cost
Private Const ParsMin As Long = 0&   ' for users used to C or C++
Private Const ParsMax As Long = 99&  ' it's not likely you'll need this many!
Private p_m(ParsMin To ParsMax) As Variant  ' the "side information" parameters

' Variable used only if we are conducting unit tests
#If UnitTest_ Then
Private ofi_m As Integer  ' output-file index used by unit-test routine
#End If

'############################# Exported Routines ###############################

'===============================================================================
Public Function AdaptiveIntegratorVersion(Optional ByVal trigger As Variant) _
As String
' Date of the latest revision to this code, as a string with format "yyyy-mm-dd"
AdaptiveIntegratorVersion = Version_
End Function

'===============================================================================
Public Function aiIntegrand(ByVal x As Double, _
  Optional ByVal which As Long = 0&) _
As Double
' User-coded functions to be integrated go here. Using which = 0, or another
' value if you need several integrands. Write code to evaluate your integrand
' as a function of x and set "f" to its value, such as f = Sin(x).
' You can use side parameters during the calculation. You set the parameters
' before calling "aiIntegrate" by using "setAiParam(j)" and read them here as
' p_m(j).  The index J can run from 0 to a large value (set above). See the
' code for which = -9 here, and in the unit tests, for an example.
Const ID_c As String = File_ & "aiIntegrand"
Dim errNum As Long, errDes As String  ' for saving error Number & Description
On Error GoTo ErrHandler

Dim f As Double  ' function value
Dim t1 As Double, t2 As Double, t3 As Double  ' general-purpose temp values

' If you know the call where there was trouble, set # & uncomment this to debug
'If # = functionCalls_m Then Stop

'---------- user-coded integrand case(s) go here -------------------------------

If 0& = which Then  ' do default case - if only one integrand, put it here
  f = 0.5 / (Sqr(x - 0.9999) * (Sqr(1.0001) - Sqr(0.0001)))  ' near-singular
ElseIf 1& = which Then  ' do an indexed case
  f = 1.5 * Pi_ * Sin(3# * Pi_ * (x - 1#))
  
'---------- unit-test integrand cases go here - they have negative indices -----

ElseIf -1& = which Then  ' quadratic - integrates exactly
  f = 6# * (x - 1#) * (2# - x)
ElseIf -2& = which Then  ' cubic - integrates exactly
  t1 = x - 1#
  f = t1 * t1 * (10# - 4# * x)
ElseIf -3& = which Then  ' square of quadratic has 4th power - not exact
  t1 = (x - 1#) * (2# - x)
  f = 30# * t1 * t1
ElseIf -4& = which Then  ' half-sine
  f = 0.5 * Pi_ * Sin(Pi_ * (x - 1#))
ElseIf -5& = which Then  ' narrow peak
  t1 = (x - 0.5 * Pi_) / 0.01
  f = Exp(-0.5 * t1 * t1) / (0.01 * Sqr(2# * Pi_))
ElseIf (-6& = which) Or (-7& = which) Then  ' oscillatory
  t1 = x * x
  f = t1 * t1 - 5.2 + 3# * Sin(34# * Pi_ * (x - 1#))
ElseIf -8& = which Then  ' step
  f = IIf(x > 0.5 * Pi_, 0#, 1# / (0.5 * Pi_ - 1#))
ElseIf -9& = which Then  ' near singular - uses "side parameter" values
  ' note: if youy modify the test case params, change the test value here
  If 0.000000001 <> p_m(0&) Then  ' called without setting param's; use defaults
    p_m(0&) = 0.000000001
    p_m(1&) = 0.5 / ((Sqr(1# + p_m(0&)) - Sqr(p_m(0&))))
  End If
  f = p_m(1&) / Sqr(x - 1# + p_m(0&))
ElseIf -10& = which Then  ' lots of narrow spikes
  ' note: if youy modify the test case params, change the test value here
  If 0.999 <> p_m(0&) Then  ' called without setting param's; use defaults
    p_m(1&) = 0.999
    p_m(2&) = 11#
    p_m(0&) = Sqr(1# - p_m(1&) * p_m(1&))
  End If
  f = p_m(0&) / (1# + p_m(1&) * Sin(p_m(2&) * 2# * Pi_ * x))
ElseIf -11& = which Then  ' singularity at a almost-not-binary-fraction value
  f = 1# / Abs(x - 1# - 1# / Pi_)  ' forces a segment-too-small error
ElseIf -12& = which Then  ' singularity at a binary-fraction value
  f = 1# / Abs(x - 1.53125)  ' forces div/0 error on call 9
Else  ' caller supplied an undefined 'which' value - report as a bad argument
  Err.Raise BadArg_, ID_c, _
    "Unknown integrand-type value 'which'" & EOL_ & _
    "Change to a known case, or write code for this case"
End If

functionCalls_m = functionCalls_m + 1&  ' update function call count
' uncomment the following line to get an Immediate-window evaluation trace
' Debug.Print which; Tab(6); functionCalls_m; Tab(14); x; Tab(35); f
aiIntegrand = f
Exit Function  '----------------------------------------------------------------

ErrHandler:  '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
errNum = Err.Number  ' save error number
errDes = Err.Description  ' save error-description text
' supplement Description; did error come from a called routine or this routine?
errDes = errDes & EOL_ & _
  "x = " & x & "  which = " & which & "  calls = " & functionCalls_m & EOL_ & _
  IIf(0& = InStr(errDes, "Problem in"), "Problem in ", "Called by ") & ID_c & _
  IIf(0 <> Erl, " line " & Erl, "")  ' undocumented error-line number
'If Designing_C Then Stop: Resume ' hit F8 twice to return to error point
' re-raise error with this routine's ID as Source, and appended to Message
Err.Raise errNum, ID_c, errDes  ' stops here if 'Debug' selected (not 'End')
Resume  ' set Next Statement here and hit F8 to return to error point
End Function

'===============================================================================
Public Function aiIntegrate( _
  ByVal xBeg As Double, _
  ByVal xEnd As Double, _
  Optional ByVal which As Long = 0&, _
  Optional ByVal relErr As Double = 0.00000000000003, _
  Optional ByVal dxMax As Double = -100#, _
  Optional ByVal maxCalls As Long = 100000) _
As Double
' This is the main worker routine in this Module, along with "aiIntegrand".
' It trys to numerically integrate a user-coded real univariate integrand
' function from "xBeg" to "xEnd". The function to be used is specified by
' the index "which" (default 0&). The function resides in "aiIntegrand" in this
' Module, where the user must write their specific integrand code.
'
' Arguments:
'  xBeg     Start of the integration interval.
'  xEnd     End of the integration interval.
'           Note: if xBeg and xEnd are reversed, the integral changes sign.
'           and so do other direction-controlled results.
'  which    Long index of function to integrate (default 0); see "aiIntegrand"
'  relErr   Approximate relative error per sub-interval. The relative error in
'           the result will be approximately this, within an order of magnitude
'           Defaults to 3E-14, which will give high accuracy.
'  dxMax    Largest sub-interval allowed. Set to small value to assure all areas
'           of integrand are sampled (finds isolated peaks). If negative, sets
'           largest sub-interval to whole interval divided by Abs(dxMax). If
'           zero, sets dxMax to entire interval. Defaults to -100; that causes
'           at least 129 evenly-spaced evaluations
'  maxCalls Maximum number of integrand function calls allowed. If this is
'           exceeded, the result is not good, perhaps due to a singularity
'           or other bad behavior. Increase relErr, or increase maxCalls, but
'           be cautious - you may be in trouble. Defaults to 100,000.
'
' Results:
'  returns an approximation to the requested integral as v = aiIntegrate(...
'
' After return, a number of informative values are available by calling:
'  getAiCallCount      Number of function calls during latest invocation
'  getAiIntegral       Value returned by function, or best result after error
'  getAiInvoked        Number of times called since cold start or reset
'  getAiLastGoodSegEnd Last good position after too-many-calls error
'  getAiLeastSegBegin  Start location of smallest segment
'  getAiLeastSegSize   Length of smallest segment
'  getAiResultCode     Numeric code explaining what happened during the call
'  getAiResultText     Text explaining the meaning of a result code
'  getAiStackDepth     Maximum number of segments stacked during this invocation
'
' Errors:
'  Raises AiSegmentTooSmall if a sub-segment size approaches the roundoff limit
'    This probably indicates an attempt to integrate a singular function
'  Raises AiTooManyCalls if the specified call count is exceeded
'    Note that there will always be at least 5 calls, and always an odd number
'  Also, aiIntegrand raises Error 5 = "Invalid procedure call or argument"
'  if there is no code section corresponding to "which", and may raise other
'  "regular" errors such as divide-by-zero, overflow, and so forth.
'
'  See an example of error trapping in the unit test code.
'
' If integration fails due to segment underflow or excessive calls, you can
' recover a portion of the integral. The result from xBeg to getAiLastGoodSegEnd
' is good, and can be found by integrating between those points with all other
' arguments the same. You still have to deal with the rest of the original
' interval, in which the integrand may be singular or unpleasantly close to it.
'
' Passing side information to an integrand:
'  Some integrands require more information than just the point of evaluation.
'  To supply this information (anything that will go into a Variant) there are a
'  number of side-parameter slots. Use "setAiParam index, value" to set a value
'  before calling aiIntegrate, and then you can use p_m(index) to access the
'  value in "aiIntegrand" code (see the example under which = -9, including the
'  integrand and unit-test code). Index values must be 0, 1, 2, ... up to the
'  maximum value ParsMax, which can be adjusted in the unlikely case that it is
'  too small. Just in case, you can see parameters from outside this Module by
'  using getAiParam(index).
'
' Implementation notes:
' This routine can numerically integrate almost any bounded function between
' two given finite limits (see numerical analysis books for methods of
' converting infinite limits to finite ones). It first does the whole interval
' by Simpson's rule in one or more equal segments, using three function values
' in each segment (note that end values are shared). It then splits segments in
' two and does a Simpson evaluation in each half by adding function points in
' the centers of the two sub-segments and finding the correction to the full-
' segment value. If the resulting correction to the integral is less than a
' user-supplied tolerance times the improved value, that segment is done. If it
' is not less, the upper half-segment is stacked and the lower half-segment is
' integrated in the same manner as the whole segment was. The splitting and
' stacking continues by tail recursion until the desired accuracy is attained
' in the lower segment. The next segment up is then unstacked and split. This
' continues until all segments have been integrated. This method concentrates
' attention on regions where the integrand is bumpy and large, while smooth or
' low regions get less effort expended on them. The routine quits if too many
' function evaluations have been made, returning a rough guess at the value of
' the integral. The result of getAiLastGoodSegEnd() is the last point where
' the tolerance was met, giving a known-integrable region to work with. The
' usual cause of the too-many-calls failure is an actual singularity in the
' region of integration. The routine can also return an erroneous value when
' the error criterion is met by accident due to the shape of the integrand. The
' probability of this kind of failure can be reduced by starting with segments
' small compared to the the entire initial segment, using the "dxMax" argument.
' Another way to reduce the probability of an erroneous return is to split the
' initial integration segment into two somewhat unequal portions and to compare
' the sum of the integrals with the integral over the whole segment. Errors can
' also arise if the integrand is oscillatory and the value of the integral is
' zero or near zero. Again the cure is small initial segments, or unequal
' splitting of the segment and addition of the separate integrals, or addition
' of a constant to the integrand followed by subtraction of the added offset.
' Experience has shown that the value returned by this routine usually has an
' error comparable to the per-segment tolerance, within an order of magnitude or
' so, so it is usually safe to set the error tolerance to the desired overall
' error, or maybe (say) ten times less. Note that polynomials up to cubic will
' be integrated exactly, after only 5 function calls.
Const ID_c As String = File_ & "aiIntegrate"

Dim errNum As Long, errDes As String  ' for saving error Number & Description
On Error GoTo ErrHandler

invoked_m = invoked_m + 1&  ' number of calls since load or reset
' If you know the call where there was trouble, set # & uncomment this to debug
'If # = invoked_m Then Stop

Const MinRelErr_c As Double = 4.4E-16  ' disallow segment sizes near roundoff
If relErr < MinRelErr_c Then relErr = MinRelErr_c
relErr_m = relErr

Dim dxMin As Double  ' minimum segment size allowed - near roundoff limit
dxMin = MinRelErr_c * Abs(xEnd - xBeg)

' special handling: negative "dxMax" values specify how many segments in domain
If dxMax < 0# Then dxMax = Abs((xEnd - xBeg) / dxMax)
' zero defaults to entire interval, since it's impossible
If 0# = dxMax Then dxMax = Abs(xEnd - xBeg)
' don't request more segments than implied by call maximum
If maxCalls < 5& Then maxCalls = 5&  ' impose sanity first
If Abs((xEnd - xBeg) / dxMax) > maxCalls - 1& Then
  dxMax = Abs(xEnd - xBeg) / (maxCalls - 1&)
End If

' specify things related to the stack that holds pending segments
Const StakBase_c As Long = 1&       ' it's best to leave this at 1
Const StakIncr_c As Long = 100&     ' even difficult integrals seldom go over 50
Const Fmid_c As Long = StakBase_c   ' index of midpoint function value
Const Xtop_c As Long = Fmid_c + 1&  ' index of top point location
Const Ftop_c As Long = Xtop_c + 1&  ' index of top point function value
Dim stak() As Double                ' stack is in a dynamically allocated array
ReDim stak(Fmid_c To Ftop_c, StakBase_c To StakIncr_c)  ' initial stack space
Dim stakPtr As Long                 ' location of stacked sub-segment
stakPtr = StakBase_c
stackDepth_m = 0&  ' max-depth monitor; 0 depth at start

' beginning, center, and end of initial segment A-C-E
Dim xa As Double, xc As Double, xe As Double
xa = xBeg
xc = xBeg + 0.5 * (xEnd - xBeg)  ' this formulation reduces overflow
xe = xEnd

Dim fa As Double, fc As Double, fe As Double  ' function values in segment
functionCalls_m = 0&  ' zero function call counter; bumped in "aiIntegrand"
fa = aiIntegrand(xa, which)
fc = aiIntegrand(xc, which)
fe = aiIntegrand(xe, which)

Dim integral As Double  ' this is the running estimate of the integral (times 3)
integral = 0.5 * (xe - xa) * (fa + 4# * fc + fe)  ' Simpson's rule

lastGoodSegmentEnd_m = xa     ' initially, nothing is good
leastSegmentSize_m = xe - xa  ' initially, smallest segment is entire interval

Do  ' main loop - split active segment in two, and see if that's good enough
  Dim dx As Double, dx2 As Double  ' sub-segment-related lengths
  dx = 0.25 * (xe - xa)  ' distance from segment ends to new midpoints
  dx2 = dx + dx  ' total sub-segment width
  If Abs(leastSegmentSize_m) > 1.5 * Abs(dx2) Then
    ' update smallest-segment-seen size; 1.5 allows for for roundoff jitter
    leastSegmentSize_m = dx2
    leastSegmentBegin_m = xa
  End If
  If Abs(dx2) < dxMin Then
    ' new segment size near roundoff limit for IEEE 754, so fail
    ' avoid this problem by using a less strict error tolerance "relErr"
    integral_m = integral / 3#  ' probably wrong; might be useful
    Err.Raise AiSegmentTooSmall, ID_c, getAiResultText(AiSegmentTooSmall)
  End If
  ' points in segment are A - B - C - D - E where A - C - E is old segment
  ' and new sub-segments will be A - B - C and C - D - E
  Dim xb As Double, fb As Double, xd As Double, fd As Double
  ' new midpoint of left segment
  xb = xa + dx  ' we could just use xa + dx in fb call; useful for tracing
  fb = aiIntegrand(xb, which)
  ' new midpoint of right segment
  xd = xe - dx  ' we could just use xe - dx in fd call; useful for tracing
  fd = aiIntegrand(xd, which)
  ' find the integral's change from the old segment to the two new segments
  Dim deltaIntegral As Double
  ' subtract old Simpson contribution, and add on Simpson from each half
  deltaIntegral = dx * (4# * (fb + fd) - 6# * fc - fa - fe)  ' still times 3
  integral = integral + deltaIntegral  ' replace old segment value with new ones
  If functionCalls_m > maxCalls Then
    ' quit on call count (after update); may go over limit by 1
    ' avoid this problem by using a less strict error tolerance "relErr"
    ' or by raising "maxCalls"
    integral_m = integral / 3#  ' probably wrong; might be useful
    Err.Raise AiTooManyCalls, ID_c, getAiResultText(AiTooManyCalls)
  End If
  Dim small As Boolean  ' was change from this update below tolerance?
  ' note that this may be difficult to meet if the integral is zero or near it
  small = Abs(deltaIntegral) <= Abs(relErr * integral)
  ' done with this segment if both change and segment size are small enough
  If small And (Abs(dx) <= dxMax) Then  ' done with this segment
    lastGoodSegmentEnd_m = xe
    stakPtr = stakPtr - 1&  ' pre-decrement to next lower stack position
    If stakPtr < StakBase_c Then  ' stack is empty; last segment; we are done
      Exit Do  ' job done; this is the only non-error exit from the  loop
    Else  ' we have one or more pending segments on the stack
      xa = xe  ' new bottom is old top, since stacked segment is just above
      fa = fe
      fc = stak(Fmid_c, stakPtr)  ' pop midpoint value (we don't need xc)
      xe = stak(Xtop_c, stakPtr)  ' pop top point location
      fe = stak(Ftop_c, stakPtr)  ' pop top point value
    End If
  Else  ' change or segment too large; split, push & do tail recursion
    If stakPtr > UBound(stak, 2&) Then  ' out of stack space; add more
      Beep  ' this is so unlikely that we complain to the user
      ReDim Preserve stak(Fmid_c To Ftop_c, _
        StakBase_c To UBound(stak, 2&) + StakIncr_c)
    End If
    stak(Fmid_c, stakPtr) = fd  ' push midpoint value (we don't need xd)
    stak(Xtop_c, stakPtr) = xe  ' push top point location
    stak(Ftop_c, stakPtr) = fe  ' push top point value
    ' keep track of maximum stack depth
    Dim deep As Long
    deep = stakPtr - StakBase_c + 1&  ' count of stack positions in use
    If stackDepth_m < deep Then stackDepth_m = deep
    ' post-increment to next stack position to use
    stakPtr = stakPtr + 1&
    ' set up the tail recursion; active segment = old lower half
    xe = xa + 0.5 * (xe - xa)  ' this formulation reduces overflow
    fe = fc
    fc = fb
  End If
Loop  ' go back and work on the active segment

Erase stak  ' should happen automatically; just being cautious
resultCode_m = AiSuccess
leastSegmentSize_m = 2# * leastSegmentSize_m  ' we used half-segment size
' we have ignored an integrand factor of 3 until now for speed; return it
integral_m = integral / 3#
aiIntegrate = integral_m
Exit Function  '----------------------------------------------------------------

ErrHandler:  '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
errNum = Err.Number  ' save error number
resultCode_m = errNum  ' stash it where the user can recover it later
errDes = Err.Description  ' save error-description text
' supplement Description; did error come from a called routine or this routine?
errDes = errDes & EOL_ & _
  "xBeg = " & xBeg & " xEnd = " & xEnd & " which = " & which & _
    "  invoked = " & invoked_m & EOL_ & _
  IIf(0& = InStr(errDes, "Problem in"), "Problem in ", "Called by ") & ID_c & _
  IIf(0 <> Erl, " line " & Erl, "")  ' undocumented error-line number
'If Designing_C Then Stop: Resume ' hit F8 twice to return to error point
' re-raise error with this routine's ID as Source, and appended to Message
Err.Raise errNum, ID_c, errDes  ' stops here if 'Debug' selected (not 'End')
Resume
End Function

'===============================================================================
Public Function getAiCallCount(Optional ByVal trigger As Variant) As Long
' Number of integrand function calls during latest invocation. Odd and >= 5.
getAiCallCount = functionCalls_m
End Function

'===============================================================================
Public Function getAiIntegral(Optional ByVal trigger As Variant) As Double
' Value if the integral. May be for less than the entire interval, or of lower
' accuracy than desired, if there was an error.
getAiIntegral = integral_m
End Function

'===============================================================================
Public Function getAiInvoked(Optional ByVal trigger As Variant) As Long
' Number of times aiIntegrate has been called since cold start or reset. See
' "setAiInvoked()". Useful to see how much work was needed to do a user's job.
getAiInvoked = invoked_m
End Function

'===============================================================================
Public Function getAiLastGoodSegEnd(Optional ByVal trigger As Variant) As Double
' The point where all previous segments have passed tolerance. If there is a
' "too-many-calls" error, re-integration with the upper limit set to this value,
' and all other arguments the same, will complete successfully. You can then try
' to deal with the rest of the interval of integration.
getAiLastGoodSegEnd = lastGoodSegmentEnd_m
End Function

'===============================================================================
Public Function getAiLeastSegBegin(Optional ByVal trigger As Variant) As Double
' Start point of smallest segment (there may be more segments of this size)
getAiLeastSegBegin = leastSegmentBegin_m
End Function

'===============================================================================
Public Function getAiLeastSegSize(Optional ByVal trigger As Variant) As Double
' Length of smallest segment (there may be more segments of this size)
getAiLeastSegSize = leastSegmentSize_m
End Function

'===============================================================================
Public Function getAiParam(ByVal index As Long) As Variant
' Return the specified parameter value, used by the routine 'aiIntegrand'
' during adaptive Simpson integration. See setAiParam for usage.
Const ID_c As String = File_ & "getAiParam"
If (index < ParsMin) Or (index > ParsMax) Then
  Err.Raise BadSub_, ID_c, _
    "Param index out of range" & EOL_ & _
    "Need " & ParsMin & " <= index <= " & ParsMax & _
    " but index = " & index & EOL_ & "Problem in " & ID_c
End If
getAiParam = p_m(index)
End Function

'===============================================================================
Public Function getAiResultCode(Optional ByVal trigger As Variant) As Long
' Return a numeric code explaining what happened during an integration call. The
' possible return codes unique to this Module are exported as AiXxx constants
' such as "AiSuccess", or the result may be a "standard" Visual Basic error.
' Normally, problems during integration will raise an error, but you can use
' On Error Resume Next before calling aiIntegrate to continue on, and then check
' the result code to guide program logic. See unit test code for examples.
getAiResultCode = resultCode_m
End Function

'===============================================================================
Public Function getAiResultText(ByVal resultCode As Long) As String
' Return a text explanation of the code produced by an integration call. For
' example, do this: Debug.Print getAiResultText(getAiResultCode())
If AiSuccess = resultCode Then
  getAiResultText = "No problems - result probably good to " & relErr_m
ElseIf AiTooManyCalls = resultCode Then
  getAiResultText = "Too many integrand calls: " & functionCalls_m
ElseIf AiSegmentTooSmall = resultCode Then
  getAiResultText = "Segment too small (integrand singularity?):" & EOL_ & _
    "  length = " & leastSegmentSize_m & " starts at x = " & leastSegmentBegin_m
ElseIf AiNoIntegration = resultCode Then  ' still at default startup value
  getAiResultText = "Adaptive integrator has not been called yet"
Else  ' if we got a "standard" VB error code, report it verbatim
  ' can also be triggered by user-defined or object-defined errors, yielding
  ' the uninformative "Application-defined or object-defined error"
  getAiResultText = "VB error code " & resultCode & " = " _
    & Error$(resultCode)
End If
End Function

'===============================================================================
Public Function getAiStackDepth(Optional ByVal trigger As Variant) As Long
' Maximum number of segments stacked during latest invocation of "aiIntegrate"
getAiStackDepth = stackDepth_m
End Function

'===============================================================================
Public Sub setAiInvoked(ByVal newValue As Long)
' Set the number of calls to "aiIntegrate". Use when you want to count the
' number of calls during some process. See "getAiInvoked()".
invoked_m = newValue
End Sub

'===============================================================================
Public Sub setAiParam(ByVal index As Long, ByVal param As Variant)
' Stash a user parameter needed by 'aiIntegrand' where that routine can see it.
' This is used to pass "side information" to the integrand routine.
Const ID_c As String = File_ & "setAiParam"
If (index < ParsMin) Or (index > ParsMax) Then
  Err.Raise BadSub_, ID_c, _
    "Param index out of range (with param = " & param & ")" & EOL_ & _
    "Need " & ParsMin & " <= index <= " & ParsMax & _
    " but index = " & index & EOL_ & "Problem in " & ID_c
End If
p_m(index) = param
End Sub

'############################# Unit Test #######################################

#If UnitTest_ Then  ' set UnitTest_ = True to use unit-test routines

'===============================================================================
Public Sub adaptIntUnitTest()
' This checks for proper implementation and does some numeric tests of the
' output of the routine.
'
' Output goes to a file, & to Immediate window if in VBA or VB6 Editor.
'
' To run this routine from VBA, put the cursor somewhere in it and hit F5.
' It also appears in the Developer | Macros window (Alt-F8)
' To run this from VB6, enter "adaptIntUnitTest" in the Immediate window.
' (If the Immediate window is not open, use View... or Ctrl-G to open it.)
'
' Note: this routine uses formatting functions from "Formats.bas"
Dim tStart As Single
tStart = Timer()

Dim modName As String
modName = Left$(File_, InStr(File_, "[") - 1&)  ' strip off version info

' get path to folder where Workbook resides
Dim path As String
' VBA file opening stuff - we presume we are running under Excel
' note: in Excel, must save a new workbook at least once before path exists
path = Excel.ActiveWorkbook.path
' if there is no workbook path, use CurDir (user's MyDocuments or Documents)
If vbNullString = path Then path = FileSystem.CurDir$
If vbNullString = path Then  ' no current directory?
  VBA.Interaction.Beep
  MsgBox _
    "Unable to find current folder" & EOL_ & _
    "Excel workbook has no disk location, & no CurDir" & EOL_ & _
    "Save workbook to disk before proceeding because" & EOL_ & _
    "we need a known location to write the file to." & EOL_ & _
    "Output will be sent to Immediate window only (Ctrl-G in editor)", _
    vbOKOnly Or vbCritical Or vbMsgBoxSetForeground, _
    File_ & " ERROR - No File Path"
  ofi_m = 0  ' no output to file
Else
  ' be sure path separator is at end of path (only C:\ etc. have it already)
  Dim ps As String
  ps = Application.PathSeparator
  If Right$(path, 1&) <> ps Then path = path & ps
  ' make up full file name, with path
  Dim filename As String
  filename = File_ & "UnitTest.txt"
  Dim ffn As String
  ffn = path & filename
  ofi_m = FreeFile  ' file index, module-global so teeOut can use it
  ' try to open output file, over-writing any existing file
  On Error Resume Next
  Open ffn For Output Access Write Lock Write As #ofi_m
  Dim errNum As Long
  errNum = Err.Number
  On Error GoTo 0     ' clear Err object & enable default error handling
  If 0& <> errNum Then  ' cannot open file; disk full?
    Dim errDesc As String
    errDesc = Err.Description
    VBA.Interaction.Beep
    MsgBox _
      "Unable to open output file:" & EOL_ & _
      """" & filename & """" & EOL_ & _
      "in folder:" & EOL_ & _
      """" & Left$(path, Len(path) - 1&) & """" & EOL_ & _
      "Error: " & errDesc & EOL_ & _
      "Error number: " & errNum & EOL_ & _
     "Output will be sent to Immediate window only (Ctrl-G in editor)", _
      vbOKOnly Or vbCritical Or vbMsgBoxSetForeground, _
      File_ & " ERROR - Can't Open File"
    ofi_m = 0  ' no output to file
  End If
End If

teeOut "########## Test of " & modName & " routines at " & Now()
teeOut "Using Module " & Left$(File_, Len(File_) - 1&)  ' trim trailing "."
teeOut
teeOut "All test cases integrate from x = 1.0 to x = 2.0"
teeOut "and should give an integral of exactly 1.0"

setAiInvoked 0&  ' we want to count the number of "aiIntegrate" calls
Dim integrandCalls As Long
integrandCalls = 0&

Dim j As Long, funcInfo As String, pow10 As Double
For j = 1& To 12&
  Dim kMax As Long, dxMax As Double
  kMax = 16&  ' default max power of 0.1
  dxMax = 9#  ' default max segment size (this value sets no limit)
  If 1& = j Then
    funcInfo = "6 * (x - 1) * (2 - x)  quadratic"
    kMax = 3&
  ElseIf 2& = j Then
    funcInfo = "(x - 1)^2 * (10 - 4 * x)  cubic"
    kMax = 3&
  ElseIf 3& = j Then
    funcInfo = "30 * ((x - 1) * (2 - x))^2  quartic"
  ElseIf 4& = j Then
    funcInfo = "0.5 * Pi * Sin(Pi * (x - 1))  half sine"
  ElseIf 5& = j Then
    funcInfo = _
      "Exp(-0.5 * ((x - Pi / 2) / 0.01)^2 / (0.01 * Sqr(2 * Pi))  narrow peak"
  ElseIf 6& = j Then
    funcInfo = "x^4 - 5.2 + 3 * Sin(34 * Pi * (x - 1))  oscillatory"
  ElseIf 7& = j Then
    dxMax = -500#
    funcInfo = "x^4 - 5.2 + 3 * Sin(34 * Pi * (x - 1))  oscillatory" & _
      " with dxMax = " & dxMax
  ElseIf 8& = j Then
    funcInfo = "IIf(x > 0.5 * Pi, 0#, 1# / (0.5 * Pi - 1#))  step"
  ElseIf 9& = j Then
    Const Eps As Double = 0.000000001
    setAiParam 0&, Eps  ' the smaller, the closer to singular
    setAiParam 1&, 0.5 / ((Sqr(1# + Eps) - Sqr(Eps)))  ' normalize to 1
    funcInfo = "C / Sqr(x - 1 + eps)  nearly singular  eps = " & Eps & EOL_ & _
      "where C = 0.5 / ((Sqr(1 + eps) - Sqr(eps))) = " & getAiParam(1&)
  ElseIf 10& = j Then
    Const Size As Double = 0.999  ' the closer to 1, the taller the spikes
    Const Ripl As Double = 11#
    setAiParam 0&, Sqr(1# - Size * Size)  ' normalize to 1
    setAiParam 1&, Size
    setAiParam 2&, Ripl
    funcInfo = "A / (1 + B * Sin(C * 2 * Pi * x))  spikes" & EOL_ & _
      "where B = " & Size & "  C = " & Ripl & _
      "  A = Sqr(1 - B^2) = " & getAiParam(0&)
  ElseIf 11& = j Then
    kMax = 3&
    funcInfo = "1 / Abs(x - 1 - 1 / Pi)  singularity not at binary fraction"
  ElseIf 12& = j Then
    kMax = 2&
    funcInfo = "1 / Abs(x - 1.53125)  singularity at binary fraction"
  Else
    Exit For
  End If
  teeOut
  teeOut "====== Test case " & j & " ======"
  teeOut "Integrating " & funcInfo
  teeOut "Req.Err  Returned          Rel.Error  Calls  Deep " & _
    "Min. Seg. Begin   Min. Seg. End"
  teeOut "-------  ----------------- --------- ------ ---- " & _
    "----------------- -----------------"
  pow10 = 1#
  Dim k As Long, result As Double
  For k = 0& To kMax
    On Error Resume Next  ' go on past any error
    result = aiIntegrate(1#, 2#, -j, 1# / pow10, dxMax, 200000)
    errNum = Err.Number
    Dim errDes As String
    errDes = Err.Description
    On Error GoTo 0
    integrandCalls = integrandCalls + getAiCallCount()
    If 0& = errNum Then  ' no error, so report results
      teeOut cram(1# / pow10, 7&, False) & "  " & cram(result, 17&) & " " & _
        cram(result - 1#, 9&, False) & strFitR(CStr(getAiCallCount()), 7&) & _
        strFitR(CStr(getAiStackDepth()), 5&) & " " & _
        cram(getAiLeastSegBegin(), 17&) & " " & _
        cram(getAiLeastSegBegin() + getAiLeastSegSize(), 17&)
    Else  ' there was an error during integration
      teeOut "== Error #" & errNum & "  Description:"
      teeOut errDes
      teeOut "== getAiResultText(errNum):"
      teeOut getAiResultText(errNum)
      teeOut "Sub-segment tolerance met only from 1.0 to " & getAiLastGoodSegEnd()
      Exit For  ' don't try any harder on this problem
    End If
    pow10 = pow10 * 10#  ' this is exact, but 0.1 isn't an exact binary fraction
  Next k
Next j

teeOut
teeOut "aiIntegrate was called " & getAiInvoked() & " times"
teeOut "aiIntegrand was called " & Format$(integrandCalls, "0,0") & " times"
teeOut "Unit tests complete in " & Round(Timer() - tStart, 3&) & " seconds"
teeOut "~~~~~~~~~~ " & modName & _
  " unit tests ~~~~~~~~ end of file ~~~~~~~~~~"

Close ofi_m
End Sub

'-------------------------------------------------------------------------------
Private Sub teeOut(Optional ByRef str As String = vbNullString)
' Send the supplied string to the output file (if it is open) and to the
' Immediate window (Ctrl-G to open) if in VB editor.
' Unit-test support routine - John Trenholme - 9 Jul 2002
Debug.Print str  ' send to Immediate window (only if in Editor; limited size)
If 0 <> ofi_m Then Print #ofi_m, str  ' send to file, if it's open
End Sub

#End If  ' UnitTest_

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


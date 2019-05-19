Attribute VB_Name = "NelderMeadF"
Attribute VB_Description = "Module with Function that reduces an arbitrary function of 2 or more parameters."
'     _   _        _      _              __  __                   _  ______
'    | \ | |      | |    | |            |  \/  |                 | ||  ____|
'    |  \| |  ___ | |  __| |  ___  _ __ | \  / |  ___   __ _   __| || |__
'    | . ` | / _ \| | / _` | / _ \| '__|| |\/| | / _ \ / _` | / _` ||  __|
'    | |\  ||  __/| || (_| ||  __/| |   | |  | ||  __/| (_| || (_| || |
'    |_| \_| \___||_| \__,_| \___||_|   |_|  |_| \___| \__,_| \__,_||_|
'
'###############################################################################
'#
'# Visual Basic 6 and VBA source file "NelderMeadF.bas"
'#
'# John Trenholme - initial version 13 Aug 2003 (as "NelderMead.bas")
'#
'# This routine reduces the value of an unconstrained multidimensional function
'# by adjusting an array of two or more Double variables. The function or
'# functions to be reduced is/are coded by the user in the Function NMfunc().
'#
'# This Module exports the following:
'#
'#   Function NMgetFunctionIndex
'#   Function NMgetParams
'#   Function NMlimitExp
'#   Function NMlimitHard
'#   Function NMlimitSqr
'#   Function NMreduce        ' this is the Nelder-Mead algorithm
'#   Function NMresultString
'#   Const NMretXXX  (various reason-for-return codes)
'#   Type NMresultType
'#   Sub NMsetFunctionIndex
'#   Sub NMsetParams
'#   Sub NMsetTraceFileNumber
'#   Sub NMsetTraceLevel
'#   Function NMversion
'#
'# Note: to reduce a nonlinear function of one variable, use "BrentMin."
'#
'###############################################################################

' Usage:
'
' Code NMfunc(p) so that it returns the function you want to reduce, given the
' argument array p().
'
' Dim res as NMresultType
' Dim a0(1& to 6&) as Double  ' initialPosition array (use any dim's you like)
' a0(1&) = 42.24  ' initial value of first argument
'   ... and so forth, for each argument
' res = NMreduce(a0, initSize, valueTol, sizTol, rngTol, callMax)
'
' Results are now in the various components of "res." In particular, the best
' values of the arguments that the routine found are in res.bestVars() and the
' function value at those argument values is in res.bestValue.
'
' The meaning of the various arguments is as follows:
'   initPoint  Array of argument values at the start point (not changed)
'   initSize   Initial size of the simplex
'   valueTol   Return as soon as a function value at or below this is seen
'   sizeTol    Return if the size of the simplex is below this
'   rangeTol   Return if range of values in the simplex is below this
'   callLimit  Return if more than this many function calls are made

Option Compare Binary
Option Explicit
Option Private Module  ' no effect in Visual Basic; globals project-only in VBA

' version (date) of this file
Private Const Version_c As String = "2010-11-17"
Private Const ThisFile As String = "NelderMeadF"  ' ID for this file
Private Const PN_c As String = ThisFile & "[" & Version_c & "]."

Private Const EOL As String = vbNewLine  ' handy abbreviation

'###############################################################################
'#
'#  Public definitions
'#
'###############################################################################

' these values return status when reduction is complete; note 0 is not used
' we don't use an Enum because users can change the capitalization of Enums
Public Const NMretValueTolMet As Long = 1&     ' function value below tolerance
Public Const NMretSizeTolMet As Long = 2&      ' simplex size below tolerance
Public Const NMretRangeTolMet As Long = 3&     ' range of val's in simplex < tol
Public Const NMretTooManyCalls As Long = 4&    ' number of function calls > max
Public Const NMretVarHuge As Long = 5&         ' variable(s) near max Double
Public Const NMretTooFewVars As Long = 6&      ' supplied function had < 2 vars
Public Const NMretWrongArrayBase As Long = 7&  ' initPoint array base not 1

' Type holding results of the reduction; "NMreduce" returns one of these
Public Type NMresultType
  bestValue As Double    ' smallest value seen during reduction
  bestVars() As Double   ' variable values at the point of smallest value
  reason As Long         ' cause of return from the function (NMretXXX value)
  finalSize As Double    ' final size of the Nelder-Mead simplex
  finalRange As Double   ' final value range in the Nelder-Mead simplex
  callsUsed As Long      ' number of function calls made during reduction
  nReflect As Long       ' number of "reflect" moves made during reduction
  nExtend As Long        ' number of "extend" moves made during reduction
  nContract As Long      ' number of "contract" moves made, both "in" & "out"
  nContractOut As Long   ' number of "contract" moves that were outside simplex
  nHuddle As Long        ' number of "huddle" (shrink) actions during reduction
  nInitialize As Long    ' number of initializations made during reduction
  nInvoked As Double     ' number of times "NMreduce" function has been called
  initialVal As Double   ' function value at the initial point
  initialPtBest As Boolean  ' True if the initial point was best
  refOrphanBest As Boolean  ' True if a discarded reflected point was best
End Type

'###############################################################################
'#
'#  Private definitions
'#
'###############################################################################

' ***** Default values for algorithm tuning parameters
Private Const ConDefault As Double = 0.5
Private Const ExtDefault As Double = 1.6
Private Const HudDefault As Double = 0.5
Private Const RefDefault As Double = 1#
Private Const ReSizeDefault As Double = 0.8
Private Const GrowDefault As Double = 1.5
Private Const CycleMulDefault As Double = 2#
Private Const DropMulDefault As Double = 4#

' ***** Numeric limits
Private Const Huge As Double = 1.79769313486231E+308 + 5.7E+293
Private Const Tiny As Double = 2.2250738585072E-308 + 1.48219693752374E-323
Private Const EpsDbl As Double = 2.22044604925031E-16 + 3E-31

' ***** Return codes set by NMtest (internal use only)
Private Const NMretNoDecrease As Long = -1&
Private Const NMretNoReason As Long = 0&

' ***** Argument values and function value at a point
Private Type NMpointType
  vars() As Double
  val As Double
End Type

'###############################################################################
'#
'#  The Module keeps information in the following module-global variables
'#
'###############################################################################

' ***** Algorithm-tuning parameters - adjust with care
Private Con As Double             ' contract distance
Private Ext As Double             ' extend distance
Private Hud As Double             ' huddle distance
Private Ref As Double             ' reflect distance
Private ReSize As Double          ' amount size is grown at restart
Private Grow As Double            ' max amount restart can be larger than init
Private CycleMul As Double        ' this times var count is min calls/cycle
Private DropMul As Double         ' this times var count is max no-drop
Private isInit_m As Boolean  ' True if tuning params set; defaults to False

Private fBot_m As Double  ' function value at lowest point in simplex
Private funcMinCycle_m As Double  ' best function value seen in this cycle
Private functionIndex_m As Long  ' index of the function NMfunc should return
Private genPoint_m As NMpointType  ' point used for several purposes
Private invoked_m As Double  ' number of times this routine has been invoked
Private kBot_m As Long  ' index of lowest point in simplex
Private kMid_m As Long  ' index of next-to-highest point in simplex
Private kTop_m As Long  ' index of highest point in simplex
Private LB_m As Long  ' lower bound of arrays
Private maxCalls_m As Long  ' caller's limit on function calls
Private n_m As Integer  ' file number of debug print file (if any)
Private nCalls_m As Long  ' total number of calls during this invocation
Private nCon_m As Long
Private nConOut_m As Long
Private nCycle_m As Long  ' number of calls during this cycle
Private nExt_m As Long
Private nHud_m As Long
Private nInit_m As Long
Private noDrop_m As Long
Private nRef_m As Long
Private nSimplex_m As Long  ' number of points in simplex
Private nVars_m As Long  ' number of variables we are varying
Private prevReason_m As Long  ' reason for previous result cycle halt
Private prevPoint_m As NMpointType  ' best point at start of this cycle
Private range_m As Double  ' present value range in the simplex
Private rangeTol_m As Double  ' caller's exit tolerance on relative value diff.
Private refPoint_m As NMpointType  ' point we reflect to
Private result_m As NMresultType   ' this holds the return information
Private scale_m() As Double  ' size of variables, used for scaling to unit size
Private simplex_m() As NMpointType  ' the simplex of N+1 points
Private size_m As Double  ' size of the simplex
Private sizeStart_m As Double  ' present-cycle start size of the simplex
Private sizeInit_m As Double  ' caller's initial size of the simplex
Private sizeTol_m As Double  ' caller's exit tolerance on normalized size
Private testResult_m As Long  ' the result of the termination tests
Private traceLevel_m As Long  ' 0=none  1=minimal  2=moderate  3=detailed
Private UB_m As Long  ' upper bound of arrays
Private x_m() As Double  ' temporary variable array
Private valueTol_m As Double  ' caller's exit tolerance on function value

'###############################################################################
'#
'#  Public routines
'#
'###############################################################################

'===============================================================================
Public Function NMgetParams() As Variant
' Return the internal parameters used in the algorithm, in a zero-based array
' of zero-based pairs of (String,Double) values. For example, to get the name
' of Ext use NMgetParams(1&)(0&); get its value with NMgetParams(1&)(1&).
' If the trace file index is non-zero, the values also are printed to that file.
checkInitialization
If n_m <> 0 Then
  Print #n_m, "Nelder-Mead algorithm tuning parameters are:"
  Print #n_m, "   Con      = "; Con
  Print #n_m, "   Ext      = "; Ext
  Print #n_m, "   Hud      = "; Hud
  Print #n_m, "   Ref      = "; Ref
  Print #n_m, "   ReSize   = "; ReSize
  Print #n_m, "   Grow     = "; Grow
  Print #n_m, "   CycleMul = "; CycleMul
  Print #n_m, "   DropMul  = "; DropMul
End If
NMgetParams = Array(Array("Con", Con), Array("Ext", Ext), Array("Hud", Hud), _
  Array("Ref", Ref), Array("ReSize", ReSize), Array("Grow", Grow), _
  Array("CycleMul", CycleMul), Array("DropMul", DropMul))
End Function

'===============================================================================
Public Function NMreduce( _
  ByRef initPoint() As Double, _
  ByVal initSize As Double, _
  ByVal valueTol As Double, _
  ByVal sizeTol As Double, _
  ByVal rangeTol As Double, _
  ByVal callLimit As Long) _
As NMresultType
' This routine carries out the reduction of the function.

' set up error handling
Const ID_c As String = PN_c & "NMreduce Function" ' name of this file & routine
Dim errNum As Long, errDsc As String  ' Err object Property holders
invoked_m = invoked_m + 1#  ' bump NM-reducer-function-called counter
' if in IDE, halt at error; if compiled, supply detailed info & traceback
If Not inDesign() Then On Error GoTo ErrorHandler

If traceLevel_m >= 1& Then
  If n_m <> 0 Then  ' there is an output file
    Dim dStart As Date
    dStart = Date
    Dim tStart As Single
    tStart = Timer()
    Print #n_m, ">>>>>>>>>>>>>>>>>>>> Enter Function NMreduce [" & _
      Version_c & "] at "; Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    Print #n_m, "Reducing user-coded Function NMfunc (Case "; _
      NMgetFunctionIndex() & ")"
    Print #n_m, "This is call number"; invoked_m; "to this routine"
    Print #n_m, "Input values:"
    Print #n_m, "  Initial size:"; initSize
    Print #n_m, "  Value tolerance:"; valueTol
    Print #n_m, "  Size tolerance:"; sizeTol
    Print #n_m, "  Range tolerance:"; rangeTol
    Print #n_m, "  Function call limit:"; callLimit
    Print #n_m, "  Trace level:"; traceLevel_m
  Else  ' no output file; cancel trace printout
    traceLevel_m = 0&
  End If
End If

checkInitialization  ' make sure the tuning parameters are set
If traceLevel_m >= 3& Then
  Print #n_m, "   Algorithm tuning parameters:"
  Print #n_m, "     Con ="; Con
  Print #n_m, "     Ext ="; Ext
  Print #n_m, "     Hud ="; Hud
  Print #n_m, "     Ref ="; Ref
  Print #n_m, "     ReSize ="; ReSize
  Print #n_m, "     Grow ="; Grow
  Print #n_m, "     CycleMul ="; CycleMul
  Print #n_m, "     DropMul ="; DropMul
End If

' localize input variables
sizeInit_m = initSize
valueTol_m = valueTol
sizeTol_m = sizeTol
rangeTol_m = rangeTol
maxCalls_m = callLimit

' localize array limits (using initPoint() as the authority)
LB_m = LBound(initPoint)
UB_m = UBound(initPoint)

sizeStart_m = sizeInit_m   ' present size of the simplex
nVars_m = UB_m - LB_m + 1& ' number of variables to adjust
nSimplex_m = nVars_m + 1&  ' number of points in the simplex

' set default values in result_m, and set activity counters to zero
NMinit

' quit early if user has asked for variation of < 2 variables
If nVars_m < 2& Then
  ReDim result_m.bestVars(1& To 1&) As Double
  result_m.bestVars(1&) = Huge  ' dummy return value
  result_m.reason = NMretTooFewVars
  If traceLevel_m >= 2& Then
    Print #n_m, " Immediate exit; variable count "; nVars_m; " is < 2"
  End If
  GoTo Wrapup
End If

' set array sizes (dynamic memory allocation)
ReDim scale_m(LB_m To UB_m) As Double
ReDim simplex_m(1& To nSimplex_m) As NMpointType
Dim j As Long
For j = LBound(simplex_m) To UBound(simplex_m)
  ReDim simplex_m(j).vars(LB_m To UB_m) As Double
Next j
ReDim x_m(LB_m To UB_m) As Double
ReDim genPoint_m.vars(LB_m To UB_m) As Double
ReDim refPoint_m.vars(LB_m To UB_m) As Double

' set result vars to input point (note variables are not yet normalized)
result_m.bestVars = initPoint  ' ReDim array (allocate mem.) & set values

' do a function call at the initial point (it might turn out to be best)
'-------------------- function call --------------------
result_m.bestValue = NMfunc(initPoint)
'-------------------------------------------------------
nCalls_m = 1&

result_m.initialVal = result_m.bestValue  ' user may want this
result_m.initialPtBest = True

If traceLevel_m >= 2& Then
  Print #n_m, " Number of var's:"; nVars_m; "("; LB_m; "to"; UB_m; ")"
  Print #n_m, " Initial point has value"; result_m.bestValue; " vars:"
  Print #n_m, arrayStr(result_m.bestVars, "  ")
End If

' inner loops of the algorithm - steps are in separate routines for clarity
Dim mustExit As Boolean
Dim reStart As Boolean
Do  ' simplex restart loop
  ' start or restart a simplex minimization cycle
  NMstart
  Do  ' simplex test-move loop
    NMsort
    NMtest  ' set testResult_m for tests below
    ' see if value or size tolerance was met
    reStart = False
    ' see if a quit-right-now condition was seen
    If (testResult_m = NMretValueTolMet) Or _
           (testResult_m = NMretTooManyCalls) Or _
           (testResult_m = NMretVarHuge) Then
      mustExit = True  ' exit restart loop immediately
      Exit Do
    ' note: we do not return as soon as simplex size or value range meets tol
    ' instead, we restart with a small size to be sure we are at a min
    ' this may be excessively cautious, but it helps with especially hard cases
    ElseIf (testResult_m = NMretRangeTolMet) Or _
       (testResult_m = NMretSizeTolMet) Then
      mustExit = False  ' stay in simplex restart loop until no function drop
      Exit Do
    ElseIf testResult_m = NMretNoDecrease Then  ' no function drop in many tries
      mustExit = False  ' stay in simplex restart loop
      reStart = True  ' force a simplex restart
      Exit Do
    End If
    ' replace an old point in the simplex with a new one
    NMmove
  Loop  ' end of test-move loop
  
  If traceLevel_m >= 1& Then
    If testResult_m <> NMretNoDecrease Then
      Print #n_m, "Restart cycle " & nInit_m & " done: " & _
        NMresultString(testResult_m)
    Else
      Print #n_m, "Restart cycle " & nInit_m & " done: no decrease in " & _
        noDrop_m & " calls"
    End If
    Dim k As Long
    For k = 1 To nSimplex_m
      Print #n_m, "  point " & k & "  value = " & simplex_m(k).val & "  vars:"
      For j = LB_m To UB_m
        x_m(j) = simplex_m(k).vars(j) * scale_m(j)
      Next j
      Print #n_m, arrayStr(x_m, "    ")
    Next k
  End If
  
  If mustExit Then Exit Do
' restart if result was improved, or if no function drop in many tries
Loop While (funcMinCycle_m < prevPoint_m.val) Or reStart

' if the previous-cycle result was better, return the reason for quitting then
If funcMinCycle_m > prevPoint_m.val Then
  result_m.reason = prevReason_m
Else
  result_m.reason = testResult_m
End If

' we are done, so set return values (if not already set)
result_m.callsUsed = nCalls_m
result_m.finalSize = size_m
result_m.finalRange = range_m
result_m.nContract = nCon_m
result_m.nContractOut = nConOut_m
result_m.nExtend = nExt_m
result_m.nHuddle = nHud_m
result_m.nInitialize = nInit_m
result_m.nReflect = nRef_m

Wrapup:

' return the result
NMreduce = result_m

If traceLevel_m >= 1& Then
  Print #n_m, "NMreduce done"
  Print #n_m, "  Exit reason: " & NMresultString(result_m.reason)
  Print #n_m, "  Best point seen had value " & result_m.bestValue & " vars:"
  Print #n_m, arrayStr(result_m.bestVars, "  ")
  Print #n_m, "  Calls used: " & result_m.callsUsed
  Print #n_m, "  Final simplex size: " & result_m.finalSize
  Print #n_m, "  Final value range: " & result_m.finalRange
  Print #n_m, "  Initializations: " & result_m.nInitialize
  Print #n_m, "  Contract moves: " & result_m.nContract
  Print #n_m, "    Contract-out moves: "; result_m.nContractOut
  Print #n_m, "  Extend moves: " & result_m.nExtend
  Print #n_m, "  Reflect moves: " & result_m.nReflect
  Print #n_m, "  Huddle moves: " & result_m.nHuddle
  Print #n_m, "Elapsed time "; Format$(Timer() - tStart + _
    86400! * DateDiff("d", dStart, Date), "#0.000"); " seconds"
  Print #n_m, "<<<<<<<<<<<<<<<<<<<< Exit Function NMreduce at " & _
    Format$(Now(), "yyyy-mm-dd hh:mm:ss") & EOL
End If

' release dynamically-allocated memory
Erase genPoint_m.vars, refPoint_m.vars, scale_m, simplex_m, x_m

' turn off any trace action (if desired, it must be set before every call)
NMsetTraceLevel 0&
NMsetTraceFileNumber 0

Exit Function '*********** routine has just one exit (for debugging); this is it

ErrorHandler:  '----------------------------------------------------------------
' save properties of Err object; "On Error GoTo 0" erases them
errNum = Err.Number
errDsc = Err.Description
' augment the error info
On Error GoTo 0  ' avoid recursion
If InStr(errDsc, "Error in") > 0& Then  ' error from below here; add traceback
  errDsc = errDsc & vbLf & "called by " & ID_c & " call " & invoked_m
Else  ' error was in this routine (could add more info here)
  errDsc = errDsc & vbLf & "Error in " & ID_c & " call " & invoked_m
End If
Err.Raise errNum, ID_c, errDsc ' send error on up the call chain
End Function

'===============================================================================
Public Function NMresultString(ByRef why As Long) As String
Attribute NMresultString.VB_Description = "Return a string explaining the supplied result-code Enum."
Dim s As String
' by calling this with a return code, the caller can get a text explanation
If why = NMretValueTolMet Then
  s = "function-value tolerance met - result good"
ElseIf why = NMretSizeTolMet Then
  s = "simplex-size tolerance met - result good"
ElseIf why = NMretRangeTolMet Then
  s = "simplex-value-range tolerance met - result good"
ElseIf why = NMretTooManyCalls Then
  s = "too many function calls - result may be bad"
ElseIf why = NMretVarHuge Then
  s = "variable became huge - result may be bad"
ElseIf why = NMretTooFewVars Then
  s = "too few variables (need 2 or more) - result bad"
Else
  s = "unknown result status"
End If
NMresultString = s & " (code " & why & ")"
End Function

'===============================================================================
Public Function NMlimitExp( _
  ByVal x As Double, _
  ByVal xLo As Double, _
  ByVal xHi As Double, _
  Optional ByVal fillet As Double = 0.0001) _
As Double
' Smoothly limit x between xLo and xHi, with a transition fillet width near the
' value given by fillet * (xHi - xLo). Use an exponential-based method, which
' will rapidly approach x as x moves away from the limits. Use this to apply
' simple lower and upper ("box") limits to argument values in NMreduce. First,
' use NMlimitExp(x, xLo, xHi) instead of x in the coding of NMfunc. Then, run
' NMreduce. Make sure the initial argument values lie within the supplied
' limits. When NMreduce is done use NMlimitExp again, with the same limits, to
' convert the returned x to its limited value. For one-sided limits, just set
' the other side well away from the value you want to limit at (but not too far
' away, to avoid roundoff problems).
If fillet > 10# Then fillet = 10#  ' clip absurd values
Dim eps As Double, uLo As Double, uHi As Double
eps = Abs(fillet * (xHi - xLo))
If eps < 0.00000001 Then eps = 0.00000001  ' we are about to divide by eps
uLo = (x - xLo) / eps  ' use this as approx. to Log(1 + Exp(x)) for large x
If uLo < -18.42 Then  ' Exp(uLo) will be less than 1E-8
  uLo = Exp(uLo)  ' approx. Log(1 + Exp(x)) as Exp(x)
ElseIf uLo < 18.421 Then  ' Exp(uLo) will be less than 1E8
  uLo = Log(1# + Exp(uLo))  ' no need to approximate; use exact result
End If
uHi = (xHi - x) / eps
If uHi < -18.42 Then
  uHi = Exp(uHi)
ElseIf uHi < 18.421 Then
  uHi = Log(1# + Exp(uHi))
End If
NMlimitExp = xLo + xHi - x + eps * (uLo - uHi)
End Function

'===============================================================================
Public Function NMlimitHard( _
  ByVal x As Double, _
  ByVal xLo As Double, _
  ByVal xHi As Double) _
As Double
' Abruptly limit x between xLo and xHi, with no smooth fillet. Use this to apply
' simple lower and upper ("box") limits to argument values in NMreduce. First,
' use NMlimitHard(x, xLo, xHi) instead of x in the coding of NMfunc. Then, run
' NMreduce. Make sure the initial argument values lie within the supplied
' limits. When NMreduce is done use NMlimitHard again, with the same limits, to
' convert the returned x to its limited value. For one-sided limits, just set
' the other side well away from the value you want to limit.
If x < xLo Then
  NMlimitHard = xLo
ElseIf x > xHi Then
  NMlimitHard = xHi
Else
  NMlimitHard = x
End If
End Function

'===============================================================================
Public Function NMlimitSqr( _
  ByVal x As Double, _
  ByVal xLo As Double, _
  ByVal xHi As Double, _
  Optional ByVal fillet As Double = 0.0001) _
As Double
' Smoothly limit x between xLo and xHi, with a transition fillet width near the
' value given by fillet * (xHi - xLo). Use a square-root-based method, which
' will more slowly approach x as x moves away from the limits. Use this to apply
' simple lower and upper ("box") limits to argument values in NMreduce. First,
' use NMlimitSqr(x, xLo, xHi) instead of x in the coding of NMfunc. Then, run
' NMreduce. Make sure the initial argument values lie within the supplied
' limits. When NMreduce is done use NMlimitSqr again, with the same limits, to
' convert the returned x to its limited value. For one-sided limits, just set
' the other side well away from the value you want to limit at (but not too far
' away, to avoid roundoff problems).
If fillet > 10# Then fillet = 10#  ' clip absurd values
Dim dLo As Double, dHi As Double, eps As Double, eps2 As Double
dLo = x - xLo
dHi = xHi - x
eps = fillet * (xHi - xLo)
eps2 = eps * eps
NMlimitSqr = 0.5 * (xLo + xHi + Sqr(dLo * dLo + eps2) - Sqr(dHi * dHi + eps2))
End Function

'===============================================================================
Public Function NMgetFunctionIndex() As Long
' Return the index of the function that NMfunc will return. There is no need to
' use this if there is only one function in your NMfunc code. If there are
' several functions, do this:
'    Select Case NMgetFunctionIndex()
'    Case 0&
'      ' code for function #0 (default)
'    Case 1&
'      ' code for function #1
'    Case 2&
'      ' code for function #2
'    '
'    ' ... and so forth
'    '
'    Case Else
'      Err.Raise 5&, "NMfunc", "Error in NMfunc: function index = " & _
'        NMgetFunctionIndex() & " but function not coded." & vbNewLine & _
'        "Check the code in NMfunc for a missing Case value."
'    End Select
NMgetFunctionIndex = functionIndex_m
End Function

'===============================================================================
Public Sub NMsetFunctionIndex(ByVal newIndex As Long)
' Set the index of the function that NMfunc will return. There is no need to use
' this if there is only one function in your NMfunc code.
functionIndex_m = newIndex
End Sub

'===============================================================================
Public Sub NMsetParams( _
  ByVal newCon As Double, _
  ByVal newExt As Double, _
  ByVal newHud As Double, _
  ByVal newRef As Double, _
  Optional ByVal newReSize As Double = ReSizeDefault, _
  Optional ByVal newGrow As Double = GrowDefault, _
  Optional ByVal newCycleMul As Double = CycleMulDefault, _
  Optional ByVal newDropMul As Double = DropMulDefault)
' Set the algorithm tuning parameters that NMreduce uses. The performance of
' the routine depends strongly on these values, so change them very cautiously.
Con = newCon            ' contract distance
Ext = newExt            ' extend distance
Hud = newHud            ' huddle distance
Ref = newRef            ' reflect distance
ReSize = newReSize      ' amount size is grown at restart
Grow = newGrow          ' max amount restart can be larger than init
CycleMul = newCycleMul  ' this times var count is min calls/cycle
DropMul = newDropMul    ' this times var count is max no-drop

If n_m <> 0 Then
  Print #n_m, "User called NMsetParams to alter tuning parameters, so now"
  NMgetParams  ' prints parameters if n_m <> 0
End If

isInit_m = True
End Sub

'===============================================================================
Public Sub NMsetTraceFileNumber(ByVal newFileNumber As Integer)
n_m = newFileNumber
End Sub

'===============================================================================
Public Sub NMsetTraceLevel(ByVal newLevel As Long)
traceLevel_m = newLevel
End Sub

'===============================================================================
Public Function NMversion() As String
NMversion = Version_c
End Function

'###############################################################################
'#
'# Private (Module-only) support routines
'#
'###############################################################################

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub NMinit()
' set default return values, except for:
'   bestVars()
'   reason
result_m.bestValue = Huge
result_m.callsUsed = 0&
result_m.finalSize = Huge
result_m.finalRange = Huge
result_m.nContract = 0&
result_m.nContractOut = 0&
result_m.nExtend = 0&
result_m.nHuddle = 0&
result_m.nInitialize = 0&
result_m.nInvoked = invoked_m
result_m.nReflect = 0&
result_m.initialPtBest = False
result_m.refOrphanBest = False

' initialize counters
nCon_m = 0&
nConOut_m = 0&
nExt_m = 0&
nHud_m = 0&
nRef_m = 0&
nInit_m = 0&
End Sub

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub NMmove()
' set up error handling
Const ID_c As String = PN_c & "NMmove Sub"  ' name of this file & routine
Static calls_s As Double  ' number of times this routine has been called
Dim errNum As Long, errDsc As String  ' Err object Property holders
calls_s = calls_s + 1#    ' stops adding at 9,007,199,254,740,992 calls
If Not inDesign() Then On Error GoTo ErrorHandler

' find geometric centroid of all points except highest
Dim sum As Double
Dim j As Long
Dim k As Long
For j = LB_m To UB_m
  sum = 0#
  For k = 1& To nSimplex_m
    If k <> kTop_m Then sum = sum + simplex_m(k).vars(j)
  Next k
  genPoint_m.vars(j) = sum / nVars_m
Next j
If traceLevel_m >= 3& Then
  Print #n_m, "  Centroid is at vars:"
  For j = LB_m To UB_m
    x_m(j) = genPoint_m.vars(j) * scale_m(j)
  Next j
  Print #n_m, arrayStr(x_m, "   ")
End If

' reflect highest point through centroid
For j = LB_m To UB_m
  refPoint_m.vars(j) = (1# + Ref) * genPoint_m.vars(j) - _
                       Ref * simplex_m(kTop_m).vars(j)
  x_m(j) = refPoint_m.vars(j) * scale_m(j)
Next j
'-------------------- function call --------------------
refPoint_m.val = NMfunc(x_m)
'-------------------------------------------------------
nCalls_m = nCalls_m + 1&
If traceLevel_m >= 3& Then
  Print #n_m, "  Reflect value " & refPoint_m.val & "  calls " & nCalls_m & _
    "  vars:"
  Print #n_m, arrayStr(x_m, "   ")
End If
nCycle_m = nCycle_m + 1&
nRef_m = nRef_m + 1&

If refPoint_m.val < fBot_m Then
  ' reflected point is lowest - good news
  ' extend farther in same direction (overlay centroid)
  For j = LB_m To UB_m
    genPoint_m.vars(j) = (1# + Ext) * genPoint_m.vars(j) - _
                         Ext * simplex_m(kTop_m).vars(j)
    x_m(j) = genPoint_m.vars(j) * scale_m(j)
  Next j
  '-------------------- function call --------------------
  genPoint_m.val = NMfunc(x_m)
  '-------------------------------------------------------
  nCalls_m = nCalls_m + 1&
  If traceLevel_m >= 3& Then
    Print #n_m, "  Reflect is lowest  extend value " & genPoint_m.val & _
      "  calls " & nCalls_m & "  vars:"
    Print #n_m, arrayStr(x_m, "   ")
  End If
  nCycle_m = nCycle_m + 1&
  nExt_m = nExt_m + 1&

  If genPoint_m.val < fBot_m Then
    ' extended below lowest - replace top with extended
    ' do this even if reflected point was lower yet - it pays off later
    simplex_m(kTop_m) = genPoint_m
    If traceLevel_m >= 3& Then
      Print #n_m, "  Extend replaces worst"
    End If
    If refPoint_m.val < genPoint_m.val Then  ' reflected point was lower yet
      If traceLevel_m >= 3& Then
        Print #n_m, "  ... even though reflect was lower than extend"
      End If
      ' we won't use this point in the simplex, but it might be the best yet
      ' seen, so check best-seen values and update if necessary
      If funcMinCycle_m > refPoint_m.val Then  ' it's the best seen this cycle
        funcMinCycle_m = refPoint_m.val
        If result_m.bestValue > refPoint_m.val Then  ' it's the best ever seen
          result_m.bestValue = refPoint_m.val
          For j = LB_m To UB_m
            result_m.bestVars(j) = refPoint_m.vars(j) * scale_m(j)
          Next j
          If traceLevel_m >= 3& Then
            Print #n_m, "  Best point now non-simplex reflect  value " & _
              refPoint_m.val
          End If
          result_m.refOrphanBest = True  ' report unusual result to user
        End If
      End If
    End If
  Else  ' extended above lowest, reflected below, so replace top with reflected
    simplex_m(kTop_m) = refPoint_m
    If traceLevel_m >= 3& Then
      Print #n_m, "  Extend above lowest; reflect replaces worst"
    End If
  End If
ElseIf refPoint_m.val < simplex_m(kMid_m).val Then
' reflected point below next-highest - replace top with reflected
  simplex_m(kTop_m) = refPoint_m
  If traceLevel_m >= 3& Then
    Print #n_m, "  Reflect above lowest; reflect replaces worst"
  End If
Else  ' ref >= second-highest
' reflected point above second-highest - this looks bad
  If refPoint_m.val < simplex_m(kTop_m).val Then  ' reflected below top; replace
    simplex_m(kTop_m) = refPoint_m
    nConOut_m = nConOut_m + 1&
    If traceLevel_m >= 3& Then
      Print #n_m, "  Reflect above next worst; reflect replaces worst"
    End If
  End If
  ' contract point toward centroid (overlay centroid)
  For j = LB_m To UB_m
    genPoint_m.vars(j) = (1# - Con) * genPoint_m.vars(j) + _
                         Con * simplex_m(kTop_m).vars(j)
    x_m(j) = genPoint_m.vars(j) * scale_m(j)
  Next j
  '-------------------- function call --------------------
  genPoint_m.val = NMfunc(x_m)
  '-------------------------------------------------------
  nCalls_m = nCalls_m + 1&
  If traceLevel_m >= 3& Then
    Print #n_m, "  Contract value " & genPoint_m.val & _
      "  calls " & nCalls_m & "  vars:"
    Print #n_m, arrayStr(x_m, "   ")
  End If
  nCycle_m = nCycle_m + 1&
  nCon_m = nCon_m + 1&
  If genPoint_m.val < simplex_m(kTop_m).val Then
    simplex_m(kTop_m) = genPoint_m
    If traceLevel_m >= 3& Then
      Print #n_m, "  Contract below worst and replaces it"
    End If
  Else  ' no point was below highest - huddle in panic around lowest
    If traceLevel_m >= 3& Then
      Print #n_m, _
        "  Reflect > next-worst & contract > worst; huddle toward lowest"
    End If
    Dim temp As Double
    For k = 1& To nSimplex_m
      If k <> kBot_m Then
        For j = LB_m To UB_m
          temp = (1# - Hud) * simplex_m(kBot_m).vars(j) + _
            Hud * simplex_m(k).vars(j)
          simplex_m(k).vars(j) = temp
          x_m(j) = temp * scale_m(j)
        Next j
        '-------------------- function call --------------------
        simplex_m(k).val = NMfunc(x_m)
        '-------------------------------------------------------
        nCalls_m = nCalls_m + 1&
        If traceLevel_m >= 3& Then
          Print #n_m, "  Huddle value " & simplex_m(k).val & _
            "  calls " & nCalls_m & "  vars:"
          Print #n_m, arrayStr(x_m, "   ")
        End If
      End If
    Next k
    nCycle_m = nCycle_m + nVars_m
    nHud_m = nHud_m + 1&
  End If
End If
Exit Sub '**************** routine has just one exit (for debugging); this is it

ErrorHandler:  '----------------------------------------------------------------
' save properties of Err object; "On Error GoTo 0" erases them
errNum = Err.Number
errDsc = Err.Description
' augment the error info
On Error GoTo 0  ' avoid recursion
If InStr(errDsc, "Error in") > 0& Then  ' error from below here; add traceback
  errDsc = errDsc & vbLf & "called by " & ID_c & " call " & calls_s
Else  ' error was in this routine (could add more info here)
  errDsc = errDsc & vbLf & "Error in " & ID_c & " call " & calls_s
End If
Err.Raise errNum, ID_c, errDsc ' send error on up the call chain
End Sub

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub NMsort()
' find the largest, next-largest and smallest function values in the simplex
' set up error handling
Const ID_c As String = PN_c & "NMsort Sub"  ' name of this file & routine
Static calls_s As Double  ' number of times this routine has been called
Dim errNum As Long, errDsc As String  ' Err object Property holders
calls_s = calls_s + 1#    ' stops adding at 9,007,199,254,740,992 calls
If Not inDesign() Then On Error GoTo ErrorHandler

' set order of first two points
If simplex_m(1&).val < simplex_m(2&).val Then
  kBot_m = 1&
  kTop_m = 2&
Else
  kBot_m = 2&
  kTop_m = 1&
End If

' shuffle third point into place
If simplex_m(3&).val < simplex_m(kBot_m).val Then  ' below bottom
  kMid_m = kBot_m
  kBot_m = 3&
ElseIf simplex_m(3&).val > simplex_m(kTop_m).val Then  ' above top
  kMid_m = kTop_m
  kTop_m = 3&
Else  ' must be between other two (or equal)
  kMid_m = 3&
End If

' adjust ranking with remaining points (if any)
Dim j As Long
Dim temp As Double
For j = 4& To nSimplex_m
  temp = simplex_m(j).val
  If temp < simplex_m(kBot_m).val Then
    kBot_m = j
  ElseIf temp > simplex_m(kTop_m).val Then
    kMid_m = kTop_m
    kTop_m = j
  ElseIf temp > simplex_m(kMid_m).val Then
    kMid_m = j
  End If
Next j

' carry out tests to see if we are making progress (function is decreasing)
If fBot_m <= simplex_m(kBot_m).val Then  ' we made no progress
  noDrop_m = noDrop_m + 1&  ' increase level of despair
Else  ' the function has decreased
  noDrop_m = 0&  ' recover our good spirits
End If
' reset value for next did-function-decrease test, and for use in move logic
fBot_m = simplex_m(kBot_m).val

' keep track of best value seen this cycle, and best overall
If funcMinCycle_m > fBot_m Then
  funcMinCycle_m = fBot_m
  ' update best-so-far point
  If result_m.bestValue > fBot_m Then
    result_m.bestValue = fBot_m
    For j = LB_m To UB_m  ' save in caller units
      result_m.bestVars(j) = simplex_m(kBot_m).vars(j) * scale_m(j)
    Next j
    result_m.refOrphanBest = False  ' best is not an unused reflected point
    If traceLevel_m >= 2& Then
      Print #n_m, "Best point updated to simplex point " & kBot_m
    End If
  End If
End If

If traceLevel_m >= 3& Then
  Print #n_m, "  Simplex values sorted to:"
  Print #n_m, "    highest: point " & kTop_m & " value " & simplex_m(kTop_m).val
  Print #n_m, "    next-hi: point " & kMid_m & " value " & simplex_m(kMid_m).val
  Print #n_m, "    lowest:  point " & kBot_m & " value " & simplex_m(kBot_m).val
  Print #n_m, "    function-didn't-decrease count: " & noDrop_m
End If
Exit Sub '**************** routine has just one exit (for debugging); this is it

ErrorHandler:  '----------------------------------------------------------------
' save properties of Err object; "On Error GoTo 0" erases them
errNum = Err.Number
errDsc = Err.Description
' augment the error info
On Error GoTo 0  ' avoid recursion
If InStr(errDsc, "Error in") > 0& Then  ' error from below here; add traceback
  errDsc = errDsc & vbLf & "called by " & ID_c & " call " & calls_s
Else  ' error was in this routine (could add more info here)
  errDsc = errDsc & vbLf & "Error in " & ID_c & " call " & calls_s
End If
Err.Raise errNum, ID_c, errDsc ' send error on up the call chain
End Sub

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub NMstart()
' start or restart a simplex minimization cycle
' set up error handling
Const ID_c As String = PN_c & "NMstart Sub"  ' name of this file & routine
Static calls_s As Double  ' number of times this routine has been called
Dim errNum As Long, errDsc As String  ' Err object Property holders
calls_s = calls_s + 1#    ' stops adding at 9,007,199,254,740,992 calls
If Not inDesign() Then On Error GoTo ErrorHandler

' save present values for use as previous values, in case a (re)start gives worse
prevPoint_m.val = result_m.bestValue
prevPoint_m.vars = result_m.bestVars
prevReason_m = result_m.reason

nInit_m = nInit_m + 1&  ' count starts & restarts

noDrop_m = 0&  ' initialize failed-to-decrease counter
fBot_m = result_m.bestValue  ' test value to see if function decreased

funcMinCycle_m = Huge  ' best value seen in this cycle

' get normalization factors for variables that make simplex coordinates near 1.0
scale_m = result_m.bestVars
' ...but avoid scaling with a variable that's equal to zero
Dim j As Long
For j = LB_m To UB_m
  If scale_m(j) = 0# Then scale_m(j) = 1#
Next j

' make regular simplex with circumsphere diameter "sizeStart_m" with its
' centroid at best known point (that's the initial point on the first call)
' note that no simplex point will coincide with the best known point
' the method used here relies on two facts for a regular unit-radius simplex:
'   1) all vectors from centroid to simplex points have same length of unity
'   2) dot product of any such vector with any other is -1 / nVars_m
' note that we could partially randomize this process by doing the points in
' random order, and by randomly choosing the sign of the square root in the
' length-normalization step (except for the last two points, which would just
' swap places on a sign change)

' first, make a regular simplex with unit-length radius-from-origin vectors
Dim jj As Long, k As Long
Dim avgVar As Double, sum As Double, temp As Double
avgVar = 0#
For j = 0& To nVars_m - 1&  ' do the coordinates in order
  For k = 1& To nSimplex_m
    If k <= j Then  ' coordinate values above diagonal = 0
      simplex_m(k).vars(LB_m + j) = 0#
    ElseIf k = j + 1& Then  ' this is the diagonal element
      sum = 0#  ' accumulate unity, less the sum of squared coordinates
      For jj = 0& To j - 1&
        temp = simplex_m(k).vars(LB_m + jj)
        sum = sum + temp * temp
      Next jj
      temp = Sqr(1# - sum)  ' will want sum and Sqr(1# - sum) for next k
      simplex_m(k).vars(LB_m + j) = temp  ' sign choice is arbitrary
    ElseIf k = j + 2& Then ' coordinate values below diagonal via dot product
      simplex_m(k).vars(LB_m + j) = -(1# / nVars_m + sum) / temp
    Else
      ' copy value on down simplex points
      simplex_m(k).vars(LB_m + j) = simplex_m(k - 1&).vars(LB_m + j)
    End If
  Next k
  ' get an approximate measure of the present simplex-coordinate size
  avgVar = avgVar + Abs(result_m.bestVars(LB_m + j))
Next j
avgVar = avgVar / nVars_m

' second, translate and scale the unit regular simplex to where we want it
' make sure the points are not seriously influenced by roundoff
Dim sizeMin As Double
sizeMin = 16# * EpsDbl * avgVar
If sizeStart_m < sizeMin Then sizeStart_m = sizeMin
For k = 1& To nSimplex_m
  For j = LB_m To UB_m
    If result_m.bestVars(j) <> 0# Then
      simplex_m(k).vars(j) = 1# + sizeStart_m * simplex_m(k).vars(j)
    Else
      simplex_m(k).vars(j) = sizeStart_m * simplex_m(k).vars(j)
    End If
  Next j
Next k

' evaluate function at vertices of simplex
For k = 1& To nSimplex_m
  For j = LB_m To UB_m
    x_m(j) = simplex_m(k).vars(j) * scale_m(j)  ' back in caller's units
  Next j
  '-------------------- function call --------------------
  simplex_m(k).val = NMfunc(x_m)
  '-------------------------------------------------------
Next k
nCalls_m = nCalls_m + nSimplex_m
nCycle_m = nSimplex_m ' count calls for this restart cycle

If traceLevel_m >= 1& Then
  Print #n_m, "Initialization " & nInit_m & "  simplex size " & _
    CSng(sizeStart_m)
  For k = 1 To nSimplex_m
    Print #n_m, "  point " & k & "  value = " & simplex_m(k).val & "  vars:"
    For j = LB_m To UB_m
      x_m(j) = simplex_m(k).vars(j) * scale_m(j)
    Next j
    Print #n_m, arrayStr(x_m, "   ")
  Next k
End If
Exit Sub '**************** routine has just one exit (for debugging); this is it

ErrorHandler:  '----------------------------------------------------------------
' save properties of Err object; "On Error GoTo 0" erases them
errNum = Err.Number
errDsc = Err.Description
' augment the error info
On Error GoTo 0  ' avoid recursion
If InStr(errDsc, "Error in") > 0& Then  ' error from below here; add traceback
  errDsc = errDsc & vbLf & "called by " & ID_c & " call " & calls_s
Else  ' error was in this routine (could add more info here)
  errDsc = errDsc & vbLf & "Error in " & ID_c & " call " & calls_s
End If
Err.Raise errNum, ID_c, errDsc ' send error on up the call chain
End Sub

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub NMtest()
' this routine determines if any exit criterion has been met
' it reports its results by setting testResult_m to a return code

' find largest variable & normalized distance from highest point to lowest point
Dim ab As Double
Dim big As Double
Dim jBig As Long
Dim sizSqrd As Double
big = Abs(simplex_m(kBot_m).vars(LB_m) * scale_m(LB_m))
jBig = LB_m
sizSqrd = (simplex_m(kTop_m).vars(LB_m) - simplex_m(kBot_m).vars(LB_m)) ^ 2
Dim j As Long
For j = LB_m + 1& To UB_m
  ab = Abs(simplex_m(kBot_m).vars(j) * scale_m(j))
  If big < ab Then  ' keep track of largest variable (in user's units)
    big = ab
    jBig = j
  End If
  sizSqrd = sizSqrd + _
    (simplex_m(kTop_m).vars(j) - simplex_m(kBot_m).vars(j)) ^ 2
Next j
size_m = Sqr(sizSqrd)

' set simplex scale for possible later minimization cycle
sizeStart_m = ReSize * size_m
' but don't use a value larger than caller's specified initial value
If sizeStart_m > Grow * sizeInit_m Then sizeStart_m = Grow * sizeInit_m
' ...and don't let it get comparable to the tolerance
If sizeStart_m < 2# * rangeTol_m Then sizeStart_m = 2# * rangeTol_m

' find relative value difference between highest and lowest points
Dim temp As Double
' handle case where simplex values straddle zero
ab = Abs(simplex_m(kTop_m).val)
If Abs(fBot_m) < ab Then
  temp = Abs(fBot_m)
  range_m = ab
Else
  temp = ab
  range_m = Abs(fBot_m)
End If
If temp < 1024# * Tiny Then temp = 1024# * Tiny
range_m = range_m / temp - 1#

' test for exit conditions; put result in testResult_m
testResult_m = NMretNoReason  ' flag value; not a legal return code value
If fBot_m <= valueTol_m Then
  testResult_m = NMretValueTolMet
  If traceLevel_m >= 1& Then
    Print #n_m, "Exit criterion met: function value = "; fBot_m
  End If
ElseIf nCalls_m >= maxCalls_m Then
  testResult_m = NMretTooManyCalls
  If traceLevel_m >= 1& Then
    Print #n_m, "Exit criterion met: function-call count = "; nCalls_m
  End If
' do not do a size exit unless enough calls have been made
ElseIf (size_m <= sizeTol_m) And (nCycle_m >= CycleMul * nVars_m) Then
  testResult_m = NMretSizeTolMet
  If traceLevel_m >= 1& Then
    Print #n_m, "Exit criterion met: simplex size = "; size_m; _
      " and cycle calls = "; nCycle_m
  End If
' do not do a range exit unless enough calls have been made
ElseIf (range_m <= rangeTol_m) And (nCycle_m >= CycleMul * nVars_m) Then
  testResult_m = NMretRangeTolMet
  If traceLevel_m >= 1& Then
    Print #n_m, "Exit criterion met: simplex value range = "; range_m; _
      " and cycle calls = "; nCycle_m
  End If
ElseIf big > Huge / 1024# Then
  testResult_m = NMretVarHuge
  If traceLevel_m >= 1& Then
    Print #n_m, "Exit criterion met: huge variable = "; big
  End If
End If

If traceLevel_m >= 2& Then  ' do basic debug printout
  Print #n_m, " Calls "; nCalls_m; "  size "; CSng(size_m); _
    "  range "; CSng(range_m); "  best "; _
    result_m.bestValue; "  vars:"
  For j = LB_m To UB_m
    x_m(j) = result_m.bestVars(j)
  Next j
  Print #n_m, arrayStr(x_m, " ")
End If

If testResult_m = NMretNoReason Then  ' no exit condition was met
  ' if function value has not decreased for "many" calls, force a restart
  If noDrop_m > DropMul * nVars_m Then
    testResult_m = NMretNoDecrease  ' flag value; not a real return code value
  End If
End If
End Sub

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Function arrayStr( _
  ByRef arrayOfValues As Variant, _
  Optional ByVal padding As String = "", _
  Optional ByVal lineLength As Long = 80&) _
As String
' Converts array values to string consisting of "lines" with linefeeds between
' them to keep line length to less than specified value, after adding
' "padding" to front of each line. Lines contain "index=value(index)" entries.
' There will always be at least one entry per line, no matter what
' "lineLength" is. The input array can contain anything that can be converted
' to a string, including different types if it is an array of Variants.
' This routine is used in trace-file printing (if trace is on).
'         John Trenholme - 2003-08-19
Dim j As Long, jFirst As Long, jLast As Long
Dim sAdd As String, sLine As String, sNow As String
If (VarType(arrayOfValues) And vbArray) = 0 Then  ' input is a scalar item
  arrayStr = padding & "S=" & arrayOfValues
Else
  jFirst = LBound(arrayOfValues)
  jLast = UBound(arrayOfValues)
  sNow = ""  ' start with empty result
  For j = jFirst To jLast
    sAdd = j & "=" & arrayOfValues(j)  ' will add "index=value(index)" entry
    If j < jLast Then sAdd = sAdd & ","  ' separator for all but last entry
    If j = jFirst Then  ' do special case
      sLine = padding & sAdd  ' make first line, with entry
    ElseIf Len(sLine) + Len(sAdd) + 1& >= lineLength Then  ' it won't fit
      sNow = sNow & sLine & EOL  ' spill line onto output string
      sLine = padding & sAdd  ' start new line, with entry
    Else  ' it will fit
      sLine = sLine & " " & sAdd  ' add entry to line, with separation space
    End If
  Next j
  sNow = sNow & sLine  ' add last line to result
  arrayStr = sNow
End If
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub checkInitialization()
' If the algorithm tuning parameters are not initialized, set them to defaults.
If Not isInit_m Then
  Con = ConDefault            ' contract distance
  Ext = ExtDefault            ' extend distance
  Hud = HudDefault            ' huddle distance
  Ref = RefDefault            ' reflect distance
  ReSize = ReSizeDefault      ' amount size is grown at restart
  Grow = GrowDefault          ' max amount restart can be larger than init
  CycleMul = CycleMulDefault  ' this times var count is min calls/cycle
  DropMul = DropMulDefault    ' this times var count is max no-drop

  isInit_m = True
  
  If traceLevel_m >= 3& Then
    Print #n_m, "   N-M algorithm tuning parameters set to default values"
    Print #n_m, "     Con      = "; Con
    Print #n_m, "     Ext      = "; Ext
    Print #n_m, "     Hud      = "; Hud
    Print #n_m, "     Ref      = "; Ref
    Print #n_m, "     ReSize   = "; ReSize
    Print #n_m, "     Grow     = "; Grow
    Print #n_m, "     CycleMul = "; CycleMul
    Print #n_m, "     DropMul  = "; DropMul
  End If
End If
End Sub

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Function inDesign() As Boolean
' Returns True if program is running in IDE (editor) design environment, and
' False if program is running as a standalone EXE. Of course, VBA is always in
' the IDE, so you always get True.
'         John Trenholme - 2009-10-21
On Error Resume Next
Debug.Assert 1 / 0  ' attempts this illegal feat only in IDE
inDesign = (Err.Number <> 0&)
'inDesign = False  ' uncomment this to get compiled behavior while in IDE
Err.Clear  ' do not pass 1/0 error back up to caller
End Function

'----------------------------- end of file -------------------------------------


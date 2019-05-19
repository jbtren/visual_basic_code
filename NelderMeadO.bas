Attribute VB_Name = "NelderMead"
'
'###############################################################################
'#
'# Visual Basic 6 source file "NelderMead.bas"
'#
'# John Trenholme - initial version 13 Aug 2003
'#
'# This routine reduces the value of an unconstrained multidimensional function
'# by adjusting an array of two or more Double variables. The function to be
'# reduced is supplied as a "function object" so that information not contained
'# in the variable array (so-called "side information") can be passed from the
'# caller to the function without this routine seeing it, and so this routine
'# can be used to reduce a variety of functions by supplying different classes
'# that implement the I_Function interface. See Class I_Function for the
'# "function object" interface and its uses.
'#
'# To minimize a function of one variable, use "BrentMin."
'#
'###############################################################################

Option Base 0
Option Compare Binary
Option Explicit
Option Private Module  ' no effect in Visual Basic; globals project-only in VBA

' Enum values return status when reduction is complete.
Public Enum NMreturnEnum
  NMretvalueToleranceMet = 1&  ' function value dropped below threshold
  NMretSizeToleranceMet        ' simplex size dropped below limit
  NMretRangeToleranceMet       ' range of values in simplex dropped below limit
  NMretTooManyCalls            ' number of allowed function calls was exceeded
  NMretVariableHuge            ' variable(s) approached max size of Doubles
  NMretTooFewVariables         ' the supplied function had 2 or fewer arguments
End Enum

' Type holding results of the reduction; "reduceNM" returns one of these.
Public Type NMresultType
  bestValue As Double    ' smallest value seen during reduction
  bestVars() As Double   ' variable values at the point of smallest value
  reason As NMreturnEnum ' cause of return from the function
  finalSize As Double    ' final size of the Nelder-Mead simplex
  finalRange As Double   ' final value range in the Nelder-Mead simplex
  callsUsed As Long      ' number of function calls made during reduction
  nVars As Long          ' number of variables varied during reduction
  nReflect As Long       ' number of "reduction" moves made during reduction
  nExtend As Long        ' number of "extension" moves made during reduction
  nContract As Long      ' number of "contraction" moves made during reduction
  nHuddle As Long        ' number of "huddle" moves made during reduction
  nInitialize As Long    ' number of initializations made during reduction
  nInvoked As Long       ' number of times "reduceNM" function has been called
End Type

' %%%%%%%%%%%%%%%%%%%%%%%%%%%% Private %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

Private Const FileName As String = "NelderMead"  ' ID for this file
' version (date) of this file
Public Const NelderMeadVersion As String = "2007-09-28"

' ***** Algorithm-tuning parameters - adjust with care
Private Const Con As Double = 0.5  ' contract distance
Private Const Ext As Double = 2.5  ' extend distance
Private Const Hud As Double = 0.5  ' huddle distance
Private Const Ref As Double = 2.5  ' reflect distance
Private Const ReSize As Double = 1.5   ' amount size is grown at restart
Private Const Grow As Double = 1.5  ' max amount restart can be larger than init
Private Const CycleMul As Double = 5#  ' this times var count is min calls/cycle
Private Const DropMul As Double = 3#  ' this times var count is max no-drop

' ***** Common factors
Private Const ConCmp As Double = 1# - Con
Private Const ExtCmp As Double = 1# - Ext
Private Const HudCmp As Double = 1# - Hud
Private Const RefCmp As Double = 1# - Ref

Private Const Huge As Double = 1.797693E+308  ' 2.0 ^ 1023.999..
Private Const Tiny As Double = 5.562685E-309  ' 2.0 ^ (-1024.0)

Private Const NoDecrease As Long = -1&
Private Const NoReason As Long = 0&

' ***** Argument values and function value at a point
Private Type NMpointType
  vars() As Double
  val As Double
End Type

' ***** The module keeps information in module-global variables
Private m_bestPoint As NMpointType  ' best point we have ever seen in this run
Private m_bestValPrev As Double
Private m_fBot As Double
Private m_genPoint As NMpointType  ' point used for several purposes
Private m_invoked As Long  ' number of times this routine has been invoked
Private m_kBot As Long  ' index of lowest point in simplex
Private m_kMid As Long  ' index of next-to-highest point in simplex
Private m_kTop As Long  ' index of highest point in simplex
Private m_LB As Long  ' lower bound of arrays
Private m_maxCalls As Long
Private m_nCalls As Long  ' total number of calls during this invocation
Private m_nCon As Long
Private m_nCycle As Long  ' number of calls during this cycle
Private m_nExt As Long
Private m_nHud As Long
Private m_nInit As Long
Private m_noDrop As Long
Private m_nRef As Long
Private m_nSimplex As Long  ' number of points in simplex
Private m_nVary As Long  ' number of variables we are varying
Private m_range As Double  ' present value range in the simplex
Private m_rangeTol As Double  ' caller's exit tolerance on relative value diff.
Private m_refPoint As NMpointType  ' point we reflect to
Private m_result As NMresultType   ' this holds the return information
Private m_scale() As Double  ' size of variables, used for scaling to unit size
Private m_simplex() As NMpointType  ' the simplex of N+1 points
Private m_size As Double  ' size of the simplex
Private m_sizeStart As Double  ' present-cycle start size of the simplex
Private m_sizeInit As Double  ' caller's initial size of the simplex
Private m_sizeTol As Double  ' caller's exit tolerance on normalized size
Private m_UB As Long  ' upper bound of arrays
Private m_x() As Double  ' temporary variable array
Private m_valueTol As Double  ' caller's exit tolerance on function value

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% Debug %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

' ***** If DebugLevel is greater than 0, log-to-file statements will execute
#Const DebugLevel = 0&

#If DebugLevel >= 1& Then
  Const DBfileName As String = "NelderMeadLog"  ' "NNN.txt" will be added
  Dim m_ofn As Integer  ' output file number
#End If

' %%%%%%%%%%%%%%%%%%%%%%%%% Routines %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

' ******************************************************************************
Public Function reduceNM( _
  ByRef fun As I_Function, _
  ByRef initialPosition() As Double, _
  ByVal initialSize As Double, _
  ByVal valueTolerance As Double, _
  ByVal sizeTolerance As Double, _
  ByVal rangeTolerance As Double, _
  ByVal callLimit As Long) _
As NMresultType
' This routine carries out the reduction of the function.

' localize input variables
m_sizeInit = initialSize
m_valueTol = valueTolerance
m_sizeTol = sizeTolerance
m_rangeTol = rangeTolerance
m_maxCalls = callLimit

' localize array limits (assumes initialPosition() has it right)
m_LB = LBound(initialPosition)
m_UB = UBound(initialPosition)

m_sizeStart = m_sizeInit  ' present size of the simplex
m_nVary = m_UB - m_LB + 1& ' number of variables to adjust
m_nSimplex = m_nVary + 1&  ' number of points in the simplex

reduceNMinit

' quit early if user has asked for variation of fewer than 2 variables
If m_nVary < 2 Then
  ReDim m_result.bestVars(0 To 0)
  m_result.bestVars(0) = Huge  ' dummy return value
  m_result.reason = NMretTooFewVariables
  reduceNM = m_result
  Exit Function
End If

' set array sizes (dynamic memory allocation)
ReDim m_scale(m_LB To m_UB)
ReDim m_simplex(1& To m_nSimplex)
Dim j As Long
For j = LBound(m_simplex) To UBound(m_simplex)
  ReDim m_simplex(j).vars(m_LB To m_UB)
Next j
ReDim m_x(m_LB To m_UB)
ReDim m_genPoint.vars(m_LB To m_UB)
ReDim m_refPoint.vars(m_LB To m_UB)

' set best-point-so-far to input point (note variables are not yet normalized)
m_bestPoint.vars = initialPosition  ' ReDim array (allocate mem.) & set values

' do a function call at the initial point (it might turn out to be best)
m_bestPoint.val = fun.Value(m_bestPoint.vars)
m_nCalls = 1&

#If DebugLevel >= 1& Then
  InitialLog fun
#End If

' inner loops of the algorithm - steps are in separate routines for clarity
Dim mustExit As Boolean
Dim reStart As Boolean
Do  ' simplex restart loop
  reduceNMstart fun
  Do  ' simplex test-move loop
    reduceNMsort
    reduceNMtest
    ' see if value or size tolerance was met
    reStart = False
    If (m_result.reason = NMretRangeToleranceMet) Or _
       (m_result.reason = NMretSizeToleranceMet) Then
      mustExit = False  ' stay in simplex restart loop until no decrease
      Exit Do
    ' see if a quit-right-now condition was seen
    ElseIf (m_result.reason = NMretvalueToleranceMet) Or _
           (m_result.reason = NMretTooManyCalls) Or _
           (m_result.reason = NMretVariableHuge) Then
      mustExit = True  ' exit restart loop immediately
      Exit Do
    ElseIf m_result.reason = NoDecrease Then  ' no function drop in many tries
      mustExit = False  ' stay in simplex restart loop
      reStart = True  ' force a simplex restart
      Exit Do
    End If
    reduceNMmove fun
  Loop  ' end of test-move loop
  
  #If DebugLevel >= 1& Then
    Print #m_ofn, "Restart cycle " & m_nInit & " complete"
    Dim k As Long
    For k = 1 To m_nSimplex
      Print #m_ofn, "  point " & k & "  value = " & m_simplex(k).val & _
        "  caller variables:"
      For j = m_LB To m_UB
        m_x(j) = m_simplex(k).vars(j) * m_scale(j)
      Next j
      Print #m_ofn, arrayStr(m_x, "    ")
    Next k
  #End If
  
  If mustExit Then Exit Do
Loop While (m_bestPoint.val < m_bestValPrev) Or reStart  ' end of restart loop

' we are done, so set return values (if not already set)
m_result.bestValue = m_bestPoint.val
m_result.bestVars = m_bestPoint.vars
m_result.callsUsed = m_nCalls
m_result.finalSize = m_size
m_result.finalRange = m_range
m_result.nContract = m_nCon
m_result.nExtend = m_nExt
m_result.nHuddle = m_nHud
m_result.nInitialize = m_nInit
m_result.nReflect = m_nRef
m_result.nVars = m_nVary

' return the result
reduceNM = m_result

#If DebugLevel >= 1& Then
  FinalLog
#End If
End Function

' ******************************************************************************
Public Function NMresultString(ByRef why As NMreturnEnum) As String
Dim s As String
' by calling this with a return code, the caller can get a text explanation
If why = NMretvalueToleranceMet Then
  s = "function-value tolerance met - result good"
ElseIf why = NMretRangeToleranceMet Then
  s = "simplex-value-range tolerance met - result good"
ElseIf why = NMretSizeToleranceMet Then
  s = "simplex-size tolerance met - result good"
ElseIf why = NMretTooManyCalls Then
  s = "too many function calls - result may be bad"
ElseIf why = NMretVariableHuge Then
  s = "variable became huge - result may be bad"
ElseIf why = NMretTooFewVariables Then
  s = "too few variables (need 2 or more) - result bad"
Else
  s = "unknown result status"
End If
NMresultString = s & " (code " & why & ")"
End Function

' ==============================================================================
' Module-only support routines
' ==============================================================================

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub reduceNMinit()
m_invoked = m_invoked + 1&  ' bump NM-function-called counter

' set default return values
m_result.bestValue = Huge
m_result.callsUsed = 0&
m_result.finalSize = Huge
m_result.finalRange = Huge
m_result.nContract = 0&
m_result.nExtend = 0&
m_result.nHuddle = 0&
m_result.nInitialize = 0&
m_result.nInvoked = m_invoked
m_result.nReflect = 0&

' initialize counters
m_nCon = 0&
m_nExt = 0&
m_nHud = 0&
m_nRef = 0&
m_nInit = 0&
End Sub

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub reduceNMmove(ByRef f As I_Function)
' find geometric centroid of all points except highest
Dim sum As Double
Dim j As Long
Dim k As Long
For j = m_LB To m_UB
  sum = m_simplex(1).vars(j)
  For k = 2 To m_nSimplex
    sum = sum + m_simplex(k).vars(j)
  Next k
  m_genPoint.vars(j) = (sum - m_simplex(m_kTop).vars(j)) / m_nVary
Next j

' reflect highest point through centroid
For j = m_LB To m_UB
  m_refPoint.vars(j) = Ref * m_genPoint.vars(j) + _
                       RefCmp * m_simplex(m_kTop).vars(j)
  m_x(j) = m_refPoint.vars(j) * m_scale(j)
Next j
m_refPoint.val = f.Value(m_x)
m_nCalls = m_nCalls + 1&
m_nCycle = m_nCycle + 1&
m_nRef = m_nRef + 1&

If m_refPoint.val < m_fBot Then
' reflected point is lowest - good news
  ' extend farther in same direction (overlay centroid)
  For j = m_LB To m_UB
    m_genPoint.vars(j) = Ext * m_genPoint.vars(j) + _
                         ExtCmp * m_simplex(m_kTop).vars(j)
    m_x(j) = m_genPoint.vars(j) * m_scale(j)
  Next j
  m_genPoint.val = f.Value(m_x)
  m_nCalls = m_nCalls + 1&
  m_nCycle = m_nCycle + 1&
  m_nExt = m_nExt + 1&
  If m_genPoint.val < m_fBot Then
    ' extended below lowest - replace top with extended
    ' do this even if reflected point was lower yet - it pays off later
    m_simplex(m_kTop) = m_genPoint
    If m_refPoint.val < m_genPoint.val Then  ' reflected point was lower yet
      If m_refPoint.val <= m_bestPoint.val Then  ' ...and lower than best
        ' updte best-so-far point
        m_bestPoint.val = m_refPoint.val
        For j = m_LB To m_UB
          m_bestPoint.vars(j) = m_refPoint.vars(j) * m_scale(j)
        Next j
      End If
    End If
  Else  ' extended above lowest, reflected below, so replace top with reflected
    m_simplex(m_kTop) = m_refPoint
  End If
ElseIf m_refPoint.val < m_simplex(m_kMid).val Then
' reflected point below next-highest - replace top with reflected
  m_simplex(m_kTop) = m_refPoint
Else
' reflected point above next-highest - this looks bad
  If m_refPoint.val < m_simplex(m_kTop).val Then  ' reflected below top; replace
    m_simplex(m_kTop) = m_refPoint
  End If
  ' contract point toward centroid (overlay centroid)
  For j = m_LB To m_UB
    m_genPoint.vars(j) = Con * m_genPoint.vars(j) + _
                         ConCmp * m_simplex(m_kTop).vars(j)
    m_x(j) = m_genPoint.vars(j) * m_scale(j)
  Next j
  m_genPoint.val = f.Value(m_x)
  m_nCalls = m_nCalls + 1&
  m_nCycle = m_nCycle + 1&
  m_nCon = m_nCon + 1&
  If m_genPoint.val < m_simplex(m_kTop).val Then
    m_simplex(m_kTop) = m_genPoint
  Else  ' no point was below highest - huddle in panic around lowest
    Dim temp As Double
    For k = 1 To m_nSimplex
      If k <> m_kBot Then
        For j = m_LB To m_UB
          temp = Hud * m_simplex(m_kBot).vars(j) + HudCmp * m_simplex(k).vars(j)
          m_simplex(k).vars(j) = temp
          m_x(j) = temp * m_scale(j)
        Next j
        m_simplex(k).val = f.Value(m_x)
      End If
    Next k
    m_nCalls = m_nCalls + m_nVary
    m_nCycle = m_nCycle + m_nVary
    m_nHud = m_nHud + 1&
  End If
End If
End Sub

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub reduceNMsort()
' find the largest, next-largest and smallest function values in the simplex

' set order of first two points
If m_simplex(1).val < m_simplex(2).val Then
  m_kBot = 1&
  m_kTop = 2&
Else
  m_kBot = 2&
  m_kTop = 1&
End If

' shuffle third point into place
If m_simplex(3).val < m_simplex(m_kBot).val Then  ' below bottom
  m_kMid = m_kBot
  m_kBot = 3&
ElseIf m_simplex(3).val > m_simplex(m_kTop).val Then  ' above top
  m_kMid = m_kTop
  m_kTop = 3&
Else  ' must be between other two (or equal)
  m_kMid = 3&
End If

' adjust ranking with remaining points (if any)
Dim j As Long
Dim temp As Double
For j = 4& To m_nSimplex
  temp = m_simplex(j).val
  If temp < m_simplex(m_kBot).val Then
    m_kBot = j
  ElseIf temp > m_simplex(m_kTop).val Then
    m_kMid = m_kTop
    m_kTop = j
  ElseIf temp > m_simplex(m_kMid).val Then
    m_kMid = j
  End If
Next j

' carry out tests to see if we are making progress (function is decreasing)
If m_fBot <= m_simplex(m_kBot).val Then  ' we made no progress
  m_noDrop = m_noDrop + 1&  ' increase level of despair
Else  ' the function has decreased
  m_noDrop = 0&  ' recover our good spirits
End If
' reset value for next did-function-decrease test, and use in move logic
m_fBot = m_simplex(m_kBot).val

#If DebugLevel >= 2& Then
  Print #m_ofn, "Simplex values sorted to:"
  Print #m_ofn, "  highest: point " & m_kTop & " value " & m_simplex(m_kTop).val
  Print #m_ofn, "  next-hi: point " & m_kMid & " value " & m_simplex(m_kMid).val
  Print #m_ofn, "  lowest:  point " & m_kBot & " value " & m_simplex(m_kBot).val
  Print #m_ofn, "  function-didn't-decrease count: " & m_noDrop
#End If
End Sub

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub reduceNMstart(ByRef f As I_Function)
' start or restart a simplex minimization cycle
m_bestValPrev = m_bestPoint.val  ' save the previous best value
m_nInit = m_nInit + 1&  ' count starts & restarts

m_noDrop = 0&  ' initialize failed-to-decrease counter
m_fBot = m_bestPoint.val  ' test value to see if function decreased

' get normalization factors for variables that make simplex coordinates near 1.0
m_scale = m_bestPoint.vars  ' set scale factors to best variable values
' ...but avoid scaling with a variable that's equal to zero
Dim j As Long
For j = m_LB To m_UB
  If m_scale(j) = 0# Then m_scale(j) = 1#
Next j

' make simplex of size "m_sizeStart" with centroid at best known point
' note that no simplex point will coincide with the best known point
If m_sizeStart <= 0# Then m_sizeStart = 0.1  ' sanity check
' back off all coordinates from best point
Dim shrunk As Double
shrunk = m_sizeStart / m_nSimplex
Dim k As Long
For k = 1& To m_nSimplex
  For j = m_LB To m_UB
    If m_bestPoint.vars(j) = 0# Then
      m_simplex(k).vars(j) = -shrunk
    Else
      m_simplex(k).vars(j) = 1# - shrunk
    End If
  Next j
Next k
' advance coordinates, one at each point (except first)
For k = 2& To m_nSimplex
  m_simplex(k).vars(k - 1&) = m_simplex(k).vars(k - 1&) + m_sizeStart
Next k

' evaluate function at vertices of simplex
For k = 1& To m_nSimplex
  For j = m_LB To m_UB
    m_x(j) = m_simplex(k).vars(j) * m_scale(j)  ' back in caller's units
  Next j
  m_simplex(k).val = f.Value(m_x)
Next k
m_nCalls = m_nCalls + m_nSimplex
m_nCycle = m_nSimplex ' count calls for this restart cycle

#If DebugLevel >= 1& Then
  Print #m_ofn, ""
  Print #m_ofn, "Initialization " & m_nInit & "  simplex size " & _
    CSng(m_sizeStart)
  For k = 1 To m_nSimplex
    Print #m_ofn, "  point " & k & "  value = " & m_simplex(k).val & _
      "  caller variables:"
    For j = m_LB To m_UB
      m_x(j) = m_simplex(k).vars(j) * m_scale(j)
    Next j
    Print #m_ofn, arrayStr(m_x, "    ")
  Next k
#End If
End Sub

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub reduceNMtest()
' this routine determines if any exit criterion has been met
' it reports its results by setting m_result.reason to a return code

' find largest variable & normalized distance from highest point to lowest point
Dim ab As Double
Dim big As Double
Dim jBig As Long
Dim sizSqrd As Double
big = Abs(m_simplex(m_kBot).vars(m_LB) * m_scale(m_LB))
jBig = m_LB
sizSqrd = (m_simplex(m_kTop).vars(m_LB) - m_simplex(m_kBot).vars(m_LB)) ^ 2
Dim j As Long
For j = m_LB + 1& To m_UB
  ab = Abs(m_simplex(m_kBot).vars(j) * m_scale(j))
  If big < ab Then
    big = ab
    jBig = j
  End If
  sizSqrd = sizSqrd + _
    (m_simplex(m_kTop).vars(j) - m_simplex(m_kBot).vars(j)) ^ 2
Next j
m_size = Sqr(sizSqrd)

' set simplex scale for possible later minimization cycle
m_sizeStart = ReSize * m_size
' but don't use a value larger than caller's specified initial value
If m_sizeStart > Grow * m_sizeInit Then m_sizeStart = Grow * m_sizeInit
' ...and don't let it get comparable to the tolerance
If m_sizeStart < 2# * m_rangeTol Then m_sizeStart = 2# * m_rangeTol

' find relative value difference between highest and lowest points
Dim temp As Double
' handle case when simplex values straddle zero
ab = Abs(m_simplex(m_kTop).val)
If Abs(m_fBot) < ab Then
  temp = Abs(m_fBot)
  m_range = ab
Else
  temp = ab
  m_range = Abs(m_fBot)
End If
If temp < 1024# * Tiny Then temp = 1024# * Tiny
m_range = m_range / temp - 1#

' update best-so-far point
If m_bestPoint.val > m_fBot Then
  m_bestPoint.val = m_simplex(m_kBot).val
  For j = m_LB To m_UB
    m_bestPoint.vars(j) = m_simplex(m_kBot).vars(j) * m_scale(j)  ' caller units
  Next j
  #If DebugLevel >= 2& Then
    Print #m_ofn, "Best point updated to simplex point " & m_kBot
  #End If
End If

' test for exit conditions; put result in m_result.reason
m_result.reason = NoReason  ' flag value; not a legal return code value
If m_fBot <= m_valueTol Then
  m_result.reason = NMretvalueToleranceMet
  #If DebugLevel >= 1& Then
    Print #m_ofn, "Exit criterion met: function value"
  #End If
ElseIf m_nCalls >= m_maxCalls Then
  m_result.reason = NMretTooManyCalls
  #If DebugLevel >= 1& Then
    Print #m_ofn, "Exit criterion met: function-call limit"
  #End If
' do not do a size exit unless enough calls have been made
ElseIf (m_size <= m_rangeTol) And (m_nCycle >= CycleMul * m_nVary) Then
  m_result.reason = NMretSizeToleranceMet
  #If DebugLevel >= 1& Then
    Print #m_ofn, "Exit criterion met: simplex size"
  #End If
' do not do a range exit unless enough calls have been made
ElseIf (m_range <= m_rangeTol) And (m_nCycle >= CycleMul * m_nVary) Then
  m_result.reason = NMretRangeToleranceMet
  #If DebugLevel >= 1& Then
    Print #m_ofn, "Exit criterion met: simplex value range"
  #End If
ElseIf big > Huge / 1024# Then
  m_result.reason = NMretVariableHuge
  #If DebugLevel >= 1& Then
    Print #m_ofn, "Exit criterion met: huge var " & jBig & " = " & big
  #End If
End If

#If DebugLevel >= 1& Then
  ' this is the most basic debug printout
  Print #m_ofn, "Calls " & m_nCalls & "  size " & CSng(m_size) & _
    "  value " & CSng(m_range) & "  fBot-1 " & CSng(m_fBot - 1#) & "  Vars:"
  For j = m_LB To m_UB
    m_x(j) = m_simplex(m_kBot).vars(j) * m_scale(j)
  Next j
  Print #m_ofn, arrayStr(m_x, "  ")
#End If

If m_result.reason = NoReason Then  ' no exit condition was met
  ' if function value has not decreased for "many" calls, force a restart
  If m_noDrop > DropMul * m_nVary Then
    m_result.reason = NoDecrease  ' flag value; not a legal return code value
    #If DebugLevel >= 1& Then
      Print #m_ofn, "Restart forced: no decrease in " & m_noDrop & " moves"
    #End If
  End If
End If
End Sub

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
#If DebugLevel >= 1& Then
  Private Sub InitialLog(ByRef f As I_Function)
  m_ofn = FreeFile
  If m_ofn = 0 Then
    MsgBox "Cannot open debug logging file """ & DBfileName & """" & vbLf & _
      "Program will now exit.", _
      vbCritical Or vbOKOnly, _
      "FATAL ERROR in NelderMead.reduceNM()"
    End  ' <><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>><><>
  End If

  ' drop a different log file for each invocation of the routine
  Open DBfileName & Format$(m_invoked, "000") & ".txt" For Output As #m_ofn
  
  Print #m_ofn, "**** Nelder-Mead Function Reduction Start " & Date & " " & Time
  #If DebugLevel = 1& Then
    Print #m_ofn, "Debug level: 1"
  #ElseIf DebugLevel = 2& Then
    Print #m_ofn, "Debug level: 2"
  #Else
    Print #m_ofn, "Debug level: unknown"
  #End If
  Print #m_ofn, "Version: " & NelderMeadVersion
  Print #m_ofn, "This is call number " & m_result.nInvoked & " to this routine"
  Print #m_ofn, "Reducing: """ & f.Name() & """"
  Print #m_ofn, "----- Input values:"
  Print #m_ofn, "Number of variables: " & m_nVary
  Print #m_ofn, "Initial point has value " & m_bestPoint.val & " at location:"
  Print #m_ofn, arrayStr(m_bestPoint.vars, "  ")
  Print #m_ofn, "Initial size: " & m_sizeStart
  Print #m_ofn, "Value tolerance: " & m_valueTol
  Print #m_ofn, "Size tolerance: " & m_rangeTol
  Print #m_ofn, "Range tolerance: " & m_rangeTol
  Print #m_ofn, "Function call limit: " & m_maxCalls
  End Sub
#End If

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
#If DebugLevel >= 1& Then
  Private Sub FinalLog()
  Print #m_ofn, "----- Exit"
  Print #m_ofn, "Exit reason: " & NMresultString(m_result.reason)
  Print #m_ofn, "Best point seen had value " & m_bestPoint.val & " at location:"
  Print #m_ofn, arrayStr(m_bestPoint.vars, "  ")
  Print #m_ofn, "Calls used: " & m_result.callsUsed
  Print #m_ofn, "Final simplex size: " & m_result.finalSize
  Print #m_ofn, "Final value range: " & m_result.finalRange
  Print #m_ofn, "Initializations: " & m_result.nInitialize
  Print #m_ofn, "Contract moves: " & m_result.nContract
  Print #m_ofn, "Extend moves: " & m_result.nExtend
  Print #m_ofn, "Reflect moves: " & m_result.nReflect
  Print #m_ofn, "Huddle moves: " & m_result.nHuddle
  Print #m_ofn, "**** Nelder-Mead Function Reduction Done " & Date & " " & Time
  Print #m_ofn, ""
  Close #m_ofn
  End Sub
#End If

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
#If DebugLevel >= 1& Then
  Private Function arrayStr(ByRef arrayOfValues As Variant, _
                            Optional ByVal padding As String = "", _
                            Optional ByVal lineLength As Long = 80) _
  As String
  ' Converts array values to string consisting of "lines" with linefeeds between
  ' them to keep line length to less than specified value, after adding
  ' "padding" to front of each line. Lines contain "index=value(index)" entries.
  ' There will always be at least one entry per line, no matter what
  ' "lineLength" is. The input array can contain anything that can be converted
  ' to a string, including different types if it is an array of Variants.
  '         John Trenholme - 19 Aug 2003
  Dim j As Long, jFirst As Long, jLast As Long
  Dim sAdd As String, sLine As String, sNow As String
  If (VarType(arrayOfValues) And vbArray) = 0 Then  ' input is a scalar item
    arrayStr = padding & "S=" & arrayOfValues
    Exit Function
  End If
  jFirst = LBound(arrayOfValues)
  jLast = UBound(arrayOfValues)
  sNow = ""  ' start with empty result
  For j = jFirst To jLast
    sAdd = j & "=" & arrayOfValues(j)  ' will add "index=value(index)" entry
    If j < jLast Then sAdd = sAdd & ","  ' separator for all but last entry
    If j = jFirst Then  ' do special case
      sLine = padding & sAdd  ' make first line, with entry
    ElseIf Len(sLine) + Len(sAdd) + 1& >= lineLength Then  ' it won't fit
      sNow = sNow & sLine & vbNewLine  ' spill line onto output string
      sLine = padding & sAdd  ' start new line, with entry
    Else  ' it will fit
      sLine = sLine & " " & sAdd  ' add entry to line, with separation space
    End If
  Next j
  sNow = sNow & sLine  ' add last line to result
  arrayStr = sNow
  End Function
#End If

'----------------------------- end of file -------------------------------------


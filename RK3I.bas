Attribute VB_Name = "RK3I"
'
'###############################################################################
'#
'# VBA Module file "RK3I.bas"
'#
'# Fixed-step Runge-Kutta order-3 integration of a system of 1st-order ODE's
'# using Hermite cubic interpolation between last two steps to get end values.
'# This method is used to avoid "jumps" when the step length or count changes.
'#
'# Started 2010-06-17 by John Trenholme
'#
'# This module exports the routines:
'#   Function getRk3iXvalues
'#   Function getRk3iYvalues
'#   Function getRk3iDvalues
'#   Function getRk3iPvalues
'#   Function getRk3iWatch
'#   Function getLocIndex
'#   Function interpH3
'#   Function isStrictlyRising
'#   Function locate
'#   Sub rk3iDerivs           ' <--- this is where you code the derivative(s)
'#   Sub rk3iGo               ' <--- this is the main integration routine
'#   Function rk3iVersion
'#   Sub setLocIndex
'#   Sub setRk3iPvalues
'#   Sub rk3iDerivsTest
'#
'# Note: the routine 'rk3iDerivs' here is coded for diode-pumped lasers
'#       and is part of Project TORRID
'#
'###############################################################################

Option Base 0          ' array base value when not specified - the default
Option Compare Binary  ' string comparison based on Asc(char) - the default
Option Explicit        ' forces explicit variable declaration - changes default

'############################## Module-Global Constants ########################
' Module-global Const values (convention: start with upper-case; suffix "_c")

Private Const Version_c As String = "2012-01-06"  ' date of latest code revision
Private Const File_c As String = "RK3I[" & Version_c & "]."  ' file name

Private Const Nwatch_c As Long = 18&  ' number of per-point watch slots

'############################## Module-Global Variables #######################
' Module-global variables (convention: suffix "_m")
' Retained between calls; initialize as 0, "" or False

Private locIndex_m As Long    ' index value last returned by 'locate'
Private resX_m() As Double    ' x at each step
Private resY_m() As Double    ' y() at each step
Private resD_m() As Double    ' dY/dX() at each step
Private step_m As Long        ' index of present point: 1, 2, ..., nStep
Private watch_m() As Variant  ' general-purpose user-coded watched item(s)
' note: the parameters are in a collection so that undefined keys cause errors
Private pars_m As Collection  ' parameters used by 'rk3iDerivs'

'############################## Exported Routines ##############################

'===============================================================================
Public Function Rk3iVersion(Optional ByVal trigger As Variant) As String
' Date of the latest revision to this code, as a string with format "yyyy-mm-dd"
Rk3iVersion = Version_c
End Function

'===============================================================================
Public Function getLocIndex() As Long
' Index value where 'locate' last found an interpolant
getLocIndex = locIndex_m
End Function

'===============================================================================
Public Function getRk3iXvalues() As Double()
' Return the array of X values used during Runke-Kutta integration. They will
' be equally-spaced, except for the last one at the caller's end point.
' Call rk3iGo before calling this routine; you get previous-integration results.
' For A = getRk3iXvalues(), A must be a dynamic array (it might get resized)
getRk3iXvalues = resX_m
End Function

'===============================================================================
Public Function getRk3iYvalues() As Double()
' Return the array of Y values used during Runke-Kutta integration. They will be
' at the points returned by getRk3iXvalues, indexed y(function_index, x_index).
' Call rk3iGo before calling this routine; you get previous-integration results.
' For A = getRk3iYvalues(), A must be a dynamic array (it might get resized)
getRk3iYvalues = resY_m
End Function

'===============================================================================
Public Function getRk3iDvalues() As Double()
' Return the array of dY/dX values used during Runke-Kutta integration. They
' will be at the points returned by getRk3iXvalues, indexed as
' y(function_index, x_index). Call rk3iGo before calling this routine; you get
' previous-integration results.
' For A = getRk3iDvalues(), A must be a dynamic array (it might get resized)
getRk3iDvalues = resD_m
End Function

'===============================================================================
Public Function getRk3iPvalues() As Collection
' Return the Collection of parameter values used by the routine 'rk3iDerivs'
' during Runke-Kutta integration.
Set getRk3iPvalues = pars_m
End Function

'===============================================================================
Public Function getRk3iWatch() As Variant
' The user can put any desired scalar into watch_m and get it here later.
getRk3iWatch = watch_m
End Function

'===============================================================================
Public Function interpH3( _
  ByVal x As Double, _
  ByRef xVals() As Double, _
  ByRef yVals() As Double, _
  ByRef derivs() As Double, _
  Optional ByVal startIndex As Variant) _
As Double()
' Find the interval where x lies in the strictly-increasing array xVals(), and
' then do cubic Hermite interpolation using the function values in yVals() and
' the derivative values in derivs() at the two ends of the interval. Return the
' array of interpolated y values. If x is outside xVals(), use extrapolation.
' If the optional argument "startIndex" is supplied, the x-interval search will
' begin in the interval that begins at xVals(startIndex). If it is not supplied,
' the search will begin in the interval where x was found on the previous call;
' this behaviour leads to quick searches when successive values of x are close
' to each other. To force a full search, supply a "startIndex" that is in the
' center of the xVals() array; this quickens the first call.
' yVals is an array of y vectors at each xVals() point, indexed as
' yVals(which-y, x-index). Similarly, we have derivs(which-deriv, x-index).
' The yVals() and derivs() arrays must align with each other, and they must
' both align with xVals(); none of this index alignment is checked for.
' See the routines "isStrictlyRising" and "locate" for useful information.
Const ID_c As String = File_c & "interpH3 Function"
Dim errNum As Long, errDes As String  ' for saving error Number & Description
On Error GoTo ErrHandler

Dim ndx As Long  ' array index at start of interval containing x
If IsMissing(startIndex) Then
  ndx = locate(x, xVals)
Else
  ndx = locate(x, xVals, startIndex)
End If

Dim h As Double  ' step size
h = xVals(ndx + 1&) - xVals(ndx)

Dim u As Double  ' fraction of interval
u = (x - xVals(ndx)) / h
Dim u2 As Double  ' common factor
u2 = u * u
' cardinal spline coefficients in interval; coefficient code is:
' cABCD -> A = ndx value, B = ndx slope, C = ndx+1 value, D = ndx+1 slope
Dim c1000 As Double, c0100 As Double, c0010 As Double, c0001 As Double
c1000 = 1# - (3# - 2# * u) * u2
c0100 = (1# - (2# - u) * u) * u * h
c0010 = (3# - 2# * u) * u2
c0001 = (u - 1#) * u2 * h
Dim jLo As Long, jHi As Long
jLo = LBound(xVals)
jHi = UBound(xVals)
Dim y() As Double  ' temporary result array
ReDim y(jLo To jHi)
Dim j As Long
For j = jLo To jHi
  y(j) = c1000 * yVals(j, ndx) + c0100 * derivs(j, ndx) + _
         c0010 * yVals(j, ndx + 1&) + c0001 * derivs(j, ndx + 1&)
Next j

interpH3 = y  ' array assignment sets dimensions of result array
Erase y  ' this should happen automatically; just being cautious
Exit Function  '----------------------------------------------------------------
ErrHandler:  '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
errNum = Err.Number  ' save error number
errDes = Err.Description  ' save error-description text
' supplement Description; did error come from a called routine or this routine?
errDes = errDes & vbLf & _
  "x = " & x & "  ndx = " & ndx & vbLf & _
  IIf(0& = InStr(errDes, "Problem in"), "Problem in ", "Called by ") & _
  ID_c & IIf(0 <> Erl, " line " & Erl, "")  ' undocumented error-line number
If Designing_C Then Stop: Resume ' hit F8 twice to return to error point
' re-raise error with this routine's ID as Source, and appended to Message
Err.Raise errNum, ID_c, errDes  ' stops here if 'Debug' selected (not 'End')
End Function

'===============================================================================
Public Function IsStrictlyRising( _
  ByRef ary() As Double, _
  Optional ByVal rejectOneElementCase As Boolean = True) _
As Boolean
' Return True if the elements in the supplied one-dimensional array "ary" are in
' strictly increasing order (i.e., a(j) < a(j+1) for all j), or False if not.
' An array of length one will cause an error unless the optional argument
' rejectOneElementCase is set to False, instead of the default True.
' Useful if you need to be sure about an array that will be sent to a routine
' that assumes strict increase, such as "locate."
Const ID_c As String = File_c & "isStrictlyRising Function"
Dim jLo As Long
On Error Resume Next  ' prepare to test dimensions
jLo = LBound(ary)  ' fails if dynamic array not dimensioned
If Err.Number <> 0& Then
  On Error GoTo 0  ' cancel "Resume Next" behavior
  Err.Raise 5&, ID_c, Error$(5&) & vbLf & _
    "Input array not dimensioned" & vbLf & _
    "Problem in " & ID_c
End If
Dim jHi As Long
jHi = UBound(ary, 2&)  ' succeeds if array has 2 or more dimensions
If Err.Number = 0& Then
  On Error GoTo 0  ' cancel "Resume Next" behavior
  Err.Raise 5&, ID_c, Error$(5&) & vbLf & _
    "Input array has more than one dimension" & vbLf & _
    "Problem in " & ID_c
End If
On Error GoTo 0  ' no more error trapping
jHi = UBound(ary)
If rejectOneElementCase Then  ' reject case that can't be tested for increase
  If jLo = jHi Then  ' you call that an array?!?
    Err.Raise 5&, ID_c, Error$(5&) & vbLf & _
      "Input array has only one element at index " & jLo & vbLf & _
      "Problem in " & ID_c
  End If
End If
IsStrictlyRising = True  ' default value
Dim j As Long
For j = jLo + 1& To jHi
  If ary(j - 1&) >= ary(j) Then  ' not strictly increasing
    IsStrictlyRising = False
    Exit For
  End If
Next j
End Function

'===============================================================================
Public Function locate( _
  ByVal x As Double, _
  ByRef xV() As Double, _
  Optional ByVal startIndex As Variant) _
As Long
' Given a value of X and an array of strictly increasing X values, return
' the index of the array value that begins the interval containing the X value.
' That is, return k such that xV(k) <= x < xV(k+1). If x < xV(lowest_index),
' then silently return lowest_index. If x >= xV(highest_index), then silently
' return highest_index - 1. If you do not want this "snap to nearest end
' interval" behavior, test for out-of-bounds before calling this routine. No
' test is made that xV is strictly increasing (because it would be costly), and
' so you will get unexpected results if that is not the case. You can use the
' "isStrictlyRising" function to test an array before using it. There is no
' limitation on the array bounds of xV, but (of course) there should be at least
' two elements in xV. If there is only one element, the return value is its
' index. If there are no elements (xV has not been dimensioned) an error is
' raised. If the optional argument "startIndex" is supplied, the search will
' begin in the interval that begins at xV(startIndex). If it is not supplied,
' the search will begin in the interval where x was found on the previous call;
' this behaviour leads to quick searches when successive values of x are close
' to each other. To force a full search, supply a "startIndex" that is in the
' center of the xV array; this quickens the first call.

Const ID_c As String = File_c & "locate Function"

Dim jLo As Long, jHi As Long  ' bounds on input array
On Error Resume Next  ' we are about to try something that might raise an error
jLo = LBound(xV)  ' this fails if xV is not (re)dimensioned
If 0& <> Err.Number Then  ' xV is not (re)dimensioned
  locIndex_m = -2147483648#  ' an extremely unlikely value, just for safety
  Err.Raise 4337&, ID_c, "Input array xV not (re)dimensioned" & vbLf & _
    "Problem in " & ID_c
End If
On Error GoTo 0  ' we assert that no error is possible from here to the end!
jHi = UBound(xV)

' take care of end cases
If jHi = jLo Then  ' xV has only 1 element
  locIndex_m = jLo
ElseIf x < xV(jLo + 1&) Then ' x is below top of first interval
  locIndex_m = jLo
ElseIf x >= xV(jHi - 1&) Then  ' x is within or above last interval
  locIndex_m = jHi - 1&
Else  ' xV(jLo+1) <= x < xV(jHi-1) ; must search for containing interval
  'Debug.Assert xV(jLo + 1&) <= x
  'Debug.Assert x < xV(jHi - 1&)
  ' phase 1: start with one interval; widen search with increasing steps
  Dim jBot As Long, jMid As Long, jTop As Long
  If Not IsMissing(startIndex) Then  ' user has supplied a starting point
    On Error Resume Next  ' Variant could contain non-numeric
    jBot = CLng(startIndex)
    If 0& <> Err.Number Then jBot = jLo + (jHi - jLo) \ 2&  ' silent fixup
    On Error GoTo 0  ' restore normal error handling
  Else  ' use the previous result, speeding up repeated closely-spaced searches
    jBot = locIndex_m
  End If
  If (jBot < jLo) Or (jBot >= jHi) Then  ' out of bounds; use entire array
    jBot = jLo + 1&  ' we know x >= xV(jLo + 1)
    jTop = jHi - 1&  ' we know x < xV(jHi - 1)
  Else  ' hunt for region containing x
    Dim incr As Long  ' number of intervals in x-containing region
    incr = 1&
    If x >= xV(jBot) Then  ' we must search up
      jTop = jBot + incr  ' initial region is one interval, above start
      Do While x >= xV(jTop)  ' x not yet in region
        jBot = jTop
        jTop = jTop + incr
        incr = incr + incr  ' double the step size
        If jTop >= jHi - 1& Then  ' ran off top - bisect
          jTop = jHi - 1&
          Exit Do
        End If
      Loop
    Else  ' we must search down
      jTop = jBot
      jBot = jBot - incr  ' initial region is one interval, below start
      Do While x < xV(jBot)  ' x not yet in region
        jTop = jBot
        jBot = jBot - incr
        incr = incr + incr  ' double the step size
        If jBot <= jLo Then  ' ran off bottom - bisect
          jBot = jLo
          Exit Do
        End If
      Loop
    End If
  End If
  'Debug.Assert x >= xV(jBot)
  'Debug.Assert x < xV(jTop)
  'Debug.Assert jLo < jBot
  'Debug.Assert jTop < jHi - 1&
  
  ' phase 2: narrow x-containing region with binary search (if necessary)
  Do While jBot + 1& < jTop
    jMid = jBot + (jTop - jBot) \ 2&  ' this form avoids overflow
    If x >= xV(jMid) Then  ' at or above midpoint
      jBot = jMid
      If x = xV(jBot) Then Exit Do  ' exact equality; exit immediately
    Else  ' below midpoint
      jTop = jMid
    End If
  Loop
  'Debug.Assert x >= xV(jBot)
  'Debug.Assert x < xV(jBot + 1&)
  locIndex_m = jBot  ' save result for next call
End If
locate = locIndex_m  ' return index at start of interval containing x
End Function

'===============================================================================
Public Sub rk3iDerivs( _
  ByVal x As Double, _
  ByRef y() As Double, _
  ByRef dydx() As Double, _
  ByVal phase As Long, _
  Optional ByVal which As Long = 0&)
' Given values of 'x' (scalar) and 'y' (array), fills in array of derivative
' values in dyDx(). Called repeatedly by rk3iGo. You can do several different
' problems by setting 'which' and writing case-handling code. The argument
' 'phase' is 1 at the first point in each 3-point step, 2 at the 1/2 point, 3
' at the 3/4 point, and 0 at the final interpolated point, in case the user
' wants access to the action being taken.
' Be careful to use the same array bounds & index values here and in rk3iGo.
' You can use the side parameters passed in "pars_m" during the calculation.
' You put the parameters in a Collection, and pass it using "setRk3iPvalues".
Const ID_c As String = File_c & "rk3iDerivs Sub"
Dim errNum As Long, errDes As String  ' for saving error Number & Description
On Error GoTo ErrHandler

Dim tmp As Double  ' general-purpose temporary

'---------- user-coded derivative case(s) go here ------------------------------

If which = 0& Then  ' do default case; this is diode pumping of slablet group
  ' quantities being integrated are:
  '   y(1)  ' ion number density in upper level [1E20 ions / cc]
  '   y(2)  ' bulk heat density per shot [J/cc]
  '   y(3)  ' edge heat per shot [J]
  '   y(4)  ' face heat (both sides) per shot [J]
  'If x >= priVarsCol_g("pumpTau").value Then Stop  ' DEBUG halt at end of diode pump pulse
  Dim iU As Double  ' ion number density in upper level (1E20 ions / cc)
  iU = y(1&)
  Dim doping As Double  ' ground-state initial ion density (1E20 ions / cc)
  doping = pars_m("dopingAvg")
  ' we need dimensions to get ASE rate, which depends on size & aspect ratio
  Dim wd As Double  ' length of "long" axis of pumped region (cm)
  wd = pars_m("slabWidthPump")  ' unchanged during pump
  Dim hi As Double  ' length of "short" axis of pumped region (cm)
  hi = pars_m("slabHeightPump")  ' unchanged during pump
  Dim hOvrW As Double  ' pump-region height-over-width - unchanged during pump
  hOvrW = hi / wd
  Dim thik As Double  ' thickness of all slablets [cm]
  thik = pars_m("slabThick")
  Dim pumpedVolume As Double  ' volume of all slablets [cc]
  pumpedVolume = wd * hi * thik  ' to convert from J / cc to J

  ' ASE calculation looks up material by name, hoping it's in tables
  Dim matName As String  ' text name of material - unchanged during pump
  matName = pars_m("matName")
  ' if material not in tables, we at least need the refractive index to get ASE
  Dim refInd As Double  ' refractive index of material - unchanged during pump
  refInd = pars_m("refInd")
  Dim fracTIR As Double, fracNot As Double  ' part that could do TIR (or not)
  fracTIR = pars_m("fracTIR")
  fracNot = 1# - fracTIR
  
  ' resonance line fraction of radiative rate, allowing for trapping
  Dim alphaR As Double  ' loss coefficient at resonance line peak [1/cm]
  ' modify the coefficient to allow for ground-state depletion
  ' note: you can invert the upper-to-ground transition; alphaR can be > 0
  alphaR = -(doping - 2# * iU) * pars_m("crossSecReson")
  Dim axisAvg As Double  ' average axis length; hopefully near-square case
  axisAvg = 0.5 * (wd + hi)  ' this won't work too far from square
  Dim gainLog As Double  ' gain along long axis is Exp(gainLog)
  ' instead of actual emission and absorption line profiles, we approximate
  ' them as square lines with partial overlap, using an "average" gain
  gainLog = 0.8 * alphaR * axisAvg  ' take 80% of actual line-peak gain
  ' fraction escaping is roughly independent of thickness for "sorta-cube" amps
  ' but depends on gain times long axis, and short / long axis ratio
  ' here, we use the result for a square slab, with an average axis length
  ' result from fit to SlabASE results for square slabs with no gas-channel gaps
  Const GlMin_c As Double = -7.041  ' bottom of parabola
  If gainLog < GlMin_c Then gainLog = GlMin_c  ' no increase after minimum
  ' fraction of resonance-line radiation escaping in overlap part
  Dim fEsc As Double
  fEsc = Exp((0.4647 + 0.033 * gainLog) * gainLog)  ' fit over -5 < alpha*L < 0
  ' fraction of overlap-region resonance radiation leaving face (rest hits edge)
  ' fit to SlabASE results for square slabs with no gas-channel gaps
  Dim fFaceOv As Double
  fFaceOv = 0.1763 / (thik / axisAvg) ^ 0.836  ' fit over 0.8 < T/L < 1.2
  ' fraction of not-overlapped resonance radiation leaving face (rest hits edge)
  Dim fFaceNo As Double
  fFaceNo = 0.17  ' TODO - improve by allowing for thickness change
  Dim bR As Double
  bR = pars_m("branchReson")
  ' overlap calculation should properly be integral over line shapes, but we
  ' use the simple model from the 2002 Ehrmann & Campbell paper (Google it)
  Dim fOvr As Double
  fOvr = pars_m("lineOverlapReson")
  Dim fR As Double  ' resonance-line decay multiplier (less than 1)
  ' we get (1-fOvr) from part not trapped + fOvr*fEsc from (partly) trapped
  fR = (1# - fOvr) + fOvr * fEsc
  If phase <= 1& Then  ' this is an output point
    watch_m(1&, step_m) = alphaR
    watch_m(2&, step_m) = fEsc
    watch_m(3&, step_m) = fFaceOv
    watch_m(4&, step_m) = fR
  End If
  
  ' laser line fraction of radiative rate, allowing for ASE
  Dim alphaL As Double  ' net gain coefficient at laser line peak [1/cm]
  ' we allow for the nearly-fixed loss of the thermal lower-level population
  alphaL = _
    iU * pars_m("crossSecLase") - pars_m("thermalLoss") * (1# - iU / doping)
  Dim LL As Double  ' laser-line effective path length for ASE [cm]
  LL = wd * AseTrap_G().efLen(alphaL * wd, hOvrW, matName, refInd)
  Dim bL As Double
  bL = pars_m("branchLase")
  ' multiplier due to ASE not undergoing TIR (hits both face & edge) from Maple
  ' note: this includes the gain for not-TIR rays, but is zero for TIR part
  ' that is, gain from Maple = gMaple = 0 * Gtir + fracNot * Gnot
  Dim faceL As Double
  faceL = AseTrap_G().faceASE(alphaL * thik, refInd)  ' times total spontaneous
  ' multiplier due to ASE hitting edge because of TIR (from SlabASE; thin slabs)
  ' note: this includes the gain for TIR rays, but is unity for non-TIR part
  ' that is, gain from SlabASE = gSlabASE = fracTIR * Gtir + fracNot * 1
  Dim expArg As Double  ' argument to exponential. limited to avoid overflow
  expArg = alphaL * LL
  If expArg > 600# Then expArg = 600#
  Dim edgeL As Double
  edgeL = Exp(expArg)
  Dim fL As Double  ' laser-line decay multiplier (> 1 unless gain very low)
  ' total gain is fracTIR * Gtir + fracNot * Gnot =  gMaple + gSlabASE - fracNot
  ' TODO correct edge vs. face heat (below) for faceASE fraction going to edge
  fL = faceL + edgeL - fracNot
  If phase <= 1& Then  ' this is an output point
    watch_m(5&, step_m) = alphaL
    watch_m(6&, step_m) = faceL
    watch_m(7&, step_m) = edgeL
    watch_m(8&, step_m) = fL
  End If
  
  ' medium-wavelength line fraction of radiative rate, allowing for ASE
  Dim alphaM As Double  ' gain coefficient at medium-wavelength line peak [1/cm]
  alphaM = iU * pars_m("crossSecMed")
  Dim Lm As Double  ' medium-line effective path length for ASE [cm]
  Lm = wd * AseTrap_G().efLen(alphaM * wd, hOvrW, matName, refInd)
  Dim bM As Double
  bM = pars_m("branchMed")
  Dim faceM As Double  ' faceASE from Maple - see above
  faceM = AseTrap_G().faceASE(alphaM * thik, refInd)
  expArg = alphaM * Lm
  If expArg > 600# Then expArg = 600#
  Dim edgeM As Double  ' edgeASE from SlabASE - see above
  edgeM = Exp(expArg)
  Dim fM As Double  ' medium-wavelength-line decay multiplier (>= 1)
  ' total gain is fracTIR * Gtir + fracNot * Gnot =  gMaple + gSlabASE - fracNot
  ' TODO correct edge vs. face heat for faceASE fraction going to edge
  fM = faceM + edgeM - fracNot
  If phase <= 1& Then  ' this is an output point
    watch_m(9&, step_m) = alphaM
    watch_m(10&, step_m) = faceM
    watch_m(11&, step_m) = edgeM
    watch_m(12&, step_m) = fM
  End If

  ' branch ratio to thermal line near 2.2 microns
  ' we are using the branching-ratio sum bR + bL + bM + bT = 1 to get bT
  Dim bT As Double
  bT = 1# - bR - bL - bM

  ' quench-rate multiplier of radiative rate = (doping/quench)^2
  Dim qr As Double  ' dope-to-quench ratio (unchanged during pump)
  'qr = pars_m("quenchRatio")
  '#####NCM 072611 - using mean-doping decay time, with no inner-slab penalty
  qr = pars_m("dopingAvg") / pars_m("quenchDoping")
  '#####NCM
  Dim qRate As Double  ' quenching-rate multiplier of radiative rate
  qRate = qr * qr

  Dim decayUndoped As Double
  decayUndoped = pars_m("decayUndoped")
  Dim radRate As Double  ' radiative rate of decay (low doping, powdered sample)
  radRate = 1# / decayUndoped  ' per 탎

  ' sum up decay rates from various channels to get total rate
  ' note that a low-doped sample will have rateMul = 1 - bR * (1 - fR)
  Dim rateMul As Double  ' multiplier of rad. decay rate; trap, ASE, & quench
  rateMul = (qRate + bR * fR + bL * fL + bM * fM + bT)
  Dim decayInst As Double  ' instantaneous decay time [탎]
  decayInst = decayUndoped / rateMul

  ' derivative of pump-level number density (1E20 ions / cc / 탎)
  Dim pumpRate As Double
  pumpRate = pars_m("pumpRate")  ' 1E20 ions / cc / 탎
  Dim extraHeatFraction As Double
  extraHeatFraction = pars_m("extraHeatFraction")
  Dim pumpRateLessExtra As Double  ' 1E20 ions / cc / 탎
  pumpRateLessExtra = pumpRate * (1# - extraHeatFraction)
  ' net upper-level ions/cc/탎 is pump input minus total output (5 channels)
  Dim decayRate As Double
  decayRate = iU * radRate * rateMul
  Dim netIonRate As Double
  netIonRate = pumpRateLessExtra - decayRate  ' 1E20 ions / cc / 탎
  '------------------------------------------------------------------
  ' derivative of upper-level ion inversion rate 1E20 ions / cc / 탎
  dydx(1&) = netIonRate
  '------------------------------------------------------------------
  If phase <= 1& Then  ' this is an output point
    watch_m(13&, step_m) = qRate
    watch_m(14&, step_m) = rateMul
    watch_m(15&, step_m) = decayInst
    watch_m(16&, step_m) = decayRate
    watch_m(17&, step_m) = netIonRate
  End If

  ' ***** power balance *****
  
  Dim wP As Double, wR As Double, wL As Double, wM As Double  ' wavelengths
  wP = pars_m("wavelengthPump")
  wR = pars_m("wavelengthReson")
  wL = pars_m("wavelengthLase")
  wM = pars_m("wavelengthMed")

  ' calculate the total input power, light + thermal, in J / 탎 = MW
  Dim powerIn As Double
  ' convert pumpRate from 1E20 ions / cc / 탎 to MW
  ' 19.864455 is Planck's constant * lightspeed (using NIST values) * scaling
  powerIn = pumpRate * pumpedVolume * 19.864455 / wP
  
  ' we now have input, so calculate upper-level rate + losses & compare
  
  ' derivative of upper-level stored energy J / 탎 = MW
  Dim powerUpper As Double
  powerUpper = netIonRate * pumpedVolume * 19.864455 / wR
  
  ' "extra" bulk heat = pump that never got to the upper level
  Dim lostPump As Double, lostPumpDen As Double
  lostPump = powerIn * extraHeatFraction  ' J / 탎 = MW
  lostPumpDen = lostPump / pumpedVolume  ' J / 탎 / cc = MW / cc

  ' power flowing out of pump level into upper level
  Dim powerToUpper As Double
  powerToUpper = (powerIn - lostPump) * wP / wR ' J / 탎 = MW
  ' heat from pump excitation transition to upper-level ion
  Dim defect As Double, defectDen As Double
  defect = powerToUpper * (wR / wP - 1#) ' J / 탎 = MW
  defectDen = defect / pumpedVolume  ' J / 탎 / cc = MW / cc

  Dim thermalRate As Double  ' common multiplier of bulk-heat decay channels
  thermalRate = iU * radRate * 19.864455  ' J * microns / 탎 / cc
  ' rates of bulk heating J / 탎 / cc = MW / cc
  Dim heatQ As Double, heatL As Double, heatM As Double, heatT As Double
  heatQ = thermalRate * qRate / wR ' from quenching transitions
  heatL = thermalRate * bL * fL * (1# / wR - 1# / wL)  ' laser lower to ground
  heatM = thermalRate * bM * fM * (1# / wR - 1# / wM)  ' medium lower to ground
  heatT = thermalRate * bT / wR  ' all decays around 2.2 탆 go directly to heat
  Dim bulkHeatDen As Double  ' total bulk heating J / 탎 / cc = MW / cc
  bulkHeatDen = lostPumpDen + defectDen + heatQ + heatL + heatM + heatT
  '------------------------------------------------------------------
  ' derivative of bulk heat density J / 탎 / cc = MW / cc
  dydx(2&) = bulkHeatDen
  '------------------------------------------------------------------
  
  Dim bulkheat As Double  ' find total heat power
  bulkheat = bulkHeatDen * pumpedVolume  ' J / 탎 = MW

  ' common resonance-line power-escaping term (edge+face)
  Dim resEscape As Double
  resEscape = pumpedVolume * thermalRate * bR / wR
  ' derivative of edge heat; laser & medium-wavelength lines
  ' parts being radiated (stim. + spon.) J / 탎 = MW
  Dim heatER As Double, heatEL As Double, heatEM As Double
  ' resonance-line overlap escape and non-overlap
  heatER = _
    resEscape * (fOvr * fEsc * (1# - fFaceOv) + (1# - fOvr) * (1# - fFaceNo))
  ' want fracTIR * gTir = gSlabASE - fracNot = edge[L,M] - fracNot
  heatEL = pumpedVolume * thermalRate * bL * (edgeL - fracNot) / wL
  heatEM = pumpedVolume * thermalRate * bM * (edgeM - fracNot) / wM
  Dim heatEdge As Double
  heatEdge = heatER + heatEL + heatEM
  '------------------------------------------------------------------
  ' derivative of edge heat (J / 탎ec)
  dydx(3&) = heatEdge
  '------------------------------------------------------------------

  ' derivative of face heat; laser & medium-wavelength lines J / 탎 = MW
  Dim heatFR As Double, heatFL As Double, heatFM As Double
  ' resonance-line overlap escape and non-overlap
  heatFR = resEscape * (fOvr * fEsc * fFaceOv + (1# - fOvr) * fFaceNo)
  ' allow negative escape (tiny negative is possible) 2012-01-06 JBT
  Debug.Assert Abs(heatER + heatFR - resEscape * fR) <= _
    Abs(0.0000000001 * resEscape)
  ' want fracNot * gNot = gMaple = face[L,M]
  heatFL = pumpedVolume * thermalRate * bL * faceL / wL
  heatFM = pumpedVolume * thermalRate * bM * faceM / wM
  Dim heatFace As Double
  heatFace = heatFR + heatFL + heatFM
  '------------------------------------------------------------------
  ' derivative of face heat (J / 탎ec)
  dydx(4&) = heatFace
  '------------------------------------------------------------------

  ' sum up the power-out terms (where the pump power went)
  Dim powerUpperPlusHeat As Double  ' storage+heat J / 탎 = MW
  powerUpperPlusHeat = powerUpper + lostPump + bulkheat + heatEdge + heatFace
  
  ' check power balance at this point, and keep watch on worst case
  Dim powerError As Double
  powerError = powerUpperPlusHeat - powerIn
  'Debug.Assert Abs(powerError) < 0.000000000001
  If phase <= 1& Then  ' this is an output point
    watch_m(18&, step_m) = powerError
  End If

  ' debug print to Immediate window
'  Debug.Print "t "; cram(x, 8); " iU "; cram(iU, 8); " qr "; cram(qr, 8)
'  Debug.Print " aR "; cram(alphaR, 8); " Lr "; cram(Lr, 8); " fR "; cram(fR, 8);
'  Debug.Print _
'    " aL "; cram(alphaL, 8); " LL "; cram(LL, 8); " fL "; cram(fL, 8)
'  Debug.Print " aM "; cram(alphaM, 8); " Lm "; cram(Lm, 8); " fM "; cram(fM, 8);
'  Debug.Print " d1 "; cram(dydx(1&), 8); " d2 "; cram(dydx(2&), 8); _
'    " d3 "; cram(dydx(3&), 8); " d4 "; cram(dydx(4&), 8)

'---------- end user-coded derivatives -----------------------------------------

ElseIf which = -1& Then  ' 3-equation test case with known solutions
  ' y = Log(1 + x)
  dydx(0&) = Exp(-y(0&))
  ' y = (4*tanh(x+atanh(1/2))-2)/3  [atanh(1/2)=0.549306144334055]
  dydx(1&) = 1# - y(1&) * (1# + 0.75 * y(1&))
  ' y = Sin(x*Pi/4)
  dydx(2&) = 0.25 * Pi * Cos(0.25 * Pi * x)

Else  ' caller supplied an undefined 'which' value
  ' Error 5 = "Invalid procedure call or argument"
  Err.Raise 5&, ID_c, _
    Error$(5&) & ":" & vbLf & _
    "Unknown derivative-type case value 'which'" & vbLf & _
    "Change to a known case, or write code for this case"
End If
Exit Sub  '----------------------------------------------------------------
ErrHandler:  '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
errNum = Err.Number  ' save error number
errDes = Err.Description  ' save error-description text
' supplement Description; did error come from a called routine or this routine?
errDes = errDes & vbLf & _
  "x = " & x & "  which = " & which & vbLf & _
  IIf(0& = InStr(errDes, "Problem in"), "Problem in ", "Called by ") & _
  ID_c & IIf(0 <> Erl, " line " & Erl, "")  ' undocumented error-line number
If Designing_C Then Stop: Resume ' hit F8 twice to return to error point
' re-raise error with this routine's ID as Source, and appended to Message
Err.Raise errNum, ID_c, errDes  ' stops here if 'Debug' selected (not 'End')
End Sub

'===============================================================================
Public Sub rk3iGo( _
  ByVal xStart As Double, _
  ByRef yStart() As Double, _
  ByVal dx As Double, _
  ByVal xStop As Double, _
  ByRef yStop() As Double, _
  Optional ByVal which As Long = 0&)
' Numerically integrates the system of ODE's specified by the derivative values
' calculated by rk3iDerivs, starting at xStart with function values yStart.
' Uses an array version of Ralston's third-order Runge-Kutta integration method.
' Integrates using fixed step size dx until x is at or past xStop, then uses
' Hermite cubic interpolation to fill in yStop with the final function values.
' If several derivative sets, for different problems, are coded in rk3iDerivs,
' uses the one specified by "which."
' Note: yStop() must be a dynamic array, since it will be ReDim'd here.
' Note: we only read from yStart once, so it can be the same as yStop.
Const ID_c As String = File_c & "rk3iGo Sub"
Dim errNum As Long, errDes As String  ' for saving error Number & Description
On Error GoTo ErrHandler

' quantities related to the step size
Dim h As Double
h = Abs(dx) * Sgn(xStop - xStart)  ' silently fix any sign error on dx
If h = 0# Then  ' we're not going to make much progress ...
  If xStop = xStart Then ' ... but that's what the caller wants
    yStop = yStart  ' array assignment; sets dimensions of yStop
    Exit Sub
  Else  ' ... caller wants an infinite amount of work - bail out
    ' Error 17 = "Can't perform requested operation"
    Err.Raise 17&, ID_c, _
      Error$(17&) & ":" & vbLf & _
      "integrate from xStart = " & xStart & " to xStop = " & xStop & vbLf & _
      "within finite time using a step size of dx = " & dx
  End If
End If

' these integer-ratio values times the step size are used in the RK steps
Dim h1d2 As Double, h3d4 As Double
h1d2 = h * 0.5
h3d4 = h * 0.75
Dim h2d9 As Double, h1d3 As Double, h4d9 As Double
h2d9 = h * 2# / 9#
h1d3 = h / 3#
h4d9 = h * 4# / 9#

Dim jLo As Long, jHi As Long  ' array bounds of function & derivative values
jLo = LBound(yStart)
jHi = UBound(yStart)

' local array storage
Dim d1() As Double, d2() As Double, d3() As Double  ' derivatives
ReDim d1(jLo To jHi), d2(jLo To jHi), d3(jLo To jHi)
Dim t() As Double, y() As Double  ' temp array, & dependent variables
ReDim t(jLo To jHi), yStop(jLo To jHi)  ' note: dimensioning caller's array here

Dim nStep As Long  ' number of integration steps (counting the initial point)
Const OverStep_c As Double = 8.8E-16  ' fraction of h past point w/o new one
nStep = Int((xStop - xStart) / h + 2# - OverStep_c)
Debug.Assert nStep >= 2&  ' must have at least a start point and a stop point

' make space for per-point watched quantities
ReDim watch_m(1& To Nwatch_c, 1& To nStep)

' saved values for later recall
ReDim resX_m(1& To nStep)
Dim k As Long
For k = 1& To nStep  ' make up list of x values that will be used
  resX_m(k) = xStart + (k - 1&) * h
Next k
ReDim resY_m(jLo To jHi, 1& To nStep)
Dim j As Long
For j = jLo To jHi  ' put initial function values into return array
  resY_m(j, 1&) = yStart(j)
Next j
ReDim resD_m(jLo To jHi, 1& To nStep)  ' space for derivatives

y = yStart  ' array assignment; sets dimensions of y
Dim v As Double
Dim x As Double
For step_m = 1& To nStep  ' step to final point; get slopes there
  x = resX_m(step_m)
  rk3iDerivs x, y, d1, 1&, which  ' get slopes at first (& last) point
  For j = jLo To jHi
    resD_m(j, step_m) = d1(j)
  Next j
  If step_m = nStep Then Exit For  ' at final point, just get slopes
  ' step from present location to new x
  For j = jLo To jHi  ' get function values using first-point slopes
    t(j) = y(j) + h1d2 * d1(j)
  Next j
  rk3iDerivs x + h1d2, t, d2, 2&, which  ' get slopes at second point
  For j = jLo To jHi  ' get function values using second-point slopes
    t(j) = y(j) + h3d4 * d2(j)
  Next j
  rk3iDerivs x + h3d4, t, d3, 3&, which  ' get slopes at third point
  For j = jLo To jHi  ' final function step is weighted average of three points
    v = y(j) + h2d9 * d1(j) + h1d3 * d2(j) + h4d9 * d3(j)
    y(j) = v
    resY_m(j, step_m + 1&) = v  ' we are at the next point now
  Next j
Next step_m
Debug.Assert (x - resX_m(nStep - 1&)) * Sgn(h) > 0#  ' after penultimate
Debug.Assert (resX_m(nStep) - x) * Sgn(h) >= 0#  ' before or at last

' interpolate using values & derivatives at RK points on each side of end point
' penultimate value,slope in resY_m(),resD_m(); final value,slope in y(),d1()
Dim u As Double, u2 As Double
u = 1# - (x - xStop) / h  ' final-step-width fraction to end point
u2 = u * u  ' common factor
' cardinal spline coefficients at end point; let p1 = point before final
' Runge-Kutta point, p2 = final runge-Kutta point; coefficient code is:
' cABCD -> A = value at p1, B = slope at p1, C = value at p2, D = slope at p2
Dim c1000 As Double, c0100 As Double, c0010 As Double, c0001 As Double
c1000 = 1# - (3# - 2# * u) * u2
c0100 = (1# - (2# - u) * u) * u * h
c0010 = (3# - 2# * u) * u2
c0001 = (u - 1#) * u2 * h
For j = jLo To jHi
  yStop(j) = c1000 * resY_m(j, nStep - 1&) + c0100 * resD_m(j, nStep - 1&) + _
             c0010 * y(j) + c0001 * d1(j)
Next j

' now set the final interpolated values into the after-execution arrays
resX_m(nStep) = xStop
For j = jLo To jHi
  resY_m(j, nStep) = yStop(j)
Next j
rk3iDerivs xStop, yStop, d1, 0&, which  ' get slopes at final point
For j = jLo To jHi
  resD_m(j, nStep) = d1(j)
Next j
Exit Sub  '---------------------------------------------------------------------

ErrHandler:  '<><><><> handler for routines that don't need wrapup code <><><>
errNum = Err.Number  ' save error number
errDes = Err.Description  ' save error-description text
' supplement Description; did error come from a called routine or this routine?
errDes = errDes & vbLf & _
  IIf(0& = InStr(errDes, "Problem in"), "Problem in ", "Called by ") & _
  ID_c & IIf(0 <> Erl, " line " & Erl, "")  ' undocumented error-line number
On Error GoTo 0  ' this clears the Err object & avoids recursion
If Designing_C Then Stop: Resume ' hit F8 twice to return to error point
' re-raise error with this routine's ID as Source, and appended to Message
Err.Raise errNum, ID_c, errDes  ' stops here if 'Debug' selected (not 'End')
End Sub

'===============================================================================
Public Sub setRk3iPvalues(ByRef params As Collection)
' Stash the parameters needed by 'rk3iDerivs' where that routine can see them.
' This is used to pass "side information" to the derivative routine.
Set pars_m = params
End Sub

'===============================================================================
Public Sub setLocIndex(ByVal newLocIndex As Long)
' Specify index value where 'locate' will first look for an interpolant
locIndex_m = newLocIndex
End Sub

'===============================================================================
Public Sub rk3iDerivsTest()  ' subroutine for testing & debugging 'rk3iDerivs'
Const ID_c As String = File_c & "testDerivs Sub"
Dim errNum As Long, errDes As String  ' for saving error Number & Description
On Error GoTo ErrHandler
Dim y(0& To 2&) As Double  ' use built-in test functions
Dim dydx() As Double
rk3iDerivs 0#, y(), dydx(), -1&
' rk3iDerivs 0#, y(), Nothing, dydx(), -31&  ' forced error - no such case
Exit Sub  '---------------------------------------------------------------------
ErrHandler:  '<><><><> handler for routines that don't need wrapup code <><><>
errNum = Err.Number  ' save error number
errDes = Err.Description  ' save error-description text
' supplement Description; did error come from a called routine or this routine?
errDes = errDes & vbLf & _
  IIf(0& = InStr(errDes, "Problem in"), "Problem in ", "Called by ") & _
  ID_c & IIf(0 <> Erl, " line " & Erl, "")  ' undocumented error-line number
If Designing_C Then Stop: Resume ' hit F8 twice to return to error point
' re-raise error with this routine's ID as Source, and appended to Message
Err.Raise errNum, ID_c, errDes  ' stops here if 'Debug' selected (not 'End')
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


Attribute VB_Name = "NiceNums"
'#
'###############################################################################
'#
'# Visual Basic for Applications (VBA) source file "NiceNums.bas"
'#
'# Routine to supply "nice" numbers in a specified interval.
'#
'#  John Trenholme - Started 2004-11-22 - this version 2006-09-19
'#
'#  Exports the routines:
'#    NiceNumbers
'#
'###############################################################################

Option Explicit

'===============================================================================
Public Function NiceNumbers( _
  ByVal intervalStart As Double, _
  ByVal intervalEnd As Double, _
  Optional ByVal atLeast As Double = 2.5, _
  Optional ByVal subIntervals As Long = 0&, _
  Optional ByVal beyond As Double = 0.003) _
As Variant  ' Note: a Variant, not declared as an array, can still be an array
' Nice numbers are 1, 2 or 5 times an integer power of 10. They are useful for
' graph-axis ticks & labels, contour values, Slider ticks, etc. This routine
' calculates and returns nice numbers in the numeric interval between
' "intervalStart" and "intervalEnd", extended by beyond * interval at both
' ends. There will be at least "atLeast" major numbers, with the interval
' between major numbers split into "subIntervals" smaller intervals. If you
' set "subIntervals" to zero (the default), this routine will select the better
' of 4 or 5 for you.
'
' Note that "atLeast" is a Double. Non-integer values allow fine tuning of the
' results when integer values would give too few or too many return values.
'
' You will get back a Variant which is an array with the results packaged into
' it. Assign the result to a local Variant (see usage below). We call the local
' Variant "niceNums" here. The return values are:
'
'   niceNums(0) = array of major values: niceNums(0)(j), j = 0, 1, .. nMajor-1
'   niceNums(1) = array of minor values: niceNums(1)(k), k = 0, 1, .. nMinor-1
'   niceNums(2) = number of major values in major array
'   niceNums(3) = number of minor values in minor array
'   niceNums(4) = size of major step (negative if intervalStart > intervalEnd)
'   niceNums(5) = size of minor step (negative if intervalStart > intervalEnd)
'   niceNums(6) = number of minor steps per major step (4 or 5 unless specified)
'   niceNums(7) = power of 10 that scales largest value to 1 <= |x| < 10
'   niceNums(8) = number of sub-minor steps per minor step (4 or 5 or 0)

' When done, niceNums(0) contains the array of major values. This array will
' contain at least "atLeast" values, and up to three times that many, depending
' on the particular values of "intervalStart" and "intervalEnd" you supply. Then
' niceNums(1) will contain an array of nice numbers between the major values;
' you can control how many per major interval with "subIntervals", but it's best
' to leave "subIntervals" at the default value of 0. This tells the routine to
' select the better of 4 (for major intervals that are 2*10^N) or 5 (for 1*10^N
' or 5*10^N).
'
' Note the infrequently-seen format for getting elements of the arrays:
'   nicenums(0)(j) = j-th element of array nicenums(0)
'   nicenums(1)(k) = k-th element of array nicenums(1)
'
' Other (scalar) values are also returned, as shown. Note that the sub-minor
' step count only makes sense if subIntervals was set to 0 in the call; you get
' back a sub-interval count of 0 otherwise.
'
' If "intervalEnd" is less than "intervalStart", the results will be supplied in
' decreasing order, rather than the more common increasing order.
'
' Note: the interval should be larger than 1.0E-290 or so, and the interval
' should be more than 1.0E-13 times the mean value of the ends, or
' near-underflow will cause irregular output intervals.
'
' Note: because of errors when dividing by 10 (since 1/10 is not an exact
' binary fraction) the values returned may differ by a few bits from exactly
' 1, 2 or 5 times a power of 10. If it is important that the returned major
' and minor values are exact (perhaps you want to print them), you may use the
' DigitRound(x, 14) function to make them exact, as indicated below. This
' will clean up the last few bits, but it's a bit slower.
'
' Note: You may wish to print the nice numbers in the format 1.000.. to 9.999..
' with a separate notation of "X 10^N", where N is some integer power. The
' return value in niceNums(7) is N, and you should multiply the returned major
' and minor values by 10^(-N) before using them. The largest major or minor
' value will then obey 1 <= |value| < 10, and all values will have 1 or fewer
' digits to the left of the decimal place.
'
' Usage example:
'
'   Dim niceNums As Variant
'   nicenums = NiceNumbers(3.5, 12.2)  ' other arguments are optional
'   Dim j as Long
'   Dim x as Double
'   For j = LBound(niceNums(0)) To UBound(niceNums(0))
'     x = niceNums(0)(j)  ' or x = grfDigitRound(niceNums(0)(j), 14)
'     ... do something with a major value in "x" ...
'   Next j
'
'   ' If you want to use minor values (perhaps for axis minor ticks) then...
'   For j = LBound(niceNums(1)) To UBound(niceNums(1))
'     x = niceNums(1)(j)  ' or x = grfDigitRound(niceNums(1)(j), 14)
'     ... do something with a minor value in "y" ...
'   Next j

' input-value protection (caller has broken contract, but we proceed anyway)

' protect against absurd values of atLeast (anything less than 2 is silly)
If atLeast < 1# Then atLeast = 1#

' negative or unity subIntervals values cause trouble; silently set them to 0
If subIntervals < 2& Then subIntervals = 0&

' protect against a beyond value so large and negative that it reverses ends
If beyond < -0.4999 Then beyond = -0.4999

' remember initial-order sign
Dim sign As Double
If intervalStart <= intervalEnd Then
  sign = 1#
Else
  sign = -1#
End If

' add on the specified extra space beyond the ends (could be negative)
Dim add As Double
add = beyond * (intervalEnd - intervalStart)

' make local values with endpoints in numeric order xLo <= xHi
Dim xHi As Double
Dim xLo As Double
xLo = sign * (intervalStart - add)
xHi = sign * (intervalEnd + add)

' protect against minor intervals getting below roundoff by limiting interval
Dim intervalMin As Double
If subIntervals = 0& Then intervalMin = 4.5 Else intervalMin = subIntervals
Const Eps_c As Double = 2.23E-16  ' 2^(-52)
intervalMin = Eps_c * (Abs(xLo) + Abs(xHi)) * atLeast * intervalMin

' protect against case xLo = xHi = 0 (or close to it)
Const Tiny_c As Double = 1E-300
If intervalMin < Tiny_c Then intervalMin = Tiny_c
' push ends out if they are too close
If xHi - xLo < intervalMin Then
  xLo = 0.5 * (xLo + xHi - intervalMin)
  xHi = xLo + intervalMin
End If

' get the unscaled tentative interval size
Dim bigStep As Double
bigStep = (xHi - xLo) / atLeast  ' note AtLeast >= 1.0, so result is >= 0

' reduce interval size to the range 1.0 <= bigStep < 10.0
Dim scale10 As Double
scale10 = 1#               ' keep track of power of 10
Do While bigStep < 1#      ' increase bigStep if too small
  bigStep = bigStep * 10#
  scale10 = scale10 / 10#  ' slight error; 1/10 is not an exact binary fraction
Loop
Do While bigStep >= 10#    ' decrease bigStep if too large
  bigStep = bigStep / 10#  ' slight error; 1/10 is not an exact binary fraction
  scale10 = scale10 * 10#
Loop

' get nice interval size (1.0, 2.0, 5.0), nice minor interval count (5, 4, 5)
' and also nice sub-minor interval count (5, 5, 4)
Dim nSub As Long
Dim subMinor As Long
If bigStep >= 5! Then        ' 5.0 <= bigStep < 10.0
  bigStep = 5! * scale10
  nSub = 5&                    ' sub is 1.0
  subMinor = 5&                ' subSub is 0.2
ElseIf bigStep >= 2! Then    ' 2.0 <= bigStep < 5.0
  bigStep = 2! * scale10
  nSub = 4&                    ' sub is 0.5
  subMinor = 5&                ' subSub is 0.1
Else                         ' 1.0 <= bigStep < 2.0
  bigStep = scale10
  nSub = 5&                    ' sub is 0.2
  subMinor = 4&                ' subSub is 0.05
End If

' set number of minor intervals per major interval
Dim stepsUse As Long
If subIntervals = 0& Then  ' special value of 0 is flag: use calculated value
  stepsUse = nSub          ' caller trusts us to do the right thing
Else
  stepsUse = subIntervals  ' caller has their own opinion; respect it
  subMinor = 0&            ' however, subMinor no longer makes sense
End If

' take off an integer number of major intervals near center of interval
' keep the result in a Double to get 53-bit accuracy (Longs are too small)
' if this is not done, jNice (below) can overflow
Dim intAvg As Double
intAvg = Int((0.5 * xLo + 0.5 * xHi) / bigStep)  ' this form reduces overflow
Dim xAvg As Double
xAvg = intAvg * bigStep
xLo = xLo - xAvg
xHi = xHi - xAvg

' get an integer number of major intervals, above the lower end
Dim jNice As Long
jNice = Int(xLo / bigStep) + 2#

' move to major value just below lower end of interval
Do While bigStep * jNice >= xLo
  jNice = jNice - 1#
Loop

' move up by minor intervals until within interval (including lower end)
' keep modulo counter for minor values
Dim jSubInt As Long
Dim subStep As Double
subStep = bigStep / stepsUse            ' switch to minor step value...
jNice = jNice * stepsUse                ' ... and adjust count to match
jSubInt = 0&                            ' modulo is zero at major interval
Do While subStep * jNice < xLo          ' keep stepping while below lower end
  jNice = jNice + 1#
  jSubInt = jSubInt + 1&                ' use modulo count for minor intervals
  If jSubInt >= stepsUse Then jSubInt = 0&
Loop

' calculate values and put into arrays
Dim xNice As Double
Dim jSpace As Long
' make sure arrays have enough space to hold results; will surely fit into this
jSpace = 3& + Int((xHi - xLo) / subStep)
Dim majors() As Double
Dim minors() As Double
ReDim majors(0 To jSpace)
ReDim minors(0 To jSpace)
Dim jMajor As Long
Dim jMinor As Long
jMajor = -1&  ' indices to be pre-incremented; start below first slot
jMinor = jMajor
Dim pwr10 As Integer, pwrNow As Integer
pwr10 = -308
Do
  xNice = subStep * jNice      ' get the nice value
  ' we have gone beyond the end, & have produced enough major values
  If (xNice > xHi) And (jMajor >= atLeast - 1#) Then Exit Do
  If jSubInt = 0& Then         ' this is a major interval
    jMajor = jMajor + 1&
    ' save the major value, restoring sign info
    majors(jMajor) = (xNice + xAvg) * sign
  Else                         ' this is a minor interval
    jMinor = jMinor + 1&
    ' save the minor value, restoring sign info
    minors(jMinor) = (xNice + xAvg) * sign
  End If
  ' Find largest power of 10 needed to scale into 1 <= |nice| < 10
  If (xNice + xAvg) <> 0# Then
    Const Log10e_c As Double = 0.43429448 + 1.90325183E-09
    pwrNow = Int(Log10e_c * Log(Abs(xNice + xAvg)))
    If pwr10 < pwrNow Then pwr10 = pwrNow
  End If
  ' Step to the next value
  jNice = jNice + 1#
  jSubInt = jSubInt + 1&       ' use modulo count for sub-intervals
  If jSubInt >= stepsUse Then jSubInt = 0&
Loop

' resize arrays so they are just large enough - then LBound & UBound work OK
If jMajor >= 0 Then
  ReDim Preserve majors(0 To jMajor)
Else  ' there has been a serious logic error
  Err.Raise 51, "Graph!NiceNumbers", _
  "Error in Graph!NiceNumbers:" & vbLf & vbLf & _
  "Serious logic failure in routine." & vbLf & _
  "Expected at least one major value but got none."
End If
If jMinor >= 0 Then
  ReDim Preserve minors(0 To jMinor)
Else  ' there has been a serious logic error
  Err.Raise 51, "Graph!NiceNumbers", _
  "Error in Graph!NiceNumbers:" & vbLf & vbLf & _
  "Serious logic failure in routine." & vbLf & _
  "Expected at least one minor value but got none."
End If

' pack the results into a Variant array, and return the packed Variant
NiceNumbers = Array(majors, minors, jMajor + 1, jMinor + 1, _
  bigStep * sign, bigStep * sign / stepsUse, stepsUse, pwr10, subMinor)
Erase majors, minors  ' release array memory
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



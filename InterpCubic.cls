VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InterpCubic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Bloom's method of cubic interpolation within an array of equi-spaced function values. This implementation devised & coded by John Trenholme."
'
'     _____         _                        _____         _      _.
'    |_   _|       | |                      / ____|       | |    (_)
'      | |   _ __  | |_   ___  _ __  _ __  | |      _   _ | |__   _   ___.
'      | |  | '_ \ | __| / _ \| '__|| '_ \ | |     | | | || '_ \ | | / __|
'     _| |_ | | | || |_ |  __/| |   | |_) || |____ | |_| || |_) || || (__.
'    |_____||_| |_| \__| \___||_|   | .__/  \_____| \__,_||_.__/ |_| \___|
'                                   | |
'                                   |_|
'
'###############################################################################
'#
'# Visual Basic Class source file "InterpCubic.cls"
'#
'# This class implements Bloom's method of cubic interpolation within an array
'# of equi-spaced function values The values to be interpolated are supplied
'# to the "setDataValues" routine in an array,. The argument (or X) values
'# corresponding to the function values are implicitly defined by specification
'# of both the argument at the lowest-index function value via "xFirst", and
'# the constant increment of the argument between function values via "xInc".
'#
'# In the April 1993 issue of the Hewlett-Packard Journal, page 41, Alan Bloom
'# published a short article explaining the interpolation method used in the
'# HP 83750 frequency sweeper series for display-graph smoothing. His method
'# uses Hermite cubics, and is related to the Catmull-Rom spline method (1974).
'#
'# Bloom cubics match in value and slope at the function values, and have a
'# piecewise-quadratic continuous first derivative. Interpolant values between
'# function values depend only on adjacent function values, and on values one
'# more node away from the interpolation interval. Thus the influence of a
'# single function value is restricted to the four intervals closest to it.
'# Because of the way the node slopes are approximated, Bloom cubics are exact
'# only up to terms quadratic in x; they do not match cubic functions exactly.
'#
'# In each interval between function values, the interpolant is a Hermite cubic
'# specified by values and slopes at the ends. We must approximate the slopes,
'# since they are not available in the function value array. Bloom cubics use
'# the slopes found by approximating the node slopes as the slope of the line
'# through the two function values that bracket an interval end node; this is
'# the same as the slope of a parabola through the interval-end node and its 2
'# bracketing nodes. Because this slope is the same for two adjacent intervals,
'# Bloom cubics match in both value and slope at the nodes.
'#
'# At the function-array ends, there is no beyond-the-end node to use in the
'# slope approximation. Bloom did not specify what to do about this. We have
'# chosen to use the slope of the parabola through the end 3 points. It turns
'# out that when we do this, the end-interval interpolant is the parabola that
'# matches value at the end points of the end interval, and has the usual slope
'# approximation at the next-to-end point. This is also the result we would get
'# by using the cubic through the three end points, with the same slope
'# approximation at the next-to-end point - that cubic also reduces to the
'# parabola. It is also the result of linear extrapolation of the slopes of the
'# end interval and the end two intervals (ignoring the center point).
'#
'# Choice of the extrapolation method beyond the data domain is available for
'# function values and derivatives, but not integrals. For function values, it
'# is set by the optional argument "extrapPwr" to the object.f() function.
'# Values are:
'#   extrapPwr = 0 -> extrapolate as the value of the function at the end
'#   extrapPwr = 1 -> extrapolate with slope, at the end, of the end parabola
'#   extrapPwr = 2 (default) -> extrapolate as continuation of the end parabola
'#   extrapPwr = anything else -> raise an error if extrapolation is attempted
'# For derivative values, extrapolation behavior is set by the optional
'# argument "extrapPwr" to the object.fDeriv() function.
'# Values are:
'#   extrapPwr = 0 -> extrapolate as the value of the derivative at the end
'#   extrapPwr = 1 -> (default) extrapolate the derivative of the end parabola
'#   extrapPwr = anything else -> raise an error if extrapolation is attempted
'#
'# Usage:
'#
'#   Dim ic As InterpCubic
'#   Set ic = New InterpCubic
'#   ic.xFirst = 3.5
'#   ic.xInc = 0.25
'#   ic.setDataValues fVals  ' supply an array of 3 or more function values
'#
'#   ic.f(x)          ' call repeatedly to get interpolated values
'#   ic.f(x, 0)       ' get interpolated or extrapolated values
'# or
'#   ic.fDeriv(x)     ' call repeatedly to get derivative of interpolated values
'#   ic.fDeriv(x, 0)  ' get derivative of interpolated or extrapolated values
'# or
'#   ic.fInt(xA, xB)  ' call repeatedly to get integral of interpolated values
'#
'#   Set ic = Nothing  ' it's good practice to explicitly destroy VBA objects
'#
'# This implementation designed and coded by John Trenholme - begun 2008-04-22
'#
'###############################################################################

' Note: line numbers for error localization added with MZtools 3.0

Option Explicit

Private Const Version_c As String = "2010-05-14"
Private Const F_c As String = "InterpCubic[" & Version_c & "]."

Private a_m() As Double    ' local copy of function array = interp. constant
Private b_m() As Double    ' coefficient of u in interpolant
Private c_m() As Double    ' coefficient of u^2 in interpolant
Private d_m() As Double    ' coefficient of u^3 in interpolant
Private xInc_m As Double   ' increment of argument value between points
Private init_m As Boolean  ' flag to disallow evaluation without initiation
Private nF_m As Long       ' number of intervals in the function array
Private nL_m As Long       ' lower-bound index of function array
Private nU_m As Long       ' upper-bound index of function array
Private nUm1_m As Long     ' upper-bound index of function array, minus 1
Private xA_m As Double     ' argument value at lower-bound function value
Private xB_m As Double     ' argument value at upper-bound function value

'################################# Events ######################################

'===============================================================================
Private Sub Class_Initialize()
' set flag to cause an error on a get-interpolant call before setDataValues
init_m = False  ' this is the default; just being cautious
xA_m = 0#  ' set some values for xStart and xInc, in case they are not called
xInc_m = 1#
End Sub

'===============================================================================
Private Sub Class_Terminate()
' dynamically-allocated arrays are automatically erased; just being cautious
Erase a_m, b_m, c_m, d_m
End Sub

'################################# Methods #####################################

'===============================================================================
Public Sub setDataValues(ByRef vals() As Double)
Attribute setDataValues.VB_Description = "Set the array of 3 or more equally-spaced function values that will be used in the interpolation."
      ' Set the array of function values that will be used in the interpolation.
      '   vals() = array of argument values at a set of 3 or more equi-spaced
      '            argument values
      ' Note: this isn't a Property to avoid problems assigning fixed-dimension arrays
      Const ID_c As String = F_c & "Sub setDataValues"
      Dim errNum As Long, errDes As String, errLin As String

1     On Error Resume Next  ' intercept access to an undefined or excess dimension

      Dim nL As Long
2     nL = LBound(vals)
      Const BadArg_c As Long = 5&  ' Invalid procedure call or argument
3     If Err.Number <> 0& Then  ' lower bound does not exist
4       On Error GoTo ErrHandler
5       Err.Raise BadArg_c, ID_c, Error(BadArg_c) & vbLf & _
          "Input array ""vals()"" not dimensioned"
6     End If

      Dim nU As Long
7     nU = UBound(vals, 2&)
8     If Err.Number = 0& Then  ' there is a second dimension (and possibly even more)
9       On Error GoTo ErrHandler
10      Err.Raise BadArg_c, ID_c, Error(BadArg_c) & vbLf & _
          "Input array ""vals()"" has more than 1 dimension"
11    End If

12    On Error GoTo ErrHandler

13    nU = UBound(vals)

      ' get array element count
      Dim n As Long
14    n = nU - nL + 1&

15    If n < 3& Then  ' raise error on array-length check failure
16      Err.Raise BadArg_c, ID_c, Error(BadArg_c) & vbLf & _
          "Array ""vals"" must have >= 3 elements, but" & vbLf & _
          "it has element count = " & n & " from Dim(" & nL & " To " & nU & ")"
17    End If

18    nF_m = n - 1&  ' save interval count for setXlast
19    setXlast

      ' save lower and upper array bounds
20    nL_m = nL
21    nU_m = nU
22    nUm1_m = nU - 1&

      ' calculate the interpolation coefficients & save them
23    ReDim a_m(nL_m To nU) ' coefficient arrays match input dimensions
24    ReDim b_m(nL_m To nU)
25    ReDim c_m(nL_m To nU)
26    ReDim d_m(nL_m To nUm1_m)

      ' lowest-interval coefficients put a parabola through the first 3 points
27    a_m(nL) = vals(nL)
28    b_m(nL) = 2# * vals(nL + 1&) - 1.5 * vals(nL) - 0.5 * vals(nL + 2&)
29    c_m(nL) = 0.5 * (vals(nL) + vals(nL + 2&)) - vals(nL + 1&)
30    d_m(nL) = 0#
      ' get coefficients for cubics in internal intervals
      Dim j As Long
31    For j = nL + 1& To nU - 2&
32      a_m(j) = vals(j)
33      b_m(j) = 0.5 * (vals(j + 1&) - vals(j - 1&))
34      c_m(j) = vals(j - 1&) - 2.5 * vals(j) + 2# * vals(j + 1&) - 0.5 * vals(j + 2&)
35      d_m(j) = 1.5 * (vals(j) - vals(j + 1&)) + 0.5 * (vals(j + 2&) - vals(j - 1&))
36    Next j
      ' highest-interval coefficients put a parabola through the last 3 points
37    a_m(nUm1_m) = vals(nUm1_m)
38    b_m(nUm1_m) = 0.5 * (vals(nU) - vals(nU - 2&))
39    c_m(nUm1_m) = 0.5 * (vals(nU) + vals(nU - 2&)) - vals(nUm1_m)
40    d_m(nUm1_m) = 0#
      ' end-point coefficients allow parabolic extrapolation above end
41    a_m(nU) = vals(nU)
42    b_m(nU) = 0.5 * vals(nU - 2&) - 2# * vals(nUm1_m) + 1.5 * vals(nU)
43    c_m(nU) = c_m(nUm1_m)

44    init_m = True
45    Exit Sub '------------------------------------------------------------------

ErrHandler:
46    errNum = Err.Number
47    errDes = Err.Description & vbLf  ' set to put call-stack info on new line
48    If Erl = 0 Then errLin = "" Else errLin = ", line " & Erl
      'If Designing_C Then Stop: Resume  ' for debug before clearing Err object; 2X F8
49    On Error GoTo 0  ' this clears the Err object & avoids recursion
50    errDes = errDes & "Problem in " & ID_c & errLin  ' base-level error
      ' re-raise error with this routine's ID as Source, and appended to Message
51    Err.Raise errNum, ID_c, errDes
52    Resume  ' if debugging, set Next Statement here and F8 back to error point
End Sub

'################################# Properties #################################

'===============================================================================
' The value of the interpolant at the supplied position. If extrapolating, the
' optional argument sets the behavior.
Property Get f( _
  ByVal x As Double, _
  Optional ByVal extrapPwr As Integer = 2) _
As Double
      Const ID_c As String = F_c & "Property Get f"
      Dim errNum As Long, errDes As String, errLin As String
1     On Error GoTo ErrHandler

      Const NoCanDo_c As Long = 17&  ' Can't perform requested operation
2     If Not init_m Then
3       Err.Raise NoCanDo_c, ID_c, Error(NoCanDo_c) & vbLf & _
          "Function data not initialized" & vbLf & _
          "You must call ""setDataValues"" before calling this routine"
4     End If

      Dim xScaled As Double  ' X expressed as an array index, with a fractional part
5     xScaled = nL_m + (x - xA_m) / xInc_m
      Dim xL As Long  ' actual integer-value array index, at or below value of xScaled
6     xL = Int(xScaled)
7     If xL < nL_m Then xL = nL_m  ' argument is below start interval
8     If xL > nUm1_m Then xL = nUm1_m  ' argument is in or above end interval
      Dim u As Double  ' local variable in interpolation interval
9     u = xScaled - CDbl(xL)
      ' decide to interpolate or extrapolate
10    If ((u < 0#) Or (u >= 1#)) And (extrapPwr <> 2) Then
        ' extrapolate below start or above end
11      If extrapPwr = 0 Then  ' use constant end value
12        If u < 0# Then f = a_m(nL_m) Else f = a_m(nU_m)
13      ElseIf extrapPwr = 1 Then  ' use end value and slope
          Dim nEx As Long
14        If u < 0# Then nEx = nL_m Else nEx = nU_m
15        If u >= 1# Then u = u - 1#  ' readjust at top end
16        f = a_m(nEx) + u * b_m(nEx)
17      Else  ' no extrapolation allowed; raise an error
18        Err.Raise NoCanDo_c, ID_c, Error(NoCanDo_c) & vbLf & _
            "Can't extrapolate with ""extrapPwr"" = " & extrapPwr & vbLf & _
            "Set ""extrapPwr"" to 0, 1, or 2 (the default) to extrapolate"
19      End If
20    Else  ' normal case; interpolate between index xL and index xL + 1
21      f = a_m(xL) + u * (b_m(xL) + u * (c_m(xL) + u * d_m(xL)))
22    End If
23    Exit Property '-----------------------------------------------------------------

ErrHandler:
24    errNum = Err.Number
25    errDes = Err.Description & vbLf  ' set to put call-stack info on new line
26    If Erl = 0 Then errLin = "" Else errLin = ", line " & Erl  ' error line number
      'If Designing_C Then Stop: Resume  ' for debug before clearing Err object; 2X F8
27    On Error GoTo 0  ' this clears the Err object & avoids recursion
28    If 0& = InStr(errDes, "Problem in") Then  ' error was in this routine
29      errDes = errDes & "Problem in " & ID_c & errLin
30    Else  ' error was in some routine called from here
31      errDes = errDes & "Called from " & ID_c & errLin
32    End If
      ' re-raise error with this routine's ID as Source, and appended to Message
33    Err.Raise errNum, ID_c, errDes
34    Resume  ' if debugging, set Next Statement here and F8 back to error point
End Property

'===============================================================================
' The value of the derivative of the interpolant at the supplied position. If
' extrapolating, the optional argument sets the behavior.
Property Get fDeriv( _
  ByVal x As Double, _
  Optional ByVal extrapPwr As Integer = 1) _
As Double
      Const ID_c As String = F_c & "Property Get fDeriv"
      Dim errNum As Long, errDes As String, errLin As String
1     On Error GoTo ErrHandler

      Const NoCanDo_c As Long = 17&  ' Can't perform requested operation
2     If Not init_m Then
3       Err.Raise NoCanDo_c, ID_c, Error(NoCanDo_c) & vbLf & _
          "Function data not initialized" & vbLf & _
          "You must call ""setDataValues"" before calling this routine"
4     End If

      Dim xScaled As Double  ' X expressed as an array index, with a fractional part
5     xScaled = nL_m + (x - xA_m) / xInc_m
      Dim xL As Long  ' actual integer-value array index, at or below value of xScaled
6     xL = Int(xScaled)
7     If xL < nL_m Then xL = nL_m  ' argument is below start interval
8     If xL > nUm1_m Then xL = nUm1_m  ' argument is in or above end interval
      Dim u As Double  ' local variable in interpolation interval
9     u = xScaled - xL
      ' decide to interpolate or extrapolate
10    If ((u < 0#) Or (u >= 1#)) And (extrapPwr <> 1) Then
        ' extrapolate below start or above end
11      If extrapPwr = 0 Then  ' use constant end slope
12        If u < 0# Then fDeriv = b_m(nL_m) Else fDeriv = b_m(nU_m)
13      ElseIf extrapPwr = 1 Then  ' use end value and slope
          Dim nEx As Long
14        If u < 0# Then nEx = nL_m Else nEx = nU_m
15        If u >= 1# Then u = u - 1#  ' readjust at top end
16        fDeriv = b_m(nEx) + u * 2# * c_m(nEx)
17      Else  ' no extrapolation allowed; raise an error
18        Err.Raise NoCanDo_c, ID_c, Error(NoCanDo_c) & vbLf & _
            "Can't extrapolate with ""extrapPwr"" = " & extrapPwr & vbLf & _
            "Set ""extrapPwr"" to 0 or 1 (the default) to extrapolate"
19      End If
20    Else  ' normal case; interpolate between index xL and index xL + 1
21      fDeriv = (b_m(xL) + u * (2# * c_m(xL) + u * 3# * d_m(xL))) / xInc_m
22    End If
23    Exit Property '-----------------------------------------------------------------

ErrHandler:
24    errNum = Err.Number
25    errDes = Err.Description & vbLf  ' set to put call-stack info on new line
26    If Erl = 0 Then errLin = "" Else errLin = ", line " & Erl  ' error line number
      'If Designing_C Then Stop: Resume  ' for debug before clearing Err object; 2X F8
27    On Error GoTo 0  ' this clears the Err object & avoids recursion
28    If 0& = InStr(errDes, "Problem in") Then  ' error was in this routine
29      errDes = errDes & "Problem in " & ID_c & errLin
30    Else  ' error was in some routine called from here
31      errDes = errDes & "Called from " & ID_c & errLin
32    End If
      ' re-raise error with this routine's ID as Source, and appended to Message
33    Err.Raise errNum, ID_c, errDes
34    Resume  ' if debugging, set Next Statement here and F8 back to error point
End Property

'===============================================================================
' The value of the integral of the interpolant between the supplied positions.
Property Get fInt(ByVal xA As Double, ByVal xB As Double) As Double
Attribute fInt.VB_Description = "The value of the integral of the interpolant at the supplied argument value."
      Const ID_c As String = F_c & "Property Get fInt"
      Dim errNum As Long, errDes As String, errLin As String
1     On Error GoTo ErrHandler

      Const NoCanDo_c As Long = 17&  ' Can't perform requested operation
2     If Not init_m Then
3       Err.Raise NoCanDo_c, ID_c, Error(NoCanDo_c) & vbLf & _
          "Function data not initialized" & vbLf & _
          "You must call ""setDataValues"" before calling this routine"
4     End If

      ' always integrate from most negative x to most positive x, then fix up later
      Dim way As Double
5     If (xA <= xB) Eqv (xInc_m >= 0#) Then ' function data is in forward order
6       way = 1#
7     Else  ' function data is in reverse order; reverse ends of integration
8       way = xA
9       xA = xB
10      xB = way
11      way = -1#
12    End If

      Dim xIB As Double  ' xB expressed as an array index, with a fractional part
13    xIB = nL_m + (xB - xA_m) / xInc_m
      Dim xLB As Long  ' actual integer-value array index, at or below value of xIB
14    xLB = Int(xIB)
15    If xLB < nL_m Then xLB = nL_m  ' extrapolate below start
16    If xLB > nUm1_m Then xLB = nUm1_m  ' extrapolate above end
      Dim u As Double  ' local variable in interpolation interval
      ' integral of interpolant between index xLB and argument xB
17    u = xIB - xLB
      Dim res As Double
      Const Third_c As Double = 1# / 3#
18    res = (a_m(xLB) + (0.5 * b_m(xLB) + (Third_c * c_m(xLB) + _
            0.25 * d_m(xLB) * u) * u) * u) * u
      Dim xIA As Double  ' xA expressed as an array index, with a fractional part
19    xIA = nL_m + (xA - xA_m) / xInc_m
      Dim xLA As Long  ' actual integer-value array index, at or below value of xIB
20    xLA = Int(xIA)
21    If xLA < nL_m Then xLA = nL_m  ' extrapolate below start
22    If xLA > nUm1_m Then xLA = nUm1_m  ' extrapolate above end
      Dim j As Long
      ' add on any complete intervals between the end-point intervals
23    For j = xLA To xLB - 1&
24      res = res + a_m(j) + 0.5 * b_m(j) + Third_c * c_m(j) + 0.25 * d_m(j)
25    Next j
      ' remove integral of interpolant between index xLA and argumentxA
26    u = xIA - xLA
27    res = res - (a_m(xLA) + (0.5 * b_m(xLA) + (Third_c * c_m(xLA) + _
            0.25 * d_m(xLA) * u) * u) * u) * u
28    fInt = res * way * xInc_m
29    Exit Property '-----------------------------------------------------------------

ErrHandler:
30    errNum = Err.Number
31    errDes = Err.Description & vbLf  ' set to put call-stack info on new line
32    If Erl = 0 Then errLin = "" Else errLin = ", line " & Erl  ' error line number
      'If Designing_C Then Stop: Resume  ' for debug before clearing Err object; 2X F8
33    On Error GoTo 0  ' this clears the Err object & avoids recursion
34    If 0& = InStr(errDes, "Problem in") Then  ' error was in this routine
35      errDes = errDes & "Problem in " & ID_c & errLin
36    Else  ' error was in some routine called from here
37      errDes = errDes & "Called from " & ID_c & errLin
38    End If
      ' re-raise error with this routine's ID as Source, and appended to Message
39    Err.Raise errNum, ID_c, errDes
40    Resume  ' if debugging, set Next Statement here and F8 back to error point
End Property

'===============================================================================
' Argument value at the first function array index.
Property Get xFirst() As Double
Attribute xFirst.VB_Description = "Argument value at the first function array index."
xFirst = xA_m
End Property
Property Let xFirst(ByVal RHS As Double)
xA_m = RHS
setXlast
End Property

'===============================================================================
' Increment of argument value between function array values; can be negative
' but not zero.
Property Get xInc() As Double
Attribute xInc.VB_Description = "Increment of argument value between function array values; can be negative but not zero."
xInc = xInc_m
End Property
Property Let xInc(ByVal RHS As Double)
      Const ID_c As String = F_c & "Property Let xInc"
      Dim errNum As Long, errDes As String, errLin As String
1     On Error GoTo ErrHandler

2     If RHS = 0# Then  ' raise error on zero separation of points
        Const BadArg_c As Long = 5&  ' Invalid procedure call or argument
3       Err.Raise BadArg_c, ID_c, Error(BadArg_c) & vbLf & _
          "Function-argument increment must be non-zero, but xInc = " & RHS
4     End If
5     xInc_m = RHS
6     setXlast
7     Exit Property '-----------------------------------------------------------------

ErrHandler:
8     errNum = Err.Number
9     errDes = Err.Description & vbLf  ' set to put call-stack info on new line
10    If Erl = 0 Then errLin = "" Else errLin = ", line " & Erl  ' error line number
      'If Designing_C Then Stop: Resume  ' for debug before clearing Err object; 2X F8
11    On Error GoTo 0  ' this clears the Err object & avoids recursion
12    errDes = errDes & "Problem in " & ID_c & errLin  ' base-level error
      ' re-raise error with this routine's ID as Source, and appended to Message
13    Err.Raise errNum, ID_c, errDes
14    Resume  ' if debugging, set Next Statement here and F8 back to error point
End Property

'===============================================================================
' Argument value at the last function array index. Only valid after xStart and
' xInc have been set, and setDataValues has been called.
Property Get xLast() As Double
Attribute xLast.VB_Description = "Argument value at the last function array index. Only valid after xStart and xInc have been set, and setDataValues has been called."
xLast = xB_m
End Property

'===============================================================================
' Date of the latest revision to this Class, as a String formatted "yyyy-mm-dd".
Property Get Version() As String
Attribute Version.VB_Description = "Date of the latest revision to this Class, as a String formatted ""yyyy-mm-dd""."
Version = Version_c
End Property

'################################# Private #####################################

'===============================================================================
Private Sub setXlast()
' Setting xFirst, xInc, or dataValues can change X at the last index of the
' function array, so we call this when any of those items changes
xB_m = xA_m + nF_m * xInc_m  ' X value at high-index end of array
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~

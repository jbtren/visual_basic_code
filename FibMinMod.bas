Attribute VB_Name = "FibMinMod"
'       ______  _  _      __  __  _.
'      |  ____|(_)| |    |  \/  |(_)
'      | |__    _ | |__  | \  / | _  _ __.
'      |  __|  | || '_ \ | |\/| || || '_ \
'      | |     | || |_) || |  | || || | | |
'      |_|     |_||_.__/ |_|  |_||_||_| |_|
'
'###############################################################################
'#
'# Visual Basic Module "FibMinMod.bas"
'#
'# Fibonacci-search function minimizer and support routines.
'#
'# Devised & coded by John Trenholme. Started 2008-10-05.
'#
'# Exports the following routines:
'#   Function fibFuncCalls
'#   Function fibFuncMin
'#   Function fibMin
'#   Function fibMinSpan
'#   Sub fibMinTest
'#   Function fibMinVersion
'#   Function funcFib
'#
'###############################################################################

Option Explicit

Private Const Version_c As String = "2008-11-04"

Private fLo_m As Double
Private nCall_m As Long

'===============================================================================
Public Function fibMin( _
  ByRef x1 As Double, _
  ByRef x2 As Double, _
  ByRef erx As Double, _
  Optional ByVal functionIndex As Long = 0&) _
As Double
' Uses Fibonacci search to localize a minimum of the function "funcFib"
' (see below). Region searched is specified by the end points 'x1' and 'x2'.
' Search terminates when minimum is known to lie within a distance 'erx' of
' the return value of 'x'. This method is most useful when you need to find
' a rough minimum location of a difficult function. See 'funcFib' for the use
' of the optional argument.
'
' Note that functions with a quadratic form near the minimum cannot have that
' minimum located to better than about 1E-8 of the minimum location (more or
' less, depending on how steep the quadratic is), because of roundoff noise.
'
' Fibonacci search always takes the same number of function calls to reduce
' the initial interval by a specified fraction, independent of the function.
'
' Fraction    Calls    Here F(N) is the N'th Fibonacci number
' --------    -----
'   1/2           1    N:    0   1   2   3   4   5   6   7   8   9  10
'   1/3           2    F(N): 0   1   1   2   3   5   8  13  21  34  55
'   1/5           3
'   1/8           4
'       ...
'  1/F(N)       N-2    For large N, F(N) is nearly ((1+Sqr(5))/2)^N / Sqr(5)
'
' Reduction of the interval to the minimum allowed by Double precision roundoff
' takes 76 function calls.
'
' Sometimes, Fibonacci search can fail because the function value is the same
' at both evaluation points. The side the minimum lies on is then uncertain.
' The logic here is set to use the side of the latest evaluation, but this is
' probably right only about half the time. This problem can be reduced by
' trying several times while jittering the ends of the initial interval by at
' least an amount equal to the desired final uncertainty. Alternatively,
' uncomment the Err.Raise statement below and catch the error in your code.
Const ErrMin_c As Double = 1.4E-16  ' near Double precision limit
If erx < ErrMin_c Then erx = ErrMin_c
Dim erInv As Double
erInv = 1# / erx
' iterate up the Fibonacci sequence
' with 'ErrMin' value above, max a = 5527939700884757, max b = 8944394323791464
' these values will cause 76 function evaluations; that's the worst case
Dim a As Double, b As Double
a = 1#  ' set values to force at least one function evaluation
b = 2#
Do While b < erInv
  b = b + a
  a = b - a
Loop
' set the initial points
Dim xHi As Double, xLo As Double
xLo = x2  ' set so first point will be in the correct location
xHi = x2 + (x2 - x1) * a / b  ' outside point forces first point inside
Dim fLo As Double, fHi As Double
Const HugeDbl As Double = 1.79769313486231E+308 + 5.7E+293
fLo = HugeDbl  ' force use of the first function evaluation
' place points at Fibonacci locations
nCall_m = 0&
Do While b >= 1.5  ' stop at a = b = 1 (after iteration below)
  a = b - a  ' iterate down the Fibonacci sequence
  b = b - a
  xHi = xLo - (xHi - xLo) * a / b  ' new point
  fHi = funcFib(xHi, functionIndex)  ' single function evaluation here
  nCall_m = nCall_m + 1&
  ' the following statement will raise an error if the two function values
  ' are equal, making it impossible to choose a side
'  If fHi = fLo Then Err.Raise 17&, "fibMin", Error(17&) & vbLf & _
'    "because function values are equal in ""fibMin""" & vbLf & _
'    "funcFib(" & xLo & "," & functionIndex & ") = " & fLo & vbLf & _
'    "funcFib(" & fHi & "," & functionIndex & ") = " & fHi
  If fHi <= fLo Then  ' new point is better; swap points
    fLo = xLo
    xLo = xHi
    xHi = fLo
    fLo = fHi
  End If
Loop
' job complete - set result values
fLo_m = fLo
fibMin = xLo
End Function

'===============================================================================
Public Function fibFuncCalls(Optional ByVal trigger As Double = 0#) As Long
' Returns the number of function evaluations used by 'fibMin'.
' Use the 'trigger' argument to force evaluation order when called from Excel.
fibFuncCalls = nCall_m
End Function

'===============================================================================
Public Function fibFuncMin(Optional ByVal trigger As Double = 0#) As Double
' Returns the function value at the argument value returned by 'fibMin'.
' Use the 'trigger' argument to force evaluation order when called from Excel.
fibFuncMin = fLo_m
End Function

'===============================================================================
Public Function funcFib( _
  ByVal x As Double, _
  Optional ByVal functionIndex As Long = 0&) _
As Double
' The function that 'fibMin' minimizes. If you need to minimize several
' different functions, code them into the Case blocks and use the optional
' argument to select among them. If there is only one function, code it in
' the Case 0& block - no need to pass a value of 'functionIndex'.
Dim r As Double, t As Double
r = Sqr(2.26)
Select Case functionIndex
  Case 0&
    If x < r Then funcFib = 1000# * (r - x) Else funcFib = x - r
  Case 1&
    t = x - r
    funcFib = 1# + t * t
  Case Else
    Err.Raise 5&, "FibMinMod!funcFib", _
      "Invalid Function call argument in FibMinMod!funcFib" & vbLf & _
      "Function with index " & functionIndex & " not defined"
End Select
End Function

'===============================================================================
Public Function fibMinSpan( _
  ByRef x1 As Double, _
  ByRef x2 As Double, _
  ByRef erx As Double, _
  Optional ByVal functionIndex As Long = 0&) _
As Double
' Informational function to return maximum error of estimated minimum location.
' The user will not usually need to call this routine.
erx = Abs(erx)  ' silent sanity enforcement
Const ErrMin As Double = 1.4E-16  ' near Double precision limit
If erx < ErrMin Then erx = ErrMin
Dim erInv As Double
erInv = 1# / erx
' iterate up the Fibonacci sequence
' with 'ErrMin' value above, max a = 5527939700884757, max b = 8944394323791464
' these values will cause 76 function evaluations; that's the worst case
Dim a As Double, b As Double
a = 1#  ' set values to force at least one function evaluation
b = 2#
Do While b < erInv
  b = b + a
  a = b - a
Loop
' set the initial points
Dim xHi As Double, xLo As Double
xLo = x2  ' set so first point will be in the correct location
xHi = x2 + (x2 - x1) * a / b  ' outside point forces first point inside
Dim fLo As Double, fHi As Double
Const HugeDbl As Double = 1.79769313486231E+308 + 5.7E+293
fLo = HugeDbl  ' force use of the first function evaluation
' place points at Fibonacci locations
nCall_m = 0&
Do While b >= 1.5  ' stop at a = b = 1 (after iteration below)
  a = b - a  ' iterate down the Fibonacci sequence
  b = b - a
  xHi = xLo - (xHi - xLo) * a / b  ' new point
  fHi = funcFib(xHi, functionIndex)  ' single function evaluation here
  nCall_m = nCall_m + 1&
  If fHi < fLo Then  ' new point is better; swap points
    fLo = xLo
    xLo = xHi
    xHi = fLo
    fLo = fHi
  End If
Loop
' job complete - set result value
fibMinSpan = Abs(xHi - xLo)
End Function

'===============================================================================
Public Sub fibMinTest()
' Quick sanity check of operation. Useful for tracing operation.
Dim x As Double
x = fibMin(1#, 2#, 0.000001)
Debug.Print "fibMin(1#, 2#, 0.000001) -> x = " & x & " error " & x - Sqr(2.26)
Debug.Print "function = " & fibFuncMin() & "  calls = " & fibFuncCalls()
End Sub

'===============================================================================
Public Function fibMinVersion() As String
' Date of latest revision, as a string formatted "yyyy-mm-dd".
Application.Volatile  ' comment out in VB6; leave active in Excel VBA
fibMinVersion = Version_c
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Attribute VB_Name = "ZeroBracketMod"
Attribute VB_Description = "Module containing routine to return an interval containing at least one zero of a user-supplied function. Devised and coded by John Trenholme."
'      ______                  ______                    _          _.
'     |___  /                  | ___ \                  | |        | |
'        / /   ___  _ __  ___  | |_/ / _ __  __ _   ___ | | __ ___ | |_.
'       / /   / _ \| '__|/ _ \ | ___ \| '__|/ _` | / __|| |/ // _ \| __|
'      / /___|  __/| |  | (_) || |_/ /| |  | (_| || (__ |   <|  __/| |_.
'     /_____/ \___||_|   \___/ \____/ |_|   \__,_| \___||_|\_\\___| \__|
'
'###############################################################################
'#
'# Visual Basic 6 & VBA code module "ZeroBracketMod.bas"
'#
'# Single-direction zero bracketing routine using interval-increase strategy.
'#
'# Exports the routines:
'#   Function zeroBracket
'#   Function ZeroBracketModVersion
'#   Function zeroBracketUnitTest
'#
'# Devised and coded by John Trenholme - begun 2012-01-01
'#
'###############################################################################

Option Base 0          ' array base value when not specified - the default
Option Compare Binary  ' string comparison based on Asc(char) - the default
Option Explicit        ' forces explicit variable declaration - changes default

Private Const Version_c As String = "2012-01-04"
Private Const File_c As String = "ZeroBracketMod[" & Version_c & "]."

'===============================================================================
Public Function ZeroBracketModVersion() As String
Attribute ZeroBracketModVersion.VB_Description = "Date of the latest revision to this routine, as a string with the format 'YYYY-MM-DD'"
' Date of latest revision to this routine, as a string with format 'YYYY-MM-DD'
ZeroBracketModVersion = Version_c
End Function

'===============================================================================
Public Function zeroBracket( _
  ByVal f As Double, _
  ByRef x As Double, _
  ByRef dx As Double, _
  ByRef calls As Long, _
  Optional ByVal callMax As Long = 20&, _
  Optional ByVal stepRatio As Double = 2#) _
As Boolean
Attribute zeroBracket.VB_Description = "Routine to return an interval containing at least one zero of  a user-coded function specified in the call to this routine."
' Bracket the zero of an arbitrary function that is coded into the call,
' allowing application to any number of functions using one routine. This
' method is necessary because Visual Basic has no easy way to pass a function.
' One variable will be adjusted, starting at a specified value and moving in
' only one direction, until a zero of the function is bracketed. The bracket
' will have different signs at its ends, so continuous functions will have
' an odd number of zeros within the bracket. Once a zero is bracketed, a
' zero-finding routine such as 'zeroBrent' can be used to locate it exactly.
'
' Function arguments:
' f         input        function value, using x (& maybe other arguments)
' x         input+output function argument that will be adjusted
' dx        input+output step of x to get to next evaluation point
' calls     input+output number of function calls - SET TO ZERO BEFORE LOOP
' callMax   input        maximum number of function calls allowed
' stepRatio input        amount by which dx is increased between steps
'
' An agressive strategy is used, in which the step size is doubled after each
' function call. You can change the step multiplier using the optional argument
' 'stepRatio'. If you want equal steps (an evenly-spaced grid search) then set
' stepRatio = 1.
'
' If there is no zero in the specified direction, many calls might be made in
' vain. A default call limit is supplied, but you can specify your own limit.
' The default values for 'callMax' and 'stepRatio' allow an increase of 'dx'
' by a factor of a million before an error halt. If there is no zero by then,
' you should rethink the situation.
'
' Usage:
'
' This routine is called inside a simple Do .. Loop structure, as follows:
'
'  Dim myX As Double, myDx As Double, calls As Long  ' names here are arbitrary
'  myX = 12.5  ' initial point of function evaluation
'  myDx = 0.1  ' initial step to next evaluation point - one direction only
'  calls = 0&  ' call count MUST be initialized to zero to start correctly
'  Do
'  Loop While zeroBracket(Cos(myX), myX, myDx, calls)  ' call to this routine
'
' In this case, the function whose zero is to be bracketed was Cos(myX), but
' any arbitrary built-in or user-coded function can be inside the function call.
' This works because arguments that are supplied as expressions are evaluated
' before the call to the routine. For more examples, see the unit test code.
'
' On successful exit, one or more zeros of the function will be specified by
' the return values of 'myX' and 'myDx' - there will be a sign change between
' myX and myX + myDx. If 'myDx' is zero, an exact zero was found, and its
' location is at 'myX'. In the above case, myX = 15.6 and myDx = -1.6, so the
' zero lies between 14 and 15.6, not including the end points.
'
' If the function is too complicated to be coded into the function call, it
' can be evaluated inside the Do .. Loop structure as follows:
'
'  Dim myX As Double, myDx As Double, calls As Long, myFunc As Double
'  myX = 7.22
'  myDx = -1.5  ' searches "backwards" from 'myX'
'  calls = 0&
'  Do
'    '----- complicated code that evaluates myFunc(myX), such as...
'    myFunc = NPcomplete(myX, arg2, argg3, ...) - Atn(otherFunc(myX))
'  Loop While zeroBracket(myFunc, myX, myDx, calls)  ' call to this routine
'
' Note that in this form the last function value, corresponding to the
' argument 'myX', is available as 'myFunc' after the loop exits.
'
' Note that errors in the function evaluation will be raised in the calling
' loop rather than this routine. The only errors seen here are argument
' errors (stepRatio < 1), call count exceeded, and overflow of x (if things
' have gotten completely out of hand).
'
Const R_c As String = "zeroBracket Function"
Const ID_c As String = File_c & R_c
Static s2 As Double  ' holds sign of previous function value between calls

On Error GoTo ErrorHandler

zeroBracket = False  ' default return value is to quit iterating
calls = calls + 1&
If calls > callMax Then Err.Raise 7334&, ID_c, "Call limit exceeded"
If stepRatio < 1# Then Err.Raise 7335&, ID_c, "stepRatio < 1.0"
If 1& = calls Then  ' this is first call; initialize previous-value sign
  s2 = Sgn(f)  ' we only care about the sign of the function
  If 0# = s2 Then  ' exact zero found
    dx = 0#  ' return with zero interval
  Else  ' set up for next call
    x = x + dx  ' step onward
    dx = dx * stepRatio  ' make step larger
    zeroBracket = True  ' tell calling loop to keep iterating
  End If
Else  ' this is call 2, 3, 4, ... so compare sign to previous sign
  Dim s1 As Double
  s1 = s2  ' save previous sign
  s2 = Sgn(f)  ' we only care about the sign of the function
  If 0# = s2 Then  ' exact zero found
    dx = 0#  ' return with zero interval
  ElseIf -1# = s1 * s2 Then  ' function's zero has been bracketed
    dx = -dx / stepRatio  ' adjust so zero between x and x+dx
  Else  ' zero not found or bracketed; adjust step size and try again
    x = x + dx  ' step onward
    dx = dx * stepRatio  ' make step larger
    zeroBracket = True  ' tell calling loop to keep iterating
  End If
End If
Exit Function  '----------------------------------------------------------------

ErrorHandler:  '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
Dim erDes As String
erDes = Err.Description & " in " & R_c & vbLf & _
  "Input on this call:" & vbLf & _
  "f = " & f & "  x = " & x & vbLf & _
  "dx = " & dx & "  stepRatio = " & stepRatio & vbLf & _
  "calls = " & calls & "  callMax = " & callMax & vbLf & _
  "Problem in " & ID_c
Err.Raise Err.Number, ID_c, erDes
Resume  ' if debugging, set Next Statement here and F8 back to error point
' or, set Next Statement on next line and F8 to print error text
Debug.Print "Run-time error '" & Err.Number & "':" & vbLf & erDes
End Function

'===============================================================================
Public Sub zeroBracketUnitTest()
Attribute zeroBracketUnitTest.VB_Description = "Tests of proper operation, with output to the Immediate Window."
' Tests of proper operation of the zero-bracketing routine. These tests were
' written to support "test-driven development" of the code - Google it!
' In Excel, put cursor here and hit F5. In VB6, paste routine's name into
' Immediate Window and hit Enter. Results in Immediate Window (Ctrl-G).
Dim f As Double, x As Double, dx As Double, calls As Long

' this case will cause an overflow error and halt execution - uncomment to run
'x = 1#
'dx = 1#
'Debug.Print vbCrLf & "***** Sqr(x) start "; x; " step "; dx; _
'  " callMax=2000 no zero"
'Debug.Print "  this will halt with an error message:"
'calls = 0&
'Do
'Loop While zeroBracket(Sqr(x), x, dx, calls, 2000&)
  
' this case will cause a stepRatio error and halt execution - uncomment to run
'x = 1#
'dx = 1#
'Debug.Print vbCrLf & "***** Sqr(x) start "; x; " step "; dx; _
'  " stepRatio=0.999"
'Debug.Print "  this will halt with an error message:"
'calls = 0&
'Do
'Loop While zeroBracket(Sqr(x), x, dx, calls, 2000&, 0.999)
  
Debug.Print "##### Unit Tests of ""zeroBracket"" Function ##### " & Now()

x = 1.5  ' Return after 2 evaluations
dx = 0.1
Debug.Print vbCrLf & "***** Cos(x) start "; x; " step "; dx
calls = 0&
Do
  Debug.Print "Working: f "; Cos(x); "x "; x; "dx "; dx; "calls"; calls
Loop While zeroBracket(Cos(x), x, dx, calls)
Debug.Print "Done: f "; Cos(x); "x "; x; "dx "; dx; "calls"; calls
Debug.Print "Zero is between "; x; "and "; x + dx

x = 2.1  ' Return after 4 evaluations; reverse direction
dx = -0.1
Debug.Print vbCrLf & "***** Cos(x) start "; x; " step "; dx
calls = 0&
Do
  f = Cos(x)
  Debug.Print "Working: f "; f; "x "; x; "dx "; dx; "calls"; calls
Loop While zeroBracket(f, x, dx, calls)
Debug.Print "Done: f "; f; "x "; x; "dx "; dx; "calls"; calls
Debug.Print "Zero is between "; x; "and "; x + dx

x = 1#  ' zero on first call
dx = 0.1
Debug.Print vbCrLf & "***** x-1 start "; x; " step "; dx; " zero on first call"
calls = 0&
Do
  Debug.Print "Working: f "; x - 1#; "x "; x; "dx "; dx; "calls"; calls
Loop While zeroBracket(x - 1#, x, dx, calls)
Debug.Print "Done: f "; x - 1#; "x "; x; "dx "; dx; "calls"; calls
Debug.Print "Zero is between "; x; "and "; x + dx

x = 0.9  ' zero on second call
dx = 0.1
Debug.Print vbCrLf & "***** x-1 start "; x; " step "; dx; " zero on second call"
calls = 0&
Do
  Debug.Print "Working: f "; x - 1#; "x "; x; "dx "; dx; "calls"; calls
Loop While zeroBracket(x - 1#, x, dx, calls)
Debug.Print "Done: f "; x - 1#; "x "; x; "dx "; dx; "calls"; calls
Debug.Print "Zero is between "; x; "and "; x + dx

x = 0.7  ' zero on third call
dx = 0.1
Debug.Print vbCrLf & "***** x-1 start "; x; " step "; dx; " zero on third call"
calls = 0&
Do
  Debug.Print "Working: f "; x - 1#; "x "; x; "dx "; dx; "calls"; calls
Loop While zeroBracket(x - 1#, x, dx, calls)
Debug.Print "Done: f "; x - 1#; "x "; x; "dx "; dx; "calls"; calls
Debug.Print "Zero is between "; x; "and "; x + dx

x = 0.7  ' zero on third call
dx = 0.03
Debug.Print vbCrLf & "***** x-1 start "; x; " step "; dx; " stepRatio=1.1"
calls = 0&
Do
  Debug.Print "Working: f "; x - 1#; "x "; x; "dx "; dx; "calls"; calls
Loop While zeroBracket(x - 1#, x, dx, calls, 20&, 1.1)
Debug.Print "Done: f "; x - 1#; "x "; x; "dx "; dx; "calls"; calls
Debug.Print "Zero is between "; x; "and "; x + dx

' this case will cause a too-many-calls error and halt execution
x = 1.002
dx = 1#
Debug.Print vbCrLf & "***** Sqr(x) start "; x; " step "; dx; " no zero"
Debug.Print "  this will halt with an error message:"
calls = 0&
Do
Loop While zeroBracket(Sqr(x), x, dx, calls)
End Sub
  
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


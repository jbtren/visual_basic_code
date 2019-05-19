Attribute VB_Name = "Vector3Dmod"
'
'###############################################################################
'#
'# Visual Basic for Applications (VBA) & VB6 Module "Vector3dMod"
'# Saved in file file "Vector3dMod.bas"
'#
'# Supports operations with 3D Cartesian vectors.
'#
'# Devised and coded by John Trenholme - Started 2012-12-02
'#
'# Exports the Type:
'#   Vector3d
'#
'# Exports the routines:
'#   Function copy
'#   Function cross
'#   Function diff
'#   Function dot
'#   Function makeV3
'#   Function neg
'#   Sub negate
'#   Function norm
'#   Function sToV3
'#   Function sum
'#   Function unit
'#   Sub unitize
'#   Function v3toS
'#   Function vAngle
'#
'###############################################################################

Option Base 0          ' array base value when not specified - the default
Option Compare Binary  ' string comparison based on Asc(char) - the default
Option Explicit        ' forces explicit variable declaration - changes default

' Module-global Const values (convention: start with upper-case; suffix "_")

Private Const Version_ As String = "2013-10-28"
Private Const File_ As String = "Vector3dMod[" & Version_ & "]."
Private Const EOL_ As String = vbNewLine

'#Const UnitTest_C = False
#Const UnitTest_C = True

'########################## Exported User-Defined Type #########################

Public Type Vector3d  ' A 3-dimensional vector in Cartesian coordinates
  x As Double: y As Double: z As Double
End Type

'########################## Exported Routines ##################################

'===============================================================================
Public Function copy(ByRef v As Vector3d) As Vector3d
' Return a copy of the input vector.
copy.x = v.x: copy.y = v.y: copy.z = v.z
End Function

'===============================================================================
Public Function cross(ByRef a As Vector3d, ByRef b As Vector3d) As Vector3d
' Return the cross product of the input vectors.
' Note that cross(a, b) = -cross(b, a).
cross.x = a.y * b.z - a.z * b.y
cross.y = a.z * b.x - a.x * b.z
cross.z = a.x * b.y - a.y * b.x
End Function

'===============================================================================
Public Function diff(ByRef a As Vector3d, ByRef b As Vector3d) As Vector3d
' Return the difference of the input vectors. Note diff(a, b) = -diff(b, a).
diff.x = a.x * b.x: diff.y = a.y * b.y: diff.z = a.z * b.z
End Function

'===============================================================================
Public Function dot(ByRef a As Vector3d, ByRef b As Vector3d) As Double
' Return the dot product of the input vectors.
dot = a.x * b.x + a.y * b.y + a.z * b.z
End Function

'===============================================================================
Public Function makeV3( _
  ByVal x As Double, ByVal y As Double, ByVal z As Double) As Vector3d
' Return a vector with the supplied components
makeV3.x = x: makeV3.y = y: makeV3.z = z
End Function

'===============================================================================
Public Function neg(ByRef v As Vector3d) As Vector3d
' Return the negative of the input vector (component by component).
neg.x = -v.x: neg.y = -v.y: neg.z = -v.z
End Function

'===============================================================================
Public Sub negate(ByRef v As Vector3d)
' Negate the input vector in place (component by component).
v.x = -v.x: v.y = -v.y: v.z = -v.z
End Function

'===============================================================================
Public Function norm(ByRef v As Vector3d) As Double
' Return the Euclidian length of the input vector. It is always >= 0.
Dim maxEl As Double: maxEl = Abs(v.x)
If maxEl < Abs(v.y) Then maxEl = Abs(v.y)
If maxEl < Abs(v.z) Then maxEl = Abs(v.z)
If 0# < maxEl Then
  Dim mult As Double, tx As Double, ty As Double, tz As Double
  mult = 1# / maxEl  ' make fractions of max element; this reduces overflow
  tx = mult * v.x: ty = v.y * mult: tz = v.z * mult  ' all are <= 1
  norm = maxEl * Sqr(tx * tx + ty * ty + tz * tz)
Else
  norm = 0#
End If
End Function

'===============================================================================
Public Function sToV3(ByRef sVec3 As String) As Vector3d
' Convert a String such as "[1.2,-3.4,5.6]" to a Vector3d.
Const ID_C As String = File_ & "sToV3"
Dim s As String: s = Trim$(sVec3)
If (Left$(s, 1&) <> "[") Or (Right$(s, 1&) <> "]") Then  ' missing bracket(s)
  Err.Raise 5&, ID_C, _
    Error$(5&) & EOL_ & _
    "Expected bracketed string like ""[xVal,yVal,zVal]"" but got:" & EOL_ & _
    """" & sVec3 & """" & EOL_ & _
    "Problem in " & ID_C
  Resume  ' if debugging, set Next Statement here and F8 back to error point
End If
' trim brackets off end
s = Trim$(Mid$(s, 2&, Len(s) - 2&))
Dim vs() As String: vs = Split(s, ",")
If (2& <> UBound(vs)) Or _
  (Not IsNumeric(vs(0&))) Or _
  (Not IsNumeric(vs(1&))) Or _
  (Not IsNumeric(vs(2&))) Then  ' wrong count, or not numeric
  Err.Raise 5&, ID_C, _
    Error$(5&) & EOL_ & _
    "Expected three comma-delimited numeric values but got:" & EOL_ & _
    """" & sVec3 & """" & EOL_ & _
    "Problem in " & ID_C
  Resume  ' if debugging, set Next Statement here and F8 back to error point
End If
sToV3.x = val(vs(0&)): sToV3.y = val(vs(1&)): sToV3.z = val(vs(2&))
End Function

'===============================================================================
Public Function sum(ByRef a As Vector3d, ByRef b As Vector3d) As Vector3d
' Return the sum of the input vectors.
sum.x = a.x + b.x: sum.y = a.y + b.y: sum.z = a.z + b.z
End Function

'===============================================================================
Public Function unit(ByRef v As Vector3d) As Vector3d
' Return a vector of unit length in the same direction as the input vector.
' See "unitize" to change the length of the input vector, in place.
Const ID_C As String = File_ & "unit"
Dim size As Double: size = norm(v)
If 0# < size Then
  Dim mult As Double: mult = 1# / size
  unit.x = mult * v.x: unit.y = mult * v.y: unit.z = mult * v.z
Else
  Err.Raise 5&, ID_C, _
    "Got null vector; cannot make unit vector" & EOL_ & _
    "Problem in " & ID_C
  Resume  ' if debugging, set Next Statement here and F8 back to error point
End If
End Function

'===============================================================================
Public Sub unitize(ByRef v As Vector3d)
' Normalize the supplied vector, in place, so it has unit length.
' See "unit" to supply a new vector, leaving the input vector unchanged.
Const ID_C As String = File_ & "unitize"
Dim size As Double: size = norm(v)
If 0# < size Then
  Dim mult As Double: mult = 1# / size
  v.x = mult * v.x: v.y = mult * v.y: v.z = mult * v.z
Else
  Err.Raise 5&, ID_C, _
    "Got null vector; cannot make into unit vector" & EOL_ & _
    "Problem in " & ID_C
  Resume  ' if debugging, set Next Statement here and F8 back to error point
End If
End Sub

'===============================================================================
Public Function v3toS(ByRef v As Vector3d) As String
' Convert a Vector3d to a String such as "[1.2,-3.4,5.6]"
v3toS = "[" & v.x & "," & v.y & "," & v.z & "]"
End Function

'===============================================================================
Public Function v3Angle(ByRef a As Vector3d, ByRef b As Vector3d) As Double
' Angle in radians from "a" to "b". 0 <= angle <= Pi
v3Angle = Atan2(norm(cross(a, b)), dot(a, b))
End Function

'########################## Unit Test Routines #################################

#If UnitTest_C Then

'===============================================================================
Public Sub testCross()
Dim c As Vector3d, d As Vector3d
c.x = 1#: c.y = 0#: c.z = 0#
d.x = 0#: d.y = 1#: d.z = 0#
Debug.Print "=== Test of 'cross' Routine ==="
Debug.Print v3toS(c) & " crossed with " & v3toS(d) & " = " & _
  v3toS(cross(c, d))
Debug.Print v3toS(d) & " crossed with " & v3toS(c) & " = " & _
  v3toS(cross(d, c))
c.x = 0#: c.y = 1#: c.z = 0#
d.x = 0#: d.y = 0#: d.z = 1#
Debug.Print v3toS(c) & " crossed with " & v3toS(d) & " = " & _
  v3toS(cross(c, d))
Debug.Print v3toS(d) & " crossed with " & v3toS(c) & " = " & _
  v3toS(cross(d, c))
c.x = 0#: c.y = 0#: c.z = 1#
d.x = 1#: d.y = 0#: d.z = 0#
Debug.Print v3toS(c) & " crossed with " & v3toS(d) & " = " & _
  v3toS(cross(c, d))
Debug.Print v3toS(d) & " crossed with " & v3toS(c) & " = " & _
  v3toS(cross(d, c))
End Sub

'===============================================================================
Public Sub testDot()
Dim c As Vector3d, d As Vector3d
c.x = 1#: c.y = 2#: c.z = 3#
d.x = 2#: d.y = -1#: d.z = 1#
Debug.Print "=== Test of 'dot' Routine ==="
Debug.Print v3toS(c) & " dotted with " & v3toS(d) & " = " & dot(c, d)
End Sub

'===============================================================================
Public Sub testNorm()
Dim c As Vector3d
c.x = 0#: c.y = 0#: c.z = 0#
Debug.Print "=== Test of 'norm' Routine ==="
Debug.Print "Norm of " & v3toS(c) & " = " & norm(c)
c.x = 1#: c.y = 2#: c.z = 2#
Debug.Print "Norm of " & v3toS(c) & " = " & norm(c)
c.x = 1#: c.y = 4#: c.z = 8#
Debug.Print "Norm of " & v3toS(c) & " = " & norm(c)
End Sub

'===============================================================================
Public Sub testUnit()
Dim c As Vector3d
c.x = 10#: c.y = 0#: c.z = 0#
Debug.Print "=== Test of 'unit' Routine ==="
Debug.Print "Unit vector from " & v3toS(c) & " = " & v3toS(unit(c))
c.x = 1#: c.y = 1#: c.z = 1#
Debug.Print "Unit vector from " & v3toS(c) & " = " & v3toS(unit(c)) & _
  " norm = " & norm(unit(c))
End Sub

'===============================================================================
Public Sub testUnitize()
Dim c As Vector3d, d As Vector3d
c.x = 10#: c.y = 0#: c.z = 0#
d = copy(c): unitize d
Debug.Print "=== Test of 'unitize' Routine ==="
Debug.Print "Unit vector from " & v3toS(c) & " = " & v3toS(d)
c.x = 1#: c.y = 1#: c.z = 1#
d = copy(c): unitize d
Debug.Print "Unit vector from " & v3toS(c) & " = " & v3toS(d) & _
  " norm = " & norm(d)
End Sub

'===============================================================================
Public Sub testV3Angle()
Const Pi_ As Double = 3.1415926 + 5.358979324E-08  ' good to the last bit
Dim c As Vector3d, d As Vector3d
c.x = 1#: c.y = 0#: c.z = 0#
d.x = 1#: d.y = 0#: d.z = 0#
Debug.Print "=== Test of 'v3Angle' Routine ==="
Debug.Print "Angle from " & v3toS(c) & " to " & v3toS(d) & " = " & v3Angle(c, d)
c.x = 1#: c.y = 0#: c.z = 0#
d.x = 0#: d.y = 1#: d.z = 0#
Debug.Print "Angle from " & v3toS(c) & " to " & v3toS(d) & " = " & _
  v3Angle(c, d) & " = " & v3Angle(c, d) / Pi_ & " times Pi"
c.x = 1#: c.y = 0#: c.z = 0#
d.x = -1#: d.y = 0#: d.z = 0#
Debug.Print "Angle from " & v3toS(c) & " to " & v3toS(d) & " = " & _
  v3Angle(c, d) & " = " & v3Angle(c, d) / Pi_ & " times Pi"
c.x = 0#: c.y = 1#: c.z = 0#
d.x = 0#: d.y = 0#: d.z = 1#
Debug.Print "Angle from " & v3toS(c) & " to " & v3toS(d) & " = " & _
  v3Angle(c, d) & " = " & v3Angle(c, d) / Pi_ & " times Pi"
c.x = 0#: c.y = 0#: c.z = 1#
d.x = 1#: d.y = 0#: d.z = 0#
Debug.Print "Angle from " & v3toS(c) & " to " & v3toS(d) & " = " & _
  v3Angle(c, d) & " = " & v3Angle(c, d) / Pi_ & " times Pi"
End Sub

#End If

'########################## Module-Private Support Routines ####################

'-------------------------------------------------------------------------------
Private Function Atan2(ByVal y As Double, ByVal x As Double) As Double
' The ANSI standard arctangent of two arguments in "reverse" order. The branch
' cut is just below the negative X axis, so the result is between -Pi and +Pi.
' The "undefined" value Atan2(0,0) is set to 0 without error.
Const ID_C As String = File_ & "Atan2"
Const Pi_ As Double = 3.1415926 + 5.358979324E-08  ' good to the last bit
Const PiOvr2_ As Double = 1.5707963 + 2.67948965E-08  ' good to the last bit
On Error GoTo ErrHandler  ' y / x might overflow (no error on underflow)
If x = 0# Then  ' on the Y axis
  Atan2 = Sgn(y) * PiOvr2_  ' also takes care of 0,0 case since Sgn(0) = 0
ElseIf y = 0# Then  ' on the X axis; return 0 if x > 0, Pi if x < 0
  Atan2 = (1# - Sgn(x)) * PiOvr2_
ElseIf x > 0# Then  ' in +X half-plane; use ordinary Atn
  Atan2 = Atn(y / x)
Else  ' x < 0; extend smoothly into -X half-plane
  Atan2 = Atn(y / x) + Sgn(y) * Pi_  ' gives -Pi just below negative X axis
End If
Exit Function  '----------------------------------------------------------------
ErrHandler:  '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
Err.Raise Err.Number, ID_C, _
  Error$(Err.Number) & " caused by invalid function argument(s)" & EOL_ & _
  "Input values are y = " & y & EOL_ & "x = " & x & EOL_ & _
  "Problem in " & ID_C
Resume  ' if debugging, set Next Statement here and F8 back to error point
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


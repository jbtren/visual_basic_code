Attribute VB_Name = "Combmrg2"
'
'###############################################################################
'#
'# VBA code file "Combmrg2.bas"
'# L'Écuyer's combined multiple recursive random-number method
'#
'# See "Good Parameters & Implementations for Combined Multiple Recursive
'# Random Number Generators" by Pierre L'Écuyer (Long version, May 4, 1998)
'# (this is the generator in Figure I of that paper; it has M8 = 0.68561,
'# M16 = 0.63940, M32 = 0.63359, M40 ~= 0.6336 and M45 ~= 0.6225 where Mnn = 1
'# is perfection)
'#
'# Coded by John Trenholme - started 2008-07-01
'#
'###############################################################################

Option Explicit

Private Const Version_c As String = "2009-05-30"

' these 6 Doubles hold the state of the two multiple recursive integer gen's
' they always contain exact integer values less than 32 bits in precision

' first 3 must be non-negative, less than 4294967087, and not all 0
Private r0 As Double
Private r1 As Double
Private r2 As Double
' second 3 must be non-negative, less than 4294944443, and not all 0
Private s0 As Double
Private s1 As Double
Private s2 As Double

Private initialized As Boolean

' constants used in the generators
Private Const m1 As Double = 4294967087#  ' 2^32 - 209
Private Const a12 As Double = 1403580#  ' note a11 = 0
Private Const a13n As Double = 810728#  ' actual value is negative
Private Const m2 As Double = 4294944443#  ' 2^32 - 22853
Private Const b21 As Double = 527612#  ' note b22 = 0
Private Const b23n As Double = 1370589#  ' actual value is negative
Private Const norm As Double = 1# / (m1 + 1#)  ' 2.32830654929573E-10
Private Const join As Double = 4.65661132131303E-10

'===============================================================================
Public Function Combmrg2Version() As String
Combmrg2Version = Version_c
End Function

'===============================================================================
Public Function rand32() As Double
' This is Pierre L'Écuyer's combined multiple recursive method (COMBMRG2)
' for 32-bit precision. The period of this version is about 2^191 = 3.14E57.
' It needs 6 seeds. It is well-behaved in dimensions up to at least 45.
Dim r As Double, s As Double, t As Double
If Not initialized Then  ' auto-initialize
  rand6Set 12345#, 12345#, 12345#, 12345#, 12345#, 12345#
End If
' Generator 1
r = a12 * r1 - a13n * r0
r = r - m1 * Int(r / m1)
If r < 0# Then r = r + m1
r0 = r1
r1 = r2
r2 = r
' Generator 2
s = b21 * s2 - b23n * s0
s = s - m2 * Int(s / m2)
If s < 0# Then s = s + m2
s0 = s1
s1 = s2
s2 = s
' Combine the outputs and reduce to region 0 < variate < 1
If r > s Then
  rand32 = (r - s) * norm  ' min value is norm = 2.32830654929573E-10
Else
  rand32 = (r - s + m1) * norm  ' max value is m1 * norm = 0.999999999767169
End If
End Function

'===============================================================================
Public Function rand53() As Double
' This is Pierre L'Écuyer's combined multiple recursive method (COMBMRG2)
' for 32-bit precision, with his method of getting 53-bit Doubles added on
' The period of this version is about 2^190 = 1.57E57. It needs 6 seeds.
If Not initialized Then  ' auto-initialize
  rand6Set 12345#, 12345#, 12345#, 12345#, 12345#, 12345#
End If
' Paste the second result onto the bottom of the first result, giving 53 bits
' Since rand32() >= norm, this is always >= 0 (min = 1.08420186369378E-19)
' The constant 'join' has been adjusted so it is as large as possible while
' keeping the maximum result < 1 (max result is 1 - 1.1102230246E-16)
rand53 = rand32() - norm + rand32() * join
End Function

'===============================================================================
Public Sub rand6Get( _
  ByRef f0 As Double, ByRef f1 As Double, ByRef f2 As Double, _
  ByRef f3 As Double, ByRef f4 As Double, ByRef f5 As Double)
' Get the present state of the 6 random integer seeds, returning them by
' changing the values of the ByRef arguments
f0 = r0
f1 = r1
f2 = r2
f3 = s0
f4 = s1
f5 = s2
End Sub

'===============================================================================
Public Sub rand6Set( _
  ByVal f0 As Double, ByVal f1 As Double, ByVal f2 As Double, _
  ByVal f3 As Double, ByVal f4 As Double, ByVal f5 As Double)
' Set the 6 random integer seeds to the supplied values, as adjusted
Dim fa As Double, fb As Double, fc As Double

fa = Int(Abs(f0))  ' switch negative to positive; round down to integer
If fa >= m1 Then fa = 12345#
r0 = fa
fb = Int(Abs(f1))
If fb >= m1 Then fb = 23456#
r1 = fb
fc = Int(Abs(f2))
If fc >= m1 Then fc = 34567#
If (fa <> 0#) Or (fb <> 0#) Or (fc <> 0#) Then
  r2 = fc
Else
  r2 = 45678#
End If

fa = Int(Abs(f3))
If fa >= m2 Then fa = 54321#
s0 = fa
fb = Int(Abs(f4))
If fb >= m2 Then fb = 65432#
s1 = fb
fc = Int(Abs(f5))
If fc >= m2 Then fc = 76543#
If (fa <> 0#) Or (fb <> 0#) Or (fc <> 0#) Then
  s2 = fc
Else
  s2 = 87654#
End If
initialized = True
End Sub

'-------------------------------------------------------------------------------
Private Sub testRand32()
Debug.Print "*** Combmrg2 Pseudo-Random Variate Routine Tests *** " & Now()
Debug.Print "error in m1 = " & m1 - (2# ^ 32# - 209#)
Debug.Print "error in a12 = " & a12 - 1403580#
Debug.Print "error in a2m = " & a13n - 810728#
Debug.Print "error in m2 = " & m2 - (2# ^ 32# - 22853#)
Debug.Print "error in b21 = " & b21 - 527612#
Debug.Print "error in b2m = " & b23n - 1370589#
Debug.Print "norm = " & norm
Debug.Print "relative error in norm = " & norm * (m1 + 1#) - 1#
Debug.Print "join = " & join
Dim big As Double
big = m1 * norm
Debug.Print "largest possible 32-bit variate = " & big
Debug.Print "  which is 1 - " & 1# - big
Debug.Print "smallest possible 32-bit variate = " & norm
big = big - norm + big * join
Debug.Print "largest possible 53-bit variate = " & big
Debug.Print "  which is 1 - " & 1# - big
Debug.Print "smallest possible 53-bit variate = " & norm * join
Dim s As Double
s = 12345#
Debug.Print "setting all 6 seeds to " & s
rand6Set s, s, s, s, s, s
Dim f0 As Double, f1 As Double, f2 As Double, f3 As Double, f4 As Double, _
  f5 As Double
rand6Get f0, f1, f2, f3, f4, f5
Debug.Print "seeds are now "; f0; f1; f2; f3; f4; f5
Dim j As Long, sum As Double
j = 10000000
Debug.Print "testing generation of " & Format$(j, "#,0") & " variates"
sum = 0#
Dim secs As Single
secs = Timer
For j = 1& To j
  sum = sum + rand32()
Next j
secs = Timer - secs
Debug.Print "calculation took " & secs & " seconds"
Debug.Print "sum of variates = " & sum & " (wanted 5001090.95)"
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Attribute VB_Name = "RandMRG32k3a"
'
'###############################################################################
'#
'# VBA code file "RandMRG32k3a.bas"
'#
'# L'Écuyer's combined multiple recursive uniform 32-bit random-number method
'# "MRG32k3a". This method has good randomness and an extremely long cycle
'# length of 2 ^ 191 = 3E57. Also supplied is a 53-bit uniform generator that
'# combines two outputs of the 32-bit generator (cycle length 1.5E57). The
'# method is well-behaved in all dimensions up to at least 45.
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

Private Const Version_c As String = "2014-03-12"

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

Private initialized As Boolean  ' controls behavior when cold-started

' constants used in the generators
Private Const m1 As Double = 4294967087#  ' 2^32 - 209
Private Const a12 As Double = 1403580#  ' note a11 = 0
Private Const a13n As Double = 810728#  ' actual value is negative
Private Const m2 As Double = 4294944443#  ' 2^32 - 22853
Private Const b21 As Double = 527612#  ' note b22 = 0
Private Const b23n As Double = 1370589#  ' actual value is negative
Private Const norm As Double = 1# / (m1 + 1#)  ' 2.32830654929573E-10
Private Const join As Double = 4.65661132131303E-10  ' near 2 ^ -31

'===============================================================================
Public Function RandCombmrg2Version() As String
RandCombmrg2Version = Version_c
End Function

'===============================================================================
Public Function rand32() As Double
' Return a 32-bit-accuracy uniform pseudo-random variate U with 0 < U < 1.
' Minimum returned value is norm = 2.32830654929573E-10
' Maximum returned value is m1 * norm = 0.999999999767169 = 1 - 2.3283065493E-10
' This is Pierre L'Écuyer's combined multiple recursive method (MRG32k3a)
' for 32-bit precision. The period of this version is about 2^191 = 3E57.
' It needs 6 seeds. It is well-behaved in dimensions up to at least 45.
Dim r As Double, s As Double, t As Double
If Not initialized Then  ' auto-initialize
  rand6Set 12345#, 23456#, 34567#, 45678#, 56789#, 67890#
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
  rand32 = (r - s) * norm
Else
  rand32 = (r - s + m1) * norm  ' max value is m1 * norm = 0.999999999767169
End If
End Function

'===============================================================================
Public Function rand53() As Double
' Return a 53-bit accuracy uniform pseudo-random variate U with 0 < U < 1.
' Minimum returned value is norm * join = 1.08420186369378E-19
' Maximum returned value is 1 - 2 ^ -53 = 1 - 1.11022302462516E-16
' This uses Pierre L'Écuyer's combined multiple recursive method (MRG32k3a)
' for 32-bit precision, with his method of getting 53-bit Doubles added on
' The period of this version is about 2^190 = 1.6E57. It needs 6 seeds.
' It is well-behaved in dimensions up to at least 45.
If Not initialized Then  ' auto-initialize
  rand6Set 12345#, 23456#, 34567#, 45678#, 56789#, 67890#
End If
' Paste the second result onto the bottom of the first result, giving 53 bits
' and throwing out the low-order bits of the second call, which are poorer.
' Since rand32() >= norm, this is always >= 0 (min = 1.08420186369378E-19)
' The constant 'join' has been adjusted so it is as large as possible while
' keeping the maximum result < 1 (max result is 1 - 1.1102230246E-16)
rand53 = rand32() - norm + rand32() * join
End Function

'===============================================================================
Public Sub rand6Get( _
  ByRef f0 As Double, ByRef f1 As Double, ByRef f2 As Double, _
  ByRef f3 As Double, ByRef f4 As Double, ByRef f5 As Double)
' Get the present state of the 6 random integer-in-Double seeds, returning them
' by changing the values of the ByRef arguments. That is:
' Dim a As Double, b As Double, c As Double, d As Double, e As Double, f As Double
'   rand6Get a, b, c, d, e, f
If Not initialized Then  ' auto-initialize
  rand6Set 12345#, 23456#, 34567#, 45678#, 56789#, 67890#
End If
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
' Set the 6 random integer-in-Double seeds to the supplied values, as adjusted
' Values should be integers, even though they are passed in Doubles (7#, not 7.5)
' The first 3 (f0, f1, f2) should obey 0 <= f < m1 (m1 = 4294967087); not all 0
' The last 3 (f3, f4, f5) should obey 0 <= f < m2 (m2 = 4294944443); not all 0
'
Dim fa As Double, fb As Double, fc As Double

fa = Int(Abs(f0))  ' silently switch negative to positive; round down to integer
Debug.Assert fa < m1  ' must be less than modulus 1 = 4294967087#
r0 = fa
fb = Int(Abs(f1))
Debug.Assert fb < m1  ' must be less than modulus 1 = 4294967087#
r1 = fb
fc = Int(Abs(f2))
Debug.Assert fc < m1  ' must be less than modulus 1 = 4294967087#
r2 = fc
Debug.Assert (fa <> 0#) Or (fb <> 0#) Or (fc <> 0#)  ' can't all be zero

fa = Int(Abs(f3))
Debug.Assert fa < m2  ' must be less than modulus 2 = 4294944443#
s0 = fa
fb = Int(Abs(f4))
Debug.Assert fb < m2  ' must be less than modulus 2 = 4294944443#
s1 = fb
fc = Int(Abs(f5))
Debug.Assert fc < m2  ' must be less than modulus 2 = 4294944443#
s2 = fc
Debug.Assert (fa <> 0#) Or (fb <> 0#) Or (fc <> 0#)  ' can't all be zero
initialized = True
End Sub

'-------------------------------------------------------------------------------
Private Sub testRand32()
Debug.Print "****** L'Ecuyer MRG32k3a Uniform Pseudo-Random Variate Routine "; _
  "Tests ****** " & Now()
Debug.Print "** check coefficients"
Debug.Print "error in m1 = " & m1 - (2# ^ 32# - 209#)
Debug.Print "error in a12 = " & a12 - 1403580#
Debug.Print "error in a2m = " & a13n - 810728#
Debug.Print "error in m2 = " & m2 - (2# ^ 32# - 22853#)
Debug.Print "error in b21 = " & b21 - 527612#
Debug.Print "error in b2m = " & b23n - 1370589#
Debug.Print "norm = " & norm & " = 2 ^ " & Log(norm) / Log(2#)
Debug.Print "relative error in norm = " & norm * (m1 + 1#) - 1#
Debug.Print "join = " & join
Debug.Print "** check algorithm"
Dim big As Double
big = m1 * norm
Debug.Print "largest possible 32-bit variate = " & big
Debug.Print "  which is 1 - " & 1# - big
Debug.Print "smallest possible 32-bit variate = " & norm
big = big - norm + big * join
Debug.Print "largest possible 53-bit variate = 0.9999 + " & _
  Format$(big - 0.9999, "0.00000000000000000000")
Debug.Print "  which is 1 - " & 1# - big
Debug.Print "smallest possible 53-bit variate = " & norm * join
Debug.Print "** check seed set-get"
Dim f0 As Double, f1 As Double, f2 As Double, f3 As Double, f4 As Double, _
  f5 As Double
rand6Get f0, f1, f2, f3, f4, f5
Debug.Print "seeds are now "; f0; f1; f2; f3; f4; f5
Dim s As Double
s = 12345#
Debug.Print "setting all 6 seeds to " & s
rand6Set s, s, s, s, s, s
rand6Get f0, f1, f2, f3, f4, f5
Debug.Print "seeds are now "; f0; f1; f2; f3; f4; f5
Const Nvars_ As Long = 10000000
Debug.Print "** testing generation of " & Format$(Nvars_, "#,0") & " variates"
Debug.Print "  getting sum - loop contains 'sum = sum + rand32()'"
Dim j As Long, sum As Double
sum = 0#
For j = 1& To Nvars_
  sum = sum + rand32()
Next j
Debug.Print "sum of variates = " & sum & " (L'Ecuyer got 5001090.95 - rounded)"
Debug.Print "  getting speed - loop contains 'Call rand32'"
Dim secs As Single
secs = Timer()
For j = 1& To Nvars_
  Call rand32
Next j
secs = Timer() - secs
Debug.Print "calculation took " & secs & " seconds = " & _
  CSng(1000000000# * secs / Nvars_) & " ns per call"
Debug.Print "~~~~~~~~~~~~~~~~ unit tests complete ~~~~~~~~~~~~~~~~~~"
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

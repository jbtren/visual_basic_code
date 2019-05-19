Attribute VB_Name = "SVD"
'
'                 ad88888ba   8b           d8  88888888ba,
'                d8"     "8b  `8b         d8'  88      `"8b
'                Y8,           `8b       d8'   88        `8b
'                `Y8aaaaa,      `8b     d8'    88         88
'                  `"""""8b,     `8b   d8'     88         88
'                        `8b      `8b d8'      88         8P
'                Y8a     a8P       `888'       88      .a8P
'                 "Y88888P"         `8'        88888888Y"'
'
'###############################################################################
'#
'# Excel Visual Basic for Applications (VBA) or VB6 Module file "SVD.bas"
'#
'# Singular value decomposition main and support routines.
'#
'# Coded by John Trenholme
'#
'# Based on routines in "Numerical Recipes", which seem to be direct
'# translations of routines in EISPACK (see http://www.netlib.org/eispack/).
'#
'# Started 2007-11-19
'#
'# Exports the routines:
'#   Sub svdDecompose
'#   Function svdHypot
'#   Sub svdSolve
'#   Sub svdTrim
'#   Function svdVersion
'#
'###############################################################################

Option Base 0
Option Compare Binary
Option Explicit
Option Private Module  ' Don't allow visibility outside this Excel Project

Public svdCondition As Double  ' singular value min/max ratio after svdDecompose

Private Const c_Version As String = "2008-02-05"
Private Const c_F As String = "SVD|"  ' name of this file, plus separator

Private Const EOL As String = vbNewLine  ' shorthand for end-of-line

#Const UnitTest = False
' #Const UnitTest = True

#Const VBA = True     ' set True in Excel (etc.) VBA ; False in VB6
' #Const VBA = False

'===============================================================================
Sub svdDecompose( _
  ByRef aMat() As Double, _
  ByRef sVec() As Double, _
  ByRef vMat() As Double)
' Given a general real M-by-N input matrix 'aMat', this routine attempts to
' compute a singular value decomposition
'   aMat = uMat * sVec * transpose(vMat)
' by calculating the matrices 'uMat' and 'vMat', and the vector 'sVec'. The
' vector of singular values 'sVec' is actually the elements of a diagonal
' matrix. When this routine returns, the matrix 'uMat' has replaced the input
' matrix 'aMat'. The input matrix 'vMat', which must be a square 2D array with
' N elements each way, is filled in (note you get back 'vMat', not its
' transpose). The input array 'sMat', which must be a 1D array holding N
' elements, is filled with the diagonal elements of the singular value matrix.
' These elements are non-negative, but they have no particular magnitude order.
' Linear problems aMat * x = rhs can then be solved for 'x' by calling
' "svdSolve". Note that this routine adapts to the input dimensions of 'aMat',
' so you can use any base you want as long as it's the same for all the arrays.

Const c_R As String = "svdDecompose"  ' name of this routine
' the global quantity RV will have been set by the highest-level routine
If RV Then On Error GoTo TidyUp  ' if release version, set to report problem

' get dimensions (M,N) of input matrix 'aMat'
Dim mLo As Long, mHi As Long  ' "rows"
mLo = LBound(aMat, 1&)
mHi = UBound(aMat, 1&)
Dim nLo As Long, nHi As Long  ' "columns"
nLo = LBound(aMat, 2&)
nHi = UBound(aMat, 2&)

If mLo <> nLo Then  ' lower bounds do not match - we don't allow that
  Err.Raise -7333&, c_F & c_R, _
    "Lower bounds mLo = " & mLo & " and nLo = " & nLo & " not equal"
End If

' allocate temporary vector to hold off-diagonal elements
Dim tmp() As Double
ReDim tmp(nLo To nHi)

' Householder reduction to bidiagonal form
Dim f As Double, g As Double, h As Double, norm As Double, size As Double, _
  sizeInv As Double, s As Double, t As Double
Dim jN As Long, jNplus1 As Long, j As Long, k As Long, kM As Long, kN As Long
g = 0#
size = 0#
norm = 0#
For jN = nLo To nHi
  jNplus1 = jN + 1&
  tmp(jN) = size * g
  g = 0#
  size = 0#
  s = 0#
  If jN <= mHi Then  ' if M < N, the "extra" values will be zero
    For kM = jN To mHi
      size = size + Abs(aMat(kM, jN))  ' accumulate L1 norm of column portion
    Next kM
    ' skip transformation if elements were zero to machine precision
    If size <> 0# Then
      sizeInv = 1# / size
      For kM = jN To mHi
        t = aMat(kM, jN) * sizeInv  ' use scaled elements for transformation
        aMat(kM, jN) = t
        s = s + t * t  ' accumulate squares into "sigma"
      Next kM
      g = Sqr(s)
      f = aMat(jN, jN)
      If f >= 0# Then g = -g  ' force opposite of f's sign onto g
      ' with the sign choice above, both terms below have the same sign
      h = f * g - s
      aMat(jN, jN) = f - g
      For kN = jNplus1 To nHi
        s = 0#
        For kM = jN To mHi
          s = s + aMat(kM, jN) * aMat(kM, kN)
        Next kM
        f = s / h
        For kM = jN To mHi
          aMat(kM, kN) = aMat(kM, kN) + f * aMat(kM, jN)
        Next kM
      Next kN
      For kM = jN To mHi
        aMat(kM, jN) = size * aMat(kM, jN)
      Next kM
    End If
  End If
  sVec(jN) = size * g  ' initialize this singular value
  g = 0#
  size = 0#
  s = 0#
  If (jN <= mHi) And (jN <> nHi) Then
    For kN = jNplus1 To nHi
      size = size + Abs(aMat(jN, kN))
    Next kN
    If size <> 0# Then
      sizeInv = 1# / size
      For kN = jNplus1 To nHi
        t = aMat(jN, kN) * sizeInv
        aMat(jN, kN) = t
        s = s + t * t
      Next kN
      g = Sqr(s)
      f = aMat(jN, jNplus1)
      If f >= 0# Then g = -g  ' force opposite of f's sign onto g
      h = f * g - s
      aMat(jN, jNplus1) = f - g
      ' form elements of "p" in unused portion of 'tmp'
      t = 1# / h
      For kN = jNplus1 To nHi
        tmp(kN) = aMat(jN, kN) * t
      Next kN
      For j = jNplus1 To mHi  ' skipped when jN = mHi
        s = 0#
        For kN = jNplus1 To nHi
          s = s + aMat(j, kN) * aMat(jN, kN)
        Next kN
        For kN = jNplus1 To nHi
          aMat(j, kN) = aMat(j, kN) + s * tmp(kN)
        Next kN
      Next j
      For kN = jNplus1 To nHi
        aMat(jN, kN) = size * aMat(jN, kN)
      Next kN
    End If
  End If
  t = Abs(sVec(jN)) + Abs(tmp(jN))
  If norm < t Then norm = t  ' update max of Abs sum
Next jN

' accumulate right-hand Householder transformations

For jN = nHi To nLo Step -1&
  If jN < nHi Then
    If g <> 0# Then
      For j = jNplus1 To nHi
        ' double division avoids possible underflow
        vMat(j, jN) = (aMat(jN, j) / aMat(jN, jNplus1)) / g
      Next j
      For j = jNplus1 To nHi
        s = 0#
        For k = jNplus1 To nHi
          s = s + aMat(jN, k) * vMat(k, j)
        Next k
        For k = jNplus1 To nHi
          vMat(k, j) = vMat(k, j) + s * vMat(k, jN)
        Next k
      Next j
    End If
    ' clear out row & column below here
    For j = jNplus1 To nHi
      vMat(jN, j) = 0#
      vMat(j, jN) = 0#
    Next j
  End If
  vMat(jN, jN) = 1#  ' diagonal element
  g = tmp(jN)
  jNplus1 = jN  ' save this value of jN, for use after loop exit
Next jN

' accumulate left-hand Householder transformations

Dim jMN As Long, jMNplus1 As Long, mnMin As Long
If mHi < nHi Then mnMin = mHi Else mnMin = nHi
For jMN = mnMin To mLo Step -1&
  jMNplus1 = jMN + 1&
  g = sVec(jMN)
  For j = jMNplus1 To nHi
    aMat(jMN, j) = 0#
  Next j
  If g <> 0# Then
    t = 1# / g  ' common factor
    For j = jMNplus1 To nHi
      s = 0#
      For k = jMNplus1 To mHi
        s = s + aMat(k, jMN) * aMat(k, j)
      Next k
      ' this form avoids possible underflow
      f = (s / aMat(jMN, jMN)) * t
      For k = jMN To mHi
        aMat(k, j) = aMat(k, j) + f * aMat(k, jMN)
      Next k
    Next j
    For j = jMN To mHi
      aMat(j, jMN) = aMat(j, jMN) * t
    Next j
  Else
    For j = jMN To mHi
      aMat(j, jMN) = 0#
    Next j
  End If
  aMat(jMN, jMN) = aMat(jMN, jMN) + 1#
Next jMN

' diagonalization of the bidiagonal form by QR transformations
' loop over singular values, and over allowed iteration count

Dim flag As Boolean
Dim c As Double, x As Double, y As Double, z As Double

Dim iteration As Long
Dim jM As Long, jNless1 As Long, kNless1 As Long, jk As Long
For jN = nHi To nLo Step -1&
  Const IterationMax As Long = 30&  ' max iterations (this is quite generous)
  For iteration = 1& To IterationMax
    ' test for splitting
    flag = True
    For kN = jN To nLo Step -1&
      kNless1 = kN - 1&
      ' because tmp(nLo) is always zero, loop will always quit via "Exit For"
      If ((Abs(tmp(kN)) + norm) = norm) Then
        flag = False
        Exit For
      End If
      If ((Abs(sVec(kNless1)) + norm) = norm) Then Exit For
    Next kN
    ' now kN and kNless1 hold saved values

    If flag Then
      ' cancellation of tmp(kN) if kN > nLo
      c = 0#
      s = 1#
      For k = kN To jN
        f = s * tmp(k)
        If ((Abs(f) + norm) = norm) Then Exit For
        g = sVec(k)
        h = svdHypot(f, g)
        sVec(k) = h
        h = 1# / h
        c = g * h
        s = -f * h
        For jM = mLo To mHi
          y = aMat(jM, kNless1)
          z = aMat(jM, k)
          aMat(jM, kNless1) = (y * c) + (z * s)
          aMat(jM, k) = (z * c) - (y * s)
        Next jM
      Next k
    End If

    ' test for convergence
    z = sVec(jN)
    If kN = jN Then
      If z < 0# Then  ' negative singular value; swap sign of it & column
        sVec(jN) = -z
        For j = nLo To nHi
          vMat(j, jN) = -vMat(j, jN)
        Next j
      End If
      Exit For
    End If
    If iteration = IterationMax Then
      Err.Raise -7334&, c_F & c_R, _
        "Max iteration count " & IterationMax & " exceeded"
    End If

    ' shift from bottom 2 by 2 minor
    x = sVec(kN)
    jNless1 = jN - 1&
    y = sVec(jNless1)
    g = tmp(jNless1)
    h = tmp(jN)
    f = ((y - z) * (y + z) + (g - h) * (g + h)) / (2# * h * y)
    g = svdHypot(f, 1#)
    If f < 0# Then g = -g
    f = ((x - z) * (x + z) + h * ((y / (f + g)) - h)) / x

    ' next QR transformation.
    c = 1#
    s = 1#
    For j = kN To jNless1
      k = j + 1&
      g = tmp(k)
      y = sVec(k)
      h = s * g
      g = c * g
      z = svdHypot(f, h)
      tmp(j) = z
      c = f / z
      s = h / z
      f = (x * c) + (g * s)
      g = g * c - x * s
      h = y * s
      y = y * c
      For jk = nLo To nHi
        x = vMat(jk, j)
        z = vMat(jk, k)
        vMat(jk, j) = (z * s) + (x * c)
        vMat(jk, k) = (z * c) - (x * s)
      Next jk
      z = svdHypot(f, h)
      sVec(j) = z
      ' rotation can be arbitrary if z = 0
      If z <> 0# Then
        t = 1# / z
        c = f * t
        s = h * t
      End If
      f = (c * g) + (s * y)
      x = c * y - s * g
      For jk = mLo To mHi
        y = aMat(jk, j)
        z = aMat(jk, k)
        aMat(jk, j) = (z * s) + (y * c)
        aMat(jk, k) = (z * c) - (y * s)
      Next jk
    Next j
    tmp(kN) = 0#
    tmp(jN) = f
    sVec(jN) = x
  Next iteration
Next jN

' find the condition number (min singular value / max singular value)
' make it publicly available after this routine runs
y = sVec(nLo) ' max
z = y ' min
For jN = nLo + 1& To nHi
  x = sVec(jN)
  If y < x Then y = x Else If z > x Then z = x
Next jN
svdCondition = z / y

TidyUp:
' release allocated memory, either after normal operation or after an error
Erase tmp

If Err.Number <> 0& Then
  ' there was an error & we are in Release-Version mode
  ' all allocated resources (temp. arrays, files, ...) have been released
  Dim errNum As Long, errDsc As String  ' save Err object properties
  errNum = Err.Number
  errDsc = Err.Description

  On Error GoTo 0  ' avoid recursion; clears the Err object
  ' re-raise the error, but with added info for user
  ' this routine calls no others, so we don't check for called-routine errors
  errDsc = errDsc & " in routine " & c_F & c_R & EOL & _
  "Dimensions: mLo = " & mLo & "  mHi = " & mHi & _
  "  nLo = " & nLo & "  nHi = " & nHi
  Err.Raise Abs(errNum), c_F & c_R, errDsc
End If

End Sub

'===============================================================================
Function svdHypot( _
  ByVal a As Double, _
  ByVal b As Double) _
As Double
' This returns SQR(a^2 + b^2) without unnecessary overflow or underflow.

Const c_R As String = "svdHypot"  ' name of this routine
' the global quantity RV will have been set by the highest-level routine
If RV Then On Error GoTo Fail  ' if release version, set to report problem

Dim AbsA As Double, AbsB As Double
AbsA = Abs(a)
AbsB = Abs(b)
If AbsA > AbsB Then
  AbsB = AbsB / AbsA  ' this is <= 1
  svdHypot = AbsA * Sqr(1# + AbsB * AbsB)
ElseIf AbsB = 0# Then  ' both inputs were zero; avoid division by either
  svdHypot = 0#
Else
  AbsA = AbsA / AbsB  ' this is <= 1
  svdHypot = AbsB * Sqr(1# + AbsA * AbsA)
End If
Exit Function

Fail:
' simple lowest-level error handler
' routine calls no other routines and allocates no resources that need disposal
Err.Raise Err.Number, c_F & c_R, _
  Err.Description & " in routine " & c_F & c_R & EOL & _
  "Arguments are: a = " & a & ", b = " & b
End Function

'===============================================================================
 Sub svdSolve( _
   ByRef uMat() As Double, _
   ByRef sVec() As Double, _
   ByRef vMat() As Double, _
   ByRef rhs() As Double, _
   ByRef xVec() As Double)
' Solves the linear system A*xVec = rhs of M equations in N unknowns for the vector
' 'xVec' of length N, given a right-hand-side vector 'rhs' and the inverse of A as
' represented by the results 'uMat', 'vMat' & 'sVec' of "svdDecompose". The
' solution is
'   xVec = vMat * inverse(sVec) * transpose(uMat) * rhs
' The output will be in the vector 'xVec', and no inputs except 'xVec' get changed.
' You must supply input arrays, and an output array, with consistent dimensions.
' Say the matrix A() that produced 'uMat', 'vMat' & 'sVec' had the dimensions
' A(1& To m, 1& To n). The input array dimensions then would have to be:
'   Dim uMat(1& To m, 1& To n), sVec(1& To n), vMat(1& To n, 1& To n),
'     rhs(1& To n)
' The output array dimensions would have to be xVec(1& to n).
' Note, however, that this routine adapts to the input dimensions of 'u', so
' you can use any base you want as long as it's the same for all the arrays.
' The user is expected to have set "small" singular values in w to zero
' before calling this routine. This can be done using the "svdTrim" routine.

Const c_R As String = "svdSolve"  ' name of this routine
' the global quantity RV will have been set by the highest-level routine
If RV Then On Error GoTo TidyUp  ' if release version, set to report problem

' get dimensions (m,n) of original matrix that was replaced by 'u'
Dim mLo As Long, mHi As Long
mLo = LBound(uMat, 1&)
mHi = UBound(uMat, 1&)
Dim nLo As Long, nHi As Long
nLo = LBound(uMat, 2&)
nHi = UBound(uMat, 2&)
If mLo <> nLo Then  ' lower bounds do not match - we don't allow that
  Err.Raise -7333&, c_F & c_R, _
    "Lower bounds mLo = " & mLo & " and nLo = " & nLo & " not equal"
End If

' allocate temporary vector
Dim tmp() As Double
ReDim tmp(nLo To nHi)  ' values initialize to 0

' calculate inverse(w) * transpose(u) * rhs
Dim jM As Long, jN As Long
Dim sum As Double
For jN = nLo To nHi
  If sVec(jN) <> 0# Then  ' non-zero singular value, so evaluate
    sum = uMat(mLo, jN) * rhs(mLo)
    For jM = mLo + 1& To mHi
      sum = sum + uMat(jM, jN) * rhs(jM)
    Next jM
    tmp(jN) = sum / sVec(jN)
  End If
Next jN

' multiply vMat * tmp to get result
Dim kN As Long
For jN = nLo To nHi
  sum = vMat(jN, nLo) * tmp(nLo)
  For kN = nLo + 1& To nHi
    sum = sum + vMat(jN, kN) * tmp(kN)
  Next kN
  xVec(jN) = sum  ' set this component of the result into the output vector
Next jN

TidyUp:
' release allocated memory, either after normal operation or after an error
Erase tmp

If Err.Number <> 0& Then  ' there was an error & we are in Release-Version mode
  Dim errNum As Long, errDsc As String  ' save Err object properties
  errNum = Err.Number
  errDsc = Err.Description

  On Error GoTo 0  ' avoid recursion; clears the Err object
  ' re-raise the error, but with added info for user
  ' this routine calls no others, so we don't check for called-routine errors
  errDsc = errDsc & " in routine " & c_F & c_R & EOL & _
  "Dimensions: mLo = " & mLo & "  mHi = " & mHi & _
  "  nLo = " & nLo & "  nHi = " & nHi
  Err.Raise Abs(errNum), c_F & c_R, errDsc
End If

End Sub

'===============================================================================
Public Sub svdTrim( _
  ByVal sizeRatio As Double, _
  ByRef vectorInOut() As Double)
' This routine sets all elements of the input-output vector that are smaller in
' magnitude than 'sizeRatio' times the largest element to zero. It can be used
' to trim off singular values that are "too small". If you have no idea what
' ratio to use, try 1E-6 and see what happens. Although singular values from
' "svdDecompose" are non-negative, this routine works with negative values.

Const c_R As String = "svdTrim"  ' name of this routine
' the global quantity RV will have been set by the highest-level routine
If RV Then On Error GoTo Fail  ' if release version, set to report problem

' get the dimension limits of the input vector
Dim nLo As Long, nHi As Long
nLo = LBound(vectorInOut)
nHi = UBound(vectorInOut)

Dim j As Long, vMax As Double
vMax = Abs(vectorInOut(nLo))
For j = nLo + 1& To nHi
  If vMax < Abs(vectorInOut(j)) Then vMax = Abs(vectorInOut(j))
Next j
Dim vMin As Double
vMin = vMax * sizeRatio
Const TinyDbl As Double = 2.2250738585072E-308 + 1.48219693752374E-323
If vMin < TinyDbl Then vMin = TinyDbl  ' avoid underflow, or negative input
For j = nLo To nHi
  If Abs(vectorInOut(j)) < vMin Then vectorInOut(j) = 0#
Next j
Exit Sub

Fail:
' simple lowest-level error handler
' routine calls no other routines and allocates no resources that need disposal
Err.Raise Err.Number, c_F & c_R, _
  Err.Description & " in routine " & c_F & c_R & EOL & _
  "Arguments are: sizeRatio = " & sizeRatio & _
  ", 2nd arg of type " & TypeName(vectorInOut)
End Sub

'===============================================================================
Public Function svdVersion() As String
' Returns date of last revision as a string in the format "YYYY-MM-DD"
svdVersion = c_Version
End Function

'*******************************************************************************

#If UnitTest Then

'-------------------------------------------------------------------------------
Public Sub svdUnitTest()
' This does some simple sanity checks on the SVD routines. Full-up testing of
' matrix-solver routines is a lengthy and difficult process, and we rely on the
' authorities we stole from to assure the quality of the algorithms.
'
' Unit-test output goes to file, & to Immediate window if in VB6 or VBA Editor.
'
' To run this routine from VBA, put the cursor somewhere in it and hit F5.
'
' To run this routine from VB6, enter "svdUnitTest" in the Immediate window.
' (If the Immediate window is not open, use View... or Ctrl-G to open it.)

Dim path As String
#If VBA Then
  path = ThisWorkbook.path  ' if VBA under Excel
#Else
  path = App.path  ' if VB6
#End If

Open path & "\svdUnitTest.txt" For Output As #outFile()
teeOut String$(80&, "*")
teeOut "***** SVD Unit Test"
teeOut String$(80&, "*")
teeOut "***** Version " & svdVersion() & " by John Trenholme"
teeOut "***** Now: " & Date & " " & Time
teeOut String$(80&, "*")
teeOut

Dim m1 As Long, m2 As Long, n1 As Long, n2 As Long
m1 = 1&: m2 = 12&: n1 = 1&: n2 = 8&

Dim matTest As Long, matType As String
Dim aMat() As Double
ReDim aMat(m1 To m2, n1 To n2)
Dim uMat() As Double, z() As Double
Dim vMat() As Double, sVec() As Double
Dim x() As Double, y() As Double, rhs() As Double
ReDim rhs(m1 To m2), x(n1 To n2), y(n1 To n2)
Dim j As Long, k As Long
Dim sMax As Double, sMin As Double
Dim sum As Double, t As Double, vile As Double, worst As Double
vile = 0#
For matTest = 1& To 999&

  If matTest = 1& Then
    For j = m1 To m2
      For k = n1 To n2
        aMat(j, k) = 2# * Rnd() - 1#
      Next k
    Next j
    matType = "uniform random -1 < r < +1"
  ElseIf matTest = 2& Then
    For j = m1 To m2
      For k = n1 To n2
        aMat(j, k) = Exp(-30# * Rnd())
      Next k
    Next j
    matType = "random exponential exp(-30*r)"
  ElseIf matTest = 3& Then
    For j = m1 To m2
      For k = n1 To n2
        aMat(j, k) = 1# / (j + k - 1&)
      Next k
    Next j
    matType = "Hilbert"
  Else
    Exit For
  End If
  
  For k = n1 To n2
    x(k) = 2# * Rnd() - 1#
  Next k
  For j = m1 To m2
    sum = 0#
    For k = n1 To n2
      sum = sum + aMat(j, k) * x(k)
    Next k
    rhs(j) = sum
  Next j
  
  teeOut "===== Initial matrix: " & matType & " ====="
  printMatrix aMat
  teeOut
  
  ReDim uMat(m1 To m2, n1 To n2), z(m1 To m2, n1 To n2)
  uMat = aMat
  
  ReDim vMat(n1 To n2, n1 To n2), sVec(n1 To n2)
   Dim mm As Integer, nn As Integer
   mm = m1
   nn = n1
  svdDecompose uMat, sVec, vMat

  sMax = sVec(n1)
  sMin = sMax
  For j = n1 + 1& To n2
    If sMax < sVec(j) Then sMax = sVec(j)
    If sMin > sVec(j) Then sMin = sVec(j)
  Next j
  teeOut "SV max: " & CSng(sMax) & "  SV min: " & CSng(sMin) & _
    "  Condition: " & CSng(sMin / sMax)
  teeOut

  teeOut "Output matrix uMat:"
  printMatrix uMat
  teeOut
  
  teeOut "Output vector sVec:"
  printVector sVec
  teeOut
  
  teeOut "Output matrix vMat:"
  printMatrix vMat
  teeOut
  
  Dim i As Long
  For i = m1 To m2
    For j = n1 To n2
      sum = 0#
      For k = n1 To n2
        sum = sum + uMat(i, k) * sVec(k) * vMat(j, k)
      Next k
      z(i, j) = sum
    Next j
  Next i
  
  teeOut "Reconstructed matrix (should equal A):"
  printMatrix z
  teeOut

  worst = 0#
  Const Tweak As Double = 10000000000000#
  For j = m1 To m2
    For k = n1 To n2
      t = (z(j, k) - aMat(j, k)) * Tweak
      z(j, k) = t
      If Abs(worst) < Abs(t) Then worst = t
    Next k
  Next j
  worst = worst / Tweak
  If Abs(vile) < Abs(worst) Then vile = worst
  
  teeOut "Reconstructed matrix minus A, times " & _
    Format$(Tweak, "0.0E-0") & " (should be 0):"
  printMatrix z
  teeOut "Worst deviation: " & CSng(worst)
  teeOut
  
  ReDim z(n1 To n2, n1 To n2)
  worst = 0#
  For j = n1 To n2
    For k = n1 To n2
      sum = 0#
      For i = m1 To m2
        sum = sum + uMat(i, j) * uMat(i, k)
      Next i
      z(j, k) = sum
      If j = k Then t = sum - 1# Else t = sum
      If Abs(worst) < t Then worst = t
    Next k
  Next j
  If Abs(vile) < Abs(worst) Then vile = worst
  
  teeOut "uMat times its transpose (should be a unit matrix):"
  printMatrix z
  teeOut "Worst deviation: " & CSng(worst)
  teeOut
  
  worst = 0#
  For j = n1 To n2
    For k = n1 To n2
      sum = 0#
      For i = n1 To n2
        sum = sum + vMat(i, j) * vMat(i, k)
      Next i
      z(j, k) = sum
      If j = k Then t = sum - 1# Else t = sum
      If Abs(worst) < t Then worst = t
    Next k
  Next j
  If Abs(vile) < Abs(worst) Then vile = worst
  
  teeOut "vMat times its transpose (should be a unit matrix):"
  printMatrix z
  teeOut "Worst deviation: " & CSng(worst)
  teeOut
  
  t = 0.000000000001
  teeOut "Trimming singular values less than " & t & " of max"
  teeOut
  svdTrim t, sVec
  
  svdSolve uMat, sVec, vMat, rhs, y
  
  teeOut "Solving equations - exact answer:"
  printVector x
  teeOut "Result of svdSolve:"
  printVector y
  worst = 0#
  For j = n1 To n2
    t = y(j) - x(j)
    y(j) = t
    If Abs(worst) < Abs(t) Then worst = t
  Next j
  If Abs(vile) < Abs(worst) Then vile = worst
  
  teeOut "Worst deviation: " & CSng(worst)
  teeOut
Next matTest

teeOut "Worst of the worst: " & vile
teeOut

teeOut "#~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
Close #outFile()
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function fmt( _
  ByVal x As Double, _
  Optional ByVal size As Long = 9&) _
As String
' Format the input number into the specified number of spaces
Dim fs As String
fs = "0." & String$(size - 3&, "0")
fmt = Right$(Space$(size - 1&) & Format$(x, fs), size)
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function outFile()
' Return a unit number available for use by a file
Static fileUnit As Integer
If fileUnit = 0 Then fileUnit = FreeFile  ' once only (we may hope)
outFile = fileUnit
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub printMatrix(ByRef M() As Double)
Dim j As Long, k As Long, s As String
For j = LBound(M, 1&) To UBound(M, 1&)
  s = ""
  For k = LBound(M, 2&) To UBound(M, 2&)
    s = s & fmt(M(j, k))
    If k < UBound(M, 2&) Then s = s & " "
  Next k
  teeOut s
Next j
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub printVector(ByRef v() As Double)
Dim j As Long, s As String
s = ""
For j = LBound(v) To UBound(v)
  s = s & fmt(v(j))
  If j < UBound(v) Then s = s & " "
Next j
teeOut s
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub teeOut(Optional ByRef str As String = "")
Debug.Print str  ' only if in Editor
Print #outFile(), str
End Sub

#End If

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


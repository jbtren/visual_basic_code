Attribute VB_Name = "FastFourier"
'
'###############################################################################
'
' Visual Basic for Applications code module "FastFourier"
'
' Holds routines related to the fast discrete Fourier transform (DFT-FFT) for 1D
' complex data of arbitrary length. Includes prime-factor calculation routines.
'
' coded by John Trenholme   started 2013-07-02
'
' Note: only one FFT routine is supplied. It is self-inverse, so that when it
' is applied twice to an array of complex numbers, the original array is
' returned (to within roundoff error).
'
' The input and output is in an array of the exported Complex_t type, so that
' the data type is Double, and the real and imaginary parts are adjacent.
'
' Exports the routines:
'   Function FastFourierVersion
'   Sub fftJBT
'   Function leastPrimeFactor
'   Function primeFactorCount
'   Function primeFactors
'   Function primeFactorsStr
'   Function speedApprox
'
'###############################################################################

Option Base 0          ' array base value when not specified - the default
Option Compare Binary  ' string comparison based on Asc(char) - the default
Option Explicit        ' forces explicit variable declaration - changes default

' Module-global Const values (convention: start with Upper-Case; suffix "_")
Private Const Version_ As String = "2013-07-17"
Private Const File_ As String = "FastFourier[" & Version_ & "]."  ' Module name
Private Const EOL_ As String = vbCrLf  ' handy abbreviation

' Exported User-Defined Type (UDT) (convention: suffix "_t")
Public Type Complex_t  ' a complex number in rectangular coordinates
  x As Double  ' in-phase or real component
  y As Double  ' quadrature or imaginary component
End Type

' Exported global constants that identify errors (convention: suffix "_g")
Global Const Pf_argTooSmall_g As Long = -1&  ' prefix "Pf_" -> primeFactor
Global Const Pf_noSmallFactor_g As Long = -2&

' Quantities local to this Module (convention: suffix "_m")
' Retained between calls; initialize as 0, "" or False (etc.)
Private LB_m As Long  ' lower bound of data array
Private fftLen_m As Long
Private primeFactorCount_m As Long
Private primeFactors_m() As Long
Private work_m() As Complex_t  ' work space equal in size to data

'===============================================================================
Public Function FastFourierVersion(Optional ByVal trigger As Variant) As String
' The date of the latest revision to this Module, as a string like "2019-09-28"
FastFourierVersion = Version_
End Function

'===============================================================================
Public Sub fftJBT(ByRef data() As Complex_t)
' Carry out the self-inverse Fourier transform of the supplied 1D vector of
' complex numbers, overwriting the transform onto the input vector. Applying
' this routine twice will return the original vector, to within roundoff errors.
' Thus this routine is equivalent to an idempotent transformation matrix.
'
' The input array should be dimensioned data(L To N - 1 + L), where N is the
' number of complex values to be transformed and L is arbitrary.
'
' This routine works for data arrays of any length, but it is much, much faster
' when the length is the product of small primes such as 360 = 2^3 * 3^2 * 5.
' The routine "speedApprox" gives guidance on the speed of a specified length.
'
' This routine has the virtues of the ability to work with any data-array length
' and compact code. It is a bit slow, but only about 2X slower than gigantic,
' highly-tuned FFT routines. For casual use, it serves quite well.
' It requires work space equal to the data size, but on modern computers that
' is hardly a problem. To do a 2-D transform, move the rows and columns to
' a 1-D array, transform them one by one, and then put them back.
'
' This Glassman-Ferguson-Cheek style routine is based on the articles:
'
'  J. G. Glassman "A Generalization of the Fast Fourier Transform"
'  IEEE TRansactions on Computers C-19, pp. 105-106 (1970)
'
'  Warren E. Ferguson, Jr. "A Simple Derivation of Glassman's General N Fast
'  Fourier Transform" Comp. & Maths With Appls. 8(6), pp. 401-411 (1982)
'
'  James B. Cheek "The Price Tag for a Fast Fourier Transform on Any Sample Size"
'  Mechanical Systems & Signal Processing 5(5), pp. 357-366 (1991)
'
' As an alternative, consider the closely related method in:
'
'  Carl de Boor "FFT as Nested Multiplication, with a Twist"
'  SIAM J. Sci. Stat. Comput. 1(1), pp. 173-178 (1980)
'
' All Fortran-style array reshaping in implementations of Glassman's method has
' been replaced by explicit 1-D array indexing, and strength reduction has been
' used to speed up the index calculations in the "step" routine
Const ID_C As String = File_ & "fftJBT"
Dim LB As Long: LB = LBound(data)
Dim UB As Long: UB = UBound(data)
Dim datLen As Long: datLen = UB - LB + 1&
If datLen <= 1& Then Exit Sub  ' length-1 transform is the identity
If (fftLen_m <> datLen) Or (LB_m <> LB) Then
  ' input length or array base has changed; update expensive stuff
  fftLen_m = datLen
  LB_m = LB
  ReDim work_m(LB To UB)  ' conform work array to input array
  primeFactors_m = primeFactors(datLen)  ' new list of prime factors
End If
Dim scal As Double: scal = 1# / Sqr(datLen)  ' so applying twice = no change
Dim J As Long
For J = LB_m To UB  ' scale and conjugate the input
  With data(J)
    .x = .x * scal
    .y = -.y * scal  ' this makes exp(2*Pi*i*k/N) -> delta-function at +k
  End With
Next J
Dim inWork As Boolean: inWork = False  ' is transform in work array?
' prime factor manipulation: after * before * current = datLen
' before = not used yet, current = in use now, after = already used
Dim after As Long, before As Long, current As Long  ' prime factors
after = 1&  ' claim we have not yet done any prime factors
before = datLen  ' put all prime factors here
For J = 1& To primeFactorCount()
  current = primeFactors_m(J)  ' get a prime factor
  before = before \ current  ' remove present factor from not-done-yet product
  If inWork Then  ' data is in work array, process it into data array
    fftJBT_step after, before, current, work_m, data ' steps 2, 4, 6, ...
  Else  ' data is in data array, process it into work array
    fftJBT_step after, before, current, data, work_m  ' steps 1, 3, 5, ...
  End If
  after = after * current  ' update the product of prime factors that are done
  inWork = Not inWork  ' toggle processing direction
Next J
If inWork Then data = work_m  ' if result in work array, put it back
End Sub

'===============================================================================
Public Sub fftJBT_step( _
  ByVal after As Long, _
  ByVal before As Long, _
  ByVal current As Long, _
  ByRef inpAr() As Complex_t, _
  ByRef outAr() As Complex_t)
' Carry out the DFT step for the factor in "current" using the values in
' "after" and "before" as well. This routine is called once for each factor.
Const ID_ As String = File_ & "fftJBT_step"
' note: a = after, b = before, c = current, with a * b * c = fftLen_m always
Dim ab As Long: ab = after * before  ' fftLen_m \ current
Dim bc As Long: bc = before * current  ' fftLen_m \ after
Const TwoPi As Double = 6.2831853071795 + 8.65E-14  ' good to the last bit
Dim delAng As Double: delAng = TwoPi / (after * current)
Dim angle As Double: angle = 0#
Dim cosAng As Double, sinAng As Double  ' turn factor (omega) parts
Dim sum As Complex_t  ' holds successive terms in Horner polynomial evaluation
Dim jInp As Long  ' index into input array
Dim jOut As Long: jOut = LB_m  ' index into output array
' initial value for jInBase; the 1& is the first value of "iBef"
Dim jInInit As Long: jInInit = LB_m + bc - before - 1&
Dim iCur As Long, iAft As Long, iBef As Long, iCur2 As Long  ' loop var's
For iCur = 1& To current
  Dim jInBase As Long: jInBase = jInInit  ' (re)set the input-data index base
  For iAft = 1& To after
    ' could get new cos-sin values by rotation, but this gives higher accuracy
    cosAng = Cos(angle): sinAng = Sin(angle)  ' specify turn factor parts
    For iBef = 1& To before  ' start of sum steps forward by 1's
      jInp = jInBase + iBef  ' index into input array
      ' jInp = LB + (iBef-1) + (c - 1) * b + (iAft-1) * b * c
      If 2& < current Then  ' factor > 2; do the full Monty
        sum = inpAr(jInp)  ' initialize the sum
        For iCur2 = 2& To current  ' we just did iCur2 = 1 as a special case
          jInp = jInp - before  ' for other terms in sum, step backwards
          ' jInp = LB + (iBef-1) + (c - iCur2) * b + (iAft-1) * b * c [scrambled]
          Dim dat As Complex_t: dat = inpAr(jInp)  ' next value in data
          With sum  ' multiply sum and add on data value (Horner poly eval)
            Dim xt As Double: xt = .x  ' temp holder for old .x
            .x = .x * cosAng - .y * sinAng + dat.x
            .y = .y * cosAng + xt * sinAng + dat.y
          End With
        Next iCur2  ' sum = Sigma( data(index(J)) * omega^J, J = 0 .. current-1)
      Else  ' the factor is 2 - use simple two-term sum
        sum = inpAr(jInp - before)
        dat = inpAr(jInp)
        With sum
          .x = .x + dat.x * cosAng - dat.y * sinAng
          .y = .y + dat.y * cosAng + dat.x * sinAng
        End With
      End If
      ' jOut = LB + (iBef-1) + (iAft-1) * b + (iCur-1) * a * b [linear map]
      outAr(jOut) = sum  ' put result into output array
      jOut = jOut + 1&  ' output steps forward by 1 for each sum
    Next iBef
    angle = angle + delAng  ' step to next turn factor's angle
    jInBase = jInBase + bc  ' step forward by b * c to next sum data set
  Next iAft
Next iCur
End Sub

'===============================================================================
Sub doIt()  ' just a test harness
Dim N As Long
N = 35&
Debug.Print ">>> Test of"; N; "point DFT  note Sqr(N) ="; Sqr(N)
Dim data() As Complex_t
ReDim data(0& To N - 1&)
Const TwoPi As Double = 6.2831853071795 + 8.65E-14  ' good to the last bit
Dim J As Long, angle As Double
Const turns As Double = 5#
Debug.Print "Input is"; turns; "turn complex exponential"
For J = 0& To N - 1&
  angle = turns * TwoPi * J / N
  data(J).x = Cos(angle)
  data(J).y = Sin(angle)
Next J
Dim K As Long
For K = 1& To 2&
  fftJBT data
  Debug.Print " J"; Tab(6&); " Real"; Tab(28&); " Imag"; Tab(51&); " Magnitude"
  For J = 0& To N - 1&
    Debug.Print J; Tab(6&); data(J).x; Tab(28&); data(J).y; Tab(51&); _
      Sqr(data(J).x ^ 2 + data(J).y ^ 2)
  Next J
Next K
Debug.Print "<<< done"
End Sub

'===============================================================================
Public Function leastPrimeFactor( _
  ByVal theInteger As Long, _
  Optional ByVal returnErrorCode As Boolean = False, _
  Optional ByVal onlySmallPrimes As Boolean = False) _
As Long
' Return smallest non-unity prime factor of the supplied long-integer argument
' To get all factors, call repeatedly, dividing off the return value until = 1
' There will always be Log2(N) or fewer factors of an integer N
' To get an array of all prime factors of an integer, see "primeFactors()"
Const ID_ As String = File_ & "leastPrimeFactor", MaxNdx_ As Long = 171&
Static primes() As Long, isInit As Boolean
If Not isInit Then  ' table is not initialized
  ReDim primes(0& To MaxNdx_)
  Dim I As Long
  For I = 0& To MaxNdx_  ' fill up array with table of low primes
    ' this table covers input integers up to a bit beyond 2^20 = 1048576
    primes(I) = Array( _
      2, 3, 5, 7, 11, 13, 17, 19, 23, 29, 31, 37, 41, 43, 47, 53, 59, 61, 67, _
      71, 73, 79, 83, 89, 97, 101, 103, 107, 109, 113, 127, 131, 137, 139, _
      149, 151, 157, 163, 167, 173, 179, 181, 191, 193, 197, 199, 211, 223, _
      227, 229, 233, 239, 241, 251, 257, 263, 269, 271, 277, 281, 283, 293, _
      307, 311, 313, 317, 331, 337, 347, 349, 353, 359, 367, 373, 379, 383, _
      389, 397, 401, 409, 419, 421, 431, 433, 439, 443, 449, 457, 461, 463, _
      467, 479, 487, 491, 499, 503, 509, 521, 523, 541, 547, 557, 563, 569, _
      571, 577, 587, 593, 599, 601, 607, 613, 617, 619, 631, 641, 643, 647, _
      653, 659, 661, 673, 677, 683, 691, 701, 709, 719, 727, 733, 739, 743, _
      751, 757, 761, 769, 773, 787, 797, 809, 811, 821, 823, 827, 829, 839, _
      853, 857, 859, 863, 877, 881, 883, 887, 907, 911, 919, 929, 937, 941, _
      947, 953, 967, 971, 977, 983, 991, 997, 1009, 1013, 1019, 1021)(I)
  Next I
  isInit = True
End If
' maximum input number that correctly uses values from table is p^2 - 1
' where p is next prime after the largest prime in the table
Const MaxIn_ As Long = 1031& ^ 2 - 1&  ' = 1062960

If 2& > theInteger Then  ' say what?
  If returnErrorCode Then  ' return with an "impossible" value
    leastPrimeFactor = Pf_argTooSmall_g
    GoTo SingleExitPoint_L
  End If
  Err.Raise 5&, ID_, Error$(5&) & EOL_ & _
    "Wanted input argument 2 <= 'theInteger' but got " & theInteger & EOL_ & _
    "Problem in " & ID_  ' error 5 is "Invalid procedure call or argument"
End If

Dim J As Long, K As Long
For I = 0& To MaxNdx_                     ' main loop over small primes in table
  J = primes(I)                           ' get trial small-prime factor
  K = theInteger \ J                      ' quotient of integer division
  If J * K = theInteger Then Exit For     ' prime evenly divided theInteger
Next I

If I <= MaxNdx_ Then  ' a prime factor in the table evenly divided theInteger
  leastPrimeFactor = J
Else  ' ran off top of table
  If theInteger <= MaxIn_ Then  ' if below p^2, no divide -> prime
    leastPrimeFactor = theInteger
  Else  ' primality unknown
    If onlySmallPrimes Then  ' won't expend effort looking for a "large" factor
      If returnErrorCode Then  ' return with an "impossible" value
        leastPrimeFactor = Pf_noSmallFactor_g
        GoTo SingleExitPoint_L
      End If
      Err.Raise 17&, ID_, Error$(17&) & EOL_ & _
        "Input value of " & theInteger & " has no factor <= " & _
          primes(MaxNdx_) & EOL_ & _
        "Problem in " & ID_  ' error 17 is "Can't perform requested operation"
    End If
    Dim jMax As Long  ' use brute-force all-odd-divisors method
    jMax = Int(Sqr(theInteger))  ' no need to test above square root
    ' Debug.Assert jMax * jMax <= theInteger  ' sanity check
    For I = J + 2& To jMax Step 2&  ' odd numbers above largest prime in table
      K = theInteger \ I
      If I * K = theInteger Then Exit For
    Next I
    If I <= jMax Then  ' there is a factor of the square root or less
      leastPrimeFactor = I
    Else  ' no factor found; number must be prime (may take much work)
      leastPrimeFactor = theInteger
    End If
  End If
End If
SingleExitPoint_L:  '*** label ***
End Function

'===============================================================================
Public Function primeFactorCount(Optional ByVal trigger As Variant) As Long
' Return the number of prime factors found in the most recent call to the
' "primeFactors" routine.
primeFactorCount = primeFactorCount_m
End Function

'===============================================================================
Public Function primeFactors( _
  ByVal theInteger As Long, _
  Optional ByVal returnErrorCode As Boolean = False, _
  Optional ByVal onlySmallPrimes As Boolean = False) _
As Long()
' Return an array of the prime factors of an integer, one per array element,
' ordered from least to largest, with duplicates if they exist.
Const ID_ As String = File_ & "primeFactors"
Dim res() As Long
primeFactorCount_m = 0&  ' default value of "could not factor"
If 2& > theInteger Then  ' say what?
  If returnErrorCode Then  ' return with an "impossible" value
    ReDim res(1& To 1&)
    res(1&) = Pf_argTooSmall_g
    GoTo SingleExitPoint_L
  End If
  Err.Raise 5&, ID_, Error$(5&) & EOL_ & _
    "Wanted input argument 2 <= 'theInteger' but got " & theInteger & EOL_ & _
    "Problem in " & ID_  ' error 5 is "Invalid procedure call or argument"
End If
On Error GoTo ErrorHandler_L
Dim I As Long, J As Long, K As Long
J = theInteger
K = 0&
Dim kMax As Long
Const Ln_2 As Double = 0.6931471 + 8.055994531E-08  ' good to the last bit
kMax = Int(Log(theInteger) / Ln_2) + 1&  ' the added 1 is just for caution
ReDim res(1& To kMax)  ' make more than enough room for all factors
Do
  I = leastPrimeFactor(J, returnErrorCode, onlySmallPrimes)
  K = K + 1&
  res(K) = I
  J = J \ I
Loop While J > 1&
ReDim Preserve res(1& To K)  ' shrink array to just hold the factors
primeFactorCount_m = K

SingleExitPoint_L:  '*** label ***
primeFactors = res
Erase res  ' sometimes VB gets confused about auto-erase of dynamic arrays
Exit Function

ErrorHandler_L:  '*** Label ***
Dim errDes As String: errDes = Err.Description
If 17& = Err.Number Then
  ReDim res(1& To 1&)
  res(1&) = Pf_noSmallFactor_g
End If
' supplement text; did error come from a called routine or this routine?
errDes = errDes & EOL_ & _
  "Input integer = " & theInteger & EOL_ & _
  IIf(0& < InStr(errDes, "Problem in"), "Called by ", "Problem in ") & ID_
' re-raise error with this routine's ID as Source, and appended to Description
Err.Raise Err.Number, ID_, errDes
Resume  ' if debugging, set "Next Statement" here and F8 back to error point
End Function

'===============================================================================
Public Function primeFactorsStr( _
  ByVal theInteger As Long, _
  Optional ByVal returnErrorCode As Boolean = False, _
  Optional ByVal onlySmallPrimes As Boolean = False) As String
' Return the prime factors in a comma-delimited string
Dim factors() As Long
factors = primeFactors(theInteger, returnErrorCode, onlySmallPrimes)
If factors(1&) < 2& Then  ' got an error; returned code
  If Pf_argTooSmall_g = factors(1&) Then
    primeFactorsStr = "ERROR: argument = " & Trim$(theInteger) & " too small"
  ElseIf Pf_noSmallFactor_g = factors(1&) Then
    ' see Const MaxIn_ in "leastPrimeFactor" for the magic number
    primeFactorsStr = "ERROR: no factor < 1031 in " & theInteger
  Else  ' undefined error code?!?
    primeFactorsStr = "ERROR: return code = " & factors(1&)
  End If
Else  ' got an array of factors
  Dim res As String
  res = Trim$(factors(1&))
  Dim J As Long
  For J = 2& To primeFactorCount()
    res = res & "," & Trim$(factors(J))
  Next J
  primeFactorsStr = res
End If
End Function

'===============================================================================
Public Function speedApprox(ByVal dataLength As Long) As Double
' Reasonable approximation of ratio of DFT speed to power-of-2 best-case code
' See the article:
'  James B. Cheek "The Price Tag for a Fast Fourier Transform on Any Sample Size"
'  Mechanical Systems & Signal Processing 5(5), pp. 357-366 (1991)
Const A As Double = 1.046, B As Double = 0.528, C As Double = 1.08, _
  D As Double = 0.08  ' timings depend on compiler, CPU, memory, cache, etc.
Dim N As Long: N = dataLength  ' to use Cheek's notation
Dim factors() As Long
factors = primeFactors(N)
Dim M As Long: M = primeFactorCount()
Dim S As Long: S = 0&  ' sum of prime factors
Dim U As Long: U = 0&  ' sum of products of first J prime factors
Dim product As Long: product = 1&
Dim fact As Long
Dim J As Long
For J = 1& To M
  fact = factors(J)
  S = S + fact
  product = product * fact
  U = U + product
Next J
Dim timeCheek As Double, timePwr2 As Double
timeCheek = A * N * M + B * U + C * N * (S - M) + D * S  ' Cheek's result
' now pretend that N is an exact power of 2
Dim P As Double: P = 1.44269504088896 * Log(N)  ' log base 2 = power of 2
timePwr2 = A * N * P + B * 2# * (N - 1#) + C * N * P + D * 2# * P
' result is a number less than or equal to 0.5 (a somewhat arbitrary scaling)
speedApprox = 0.5 * timePwr2 / timeCheek
End Function

'###############################################################################
'#
'# Unit test routines
'#
'###############################################################################

#If True Then  ' include unit-test code
' #If False Then  ' exclude unit-test code

'===============================================================================
Public Sub FFTunitTest()
Const N As Long = 1350&
Dim data() As Complex_t
ReDim data(0& To N - 1&)
Dim res() As Double
ReDim res(0& To N - 1&, 1& To 4&)
Const TwoPi As Double = 6.2831853071795 + 8.65E-14  ' good to the last bit
Const Pi As Double = 3.1415926 + 5.358979324E-08
Const E As Double = 1E-99
Const turns As Double = 20#
Dim J As Long
For J = 0& To N - 1&
  With data(J)
    .x = Sin(Pi * turns * (J + E) / N) / _
      (turns * Tan(Pi * (J + E) / N))
    .y = 0#
  End With
  res(J, 1&) = J
  res(J, 2&) = data(J).x
Next J

fftJBT data

For J = 0& To N - 1&
  With data(J)
    res(J, 3&) = .x
    res(J, 4&) = .y
  End With
Next J

Range("testCase").Resize(N, 4&) = res
End Sub

'===============================================================================
Public Sub leastPrimeFactorUnitTest()
Debug.Print ">>> Unit tests of " & File_ & "leastPrimeFactor at " & now()
Dim I As Long, J As Long, K As Long, M As Long, N As Long, tStart As Single
tStart = Timer()

J = leastPrimeFactor(1&, True)
If J < 2& Then
  Debug.Print "Low-end error code returns "; J; "at N = 1 as expected"
Else
  Stop
End If

Debug.Print "Smallest allowed input 2 ->"; leastPrimeFactor(2&, True)

Debug.Print "Tests of smallest factor of values near long-integer upper limit:"
Debug.Print "Note: primes here are 2147483629 & 2147483647"
Dim x As Double  ' need one beyond Long max to exit loop
For x = 2147483626# To 2147483647#
  Debug.Print Int(x); "->"; leastPrimeFactor(Int(x), True)
Next x

Debug.Print "Tests of all factors of values near prime-table upper limit:"
Debug.Print "Note: primes here are 1062947 & 1062949, and 1062961 = 1031^2"
Dim factors(1& To 21&) As Long
For I = 1062940 To 1062965
  K = I
  M = 0&
  Do
    J = leastPrimeFactor(K, True)
    M = M + 1&
    factors(M) = J
    K = K \ J
  Loop While K > 1&
  Debug.Print I; "->";
  For J = 1& To M
    Debug.Print factors(J);
  Next J
  Debug.Print
Next I

Debug.Print "Test of only-small-factors logic:"
I = 1042441
Debug.Print I; "->"; leastPrimeFactor(I, True, True)
I = 1062961
Debug.Print I; "-> "; leastPrimeFactor(I, True, True)

Dim P As Long, jMax As Long, jMin As Long
M = 1000000
Debug.Print "Testing all factors of"; M; _
  "random inputs 2 <= theInteger <= 1062960"
jMax = 0&
jMin = 2147483647
If Rnd(-1) < 2! Then Randomize Timer()
For I = 1& To M
  J = Int(2# + 1062958.99999999 * Rnd())
  If jMin > J Then jMin = J Else If jMax < J Then jMax = J
  K = J
  P = 1&
  Do
    N = leastPrimeFactor(K)
    P = P * N
    K = K \ N
  Loop While K > 1&
  If P <> J Then Debug.Print "FAIL at"; J: Stop
Next I
Debug.Print "  Input min"; jMin; "  max"; jMax; "  no errors"

M = 100000
Debug.Print "Testing all factors of"; M; _
  "random inputs 1000000 <= theInteger <= 2147483647"
jMax = 0&
jMin = 2147483647
For I = 1& To M
  J = Int(1000000# + 2146483647.9999 * Rnd())
  If jMin > J Then jMin = J Else If jMax < J Then jMax = J
  K = J
  P = 1&
  Do
    N = leastPrimeFactor(K)
    P = P * N
    K = K \ N
  Loop While K > 1&
  If P <> J Then Debug.Print "FAIL at"; J: Stop
Next I
Debug.Print "  Input min"; jMin; "  max"; jMax; "  no errors"

Debug.Print "<<< unit tests complete in"; Round(Timer() - tStart, 3); "seconds"

' force an error
J = leastPrimeFactor(1&)
End Sub

'===============================================================================
Public Sub primeFactorsUnitTest()
Debug.Print ">>> Unit tests of " & File_ & "primeFactors at " & now()
Dim tStart As Single: tStart = Timer()
Dim factors() As Long, J As Long, K As Long
For K = 1024& - 10& To 1024& + 10&
  Debug.Print "Factors of"; K; "=";
  factors = primeFactors(K)
  For J = 1& To UBound(factors)
    Debug.Print factors(J);
  Next J
  Debug.Print
Next K
For K = 1024& ^ 2& - 10& To 1024& ^ 2& + 10&
  Debug.Print "Factors of"; K; "=";
  factors = primeFactors(K)
  For J = 1& To UBound(factors)
    Debug.Print factors(J);
  Next J
  Debug.Print
Next K
For K = 2147483626 To 2147483646
  Debug.Print "Factors of"; K + 1&; "=";
  factors = primeFactors(K + 1&)
  For J = 1& To UBound(factors)
    Debug.Print factors(J);
  Next J
  Debug.Print
Next K
Debug.Print "<<< unit tests complete in"; Round(Timer() - tStart, 3); "seconds"
End Sub

#End If

'~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

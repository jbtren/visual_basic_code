Attribute VB_Name = "Random"
Attribute VB_Description = "Pseudo-random sequence routines of high quality & long cycle length. Devised and coded by John Trenholme."
'
'###############################################################################
'#    _____    ___   _   _  _____    ____   __  __     ____    ___    _____.
'#   |  __ \  / _ \ | \ | ||  __ \  / __ \ |  \/  |   |  _ \  / _ \  / ____|
'#   | |__) || |_| ||  \| || |  | || |  | || \  / |   | |_) || |_| || (___.
'#   |  _  / |  _  || . ` || |  | || |  | || |\/| |   |  _ < |  _  | \___ \
'#   | | \ \ | | | || |\  || |__| || |__| || |  | | _ | |_) || | | | ____) |
'#   |_|  \_\|_| |_||_| \_||_____/  \____/ |_|  |_|(_)|____/ |_| |_||_____/
'#
'# Visual Basic (VB6 or VBA) Module "Random"
'# Saved in text file "Random.bas"
'#
'# Pseudo-Random variate routines.
'#
'# This Module exports the following routines:
'#
'# Function gammaLog
'# Function gammaSign
'# Function RandomVersion
'# Function rndDbl
'# Function rndExponential
'# Function rndGamma
'# Function rndGumbel
'# Function rndLng
'# Function rndLngMax
'# Function rndLogNormal
'# Function rndNormal
'# Function rndPoisson
'# Function rndRicianInten
'# Sub rndSeedGet
'# Function rndSeedGetA
'# Function rndSeedGetB
'# Sub rndSeedSet
'# Function rndSng
'# Function time2Long
'#
'# Sub randomUnitTest  if UnitTest_C is True
'#
'# Devised and coded by John Trenholme
'#
'###############################################################################

Option Base 0          ' array base value when not specified - the default
Option Compare Binary  ' string comparison based on Asc(char) - the default
Option Explicit        ' forces explicit variable declaration - changes default
' Option Private Module  ' no effect in VB6; visible-this-Project-only in VBA

' Module-global Const values (convention: start with upper-case; suffix "_c")
Private Const Version_c As String = "2013-05-26"  ' update manually on each edit
Private Const File_c As String = "Random[" & Version_c & "]."
Private Const InvalidArg_c As Long = 5&  ' "Invalid procedure call or argument"

'-------------------------------------------------------------------------------
' I wrote this module because the built-in pseudo-random function Rnd() in VBA &
' VB6 has too short a cycle length (only 2^24 or 16,777,216 numbers) to be used
' in a serious Monte Carlo job. It also has the serious defect that succesive
' N-tuples lie on planes in the unit hypercube. Also, it has only single-
' precision resolution. A longer-period generator of known high quality, with
' double-precision resolution available, is needed. This module supplies that
' generator. The cycle length is 4.61E18, after which the sequnce repeats. Note,
' however, that Don Knuth recommends that you use only a small fraction of the
' sequence length of any generator, so you may want to quit after (say) 1E15
' values. Some ultra-cautious authors suggest using only the square root of the
' sequence length, which would be around 2 billion here; that is probably much
' too fastidious.
'
' By including this module in your Project, you can replace VB's Rnd() with
' rndSng(), and replace Randomize, with its confusing and inconsistent behavior,
' with rndSeedSet(). Instead of Randomize with no argument, which seeds with
' the system timer, you can use time2Long() as one of the arguments to
' rndSeedSet() if you want different results on every run.
'
' The price of improved quality is that these routines are slower. For example,
' the rndDbl() function is about 4 times slower than Rnd(), and rndSng() is
' only slightly better at about 3.8 times slower. Given this small difference
' in speed, you may as well use the higher-precision rndDbl() variates.
'
' If the name of a Public routine exported by this module conflicts with a
' global name exported from another Module or Form, use the form "Rnd.name"
' instead of "name" (and "Other.name" for the other name). Visual Basic will
' then know which routine is which, since they will be in separate namespaces.
'
' In addition to the Rnd() replacement rndSng(), you also get the double-
' precision pseudo-random sequence rndDbl(), and routines to produce pseudo-
' random sequences drawn from a number of common probability distributions. The
' sequences and supporting routines provided are:
'
'   rndDbl()          uniform double-precision variates  0.0 < rndDbl() < 1.0
'   rndExponential()  exponentially-distributed double-precision variates
'   rndGamma()        gamma-distributed double-precision variates
'   rndGumbel()       Gumbel extreme-value double-precision variates
'   rndLogNormal()    log-normal-distributed double-precision variates
'   rndLng()          uniform long-integer values  0 <= rndLng() <= 2147483562
'   rndLngMax()       the largest value that rndLng() returns (2147483562)
'   rndNormal()       normal (or Gaussian) double-precision variates
'   rndPoisson()      Poisson-distributed long-integer variates
'   rndRicianInten()  Rician intensity double-precision variates
'   rndSng()          uniform single-precision variates  0.0 < rndSng() < 1.0
'
' All the sequences are derived from one underlying Long generator based on the
' merging of two mixed congruential Long sequences. The present point in the
' generator cycle is specified by two Long variables, called seeds (if you only
' want to use one, set the other to any arbitrary constant value; 0 will work
' fine). You set and read the seed values by using the routines:
'
'   rndSeedGet seedA As Long, seedB As Long   ' get both seed values ByRef
'   rndSeedGetA()                             ' get the first seed value
'   rndSeedGetB()                             ' get the second seed value
'   rndSeedSet seedA As Long, seedB As Long   ' set both seed values
'
' Setting the seeds allows you to force a different set of pseudo-random
' variates on different runs, or to repeat a specific pseudo-random sequence
' for debugging or comparison.
'
' Repeating a sequence is simple. At the point where you want to start a
' sequence, set the seeds (or read the present values) and then save the values.
' Then use some number of pseudo-random variates, and later call rndSeedSet
' with the saved seeds. The exact same sequence of pseudo-random variates will
' then be supplied (assuming you make the same calls in the same order). You
' must call rndSeedSet with the saved seeds immediately after any seed save if
' you want the same sequence from rndNormal and rndLogNormal, since they save
' values internally but get reset by rndSeedSet (see example in rndNormal).
'
' Three support routines used in the module are made publicly available:
'
'   gammaLog()     logarithm of absolute value of Gamma function
'   gammaSign()    sign of the Gamma function (+1, -1, or 0 at poles)
'   time2Long()    count of 67th-of-a-second intervals since New Year's Day
'                     this can be used in rndSeedSet to "randomize" results
'                     the way Randomize with no argument does in Visual Basic
'
'-------------------------------------------------------------------------------

' IMPORTANT: set the following constant to False if using with VB6, or to True
' if using with VBA in Excel (if Excel will use functions from here in
' spreadsheet cell formulas). If you leave this True in a VB6 program, execution
' of the routines can be extremely slow. If you leave it at False under Excel,
' spreadsheet cells that use the routines will not be updated properly.
#Const ExcelFunction_C = True
' #Const ExcelFunction_C = False

' Set this True to include unit-test code, False to exclude it.
#Const UnitTest_C = True
' #Const UnitTest_C = False

#If UnitTest_C Then  ' Windows API functions needed for microsecond timing
Private Declare Function QueryPerformanceFrequency _
  Lib "kernel32" (f As Currency) As Boolean
Private Declare Function QueryPerformanceCounter _
  Lib "kernel32" (p As Currency) As Boolean
#End If

' This constant is written as the sum of two parts to maintain full accuracy.
' VB[6|A] will otherwise truncate a digit when module is saved to file & loaded.
' This rude behavior is a known fault of VB.
' In this form, Pi is good to the last bit in IEEE 754 floating point.
Private Const Pi_c As Double = 3.1415926 + 5.358979324E-08

' ===== Details of the pseudo-random generator and routines ====================

' The basis for good pseudo-random number sequence generators is often a good
' integer sequence generator. Fortunately, many superior integer generators are
' available. Some even work with signed 31-bit integers (i.e., Longs), which VB
' is restricted to. In this module, we use a pair of portable pseudo-random
' integer generators for 31-bit signed integer arithmetic that work without
' overflow of intermediate results (another VB restriction). They are of the
' linear congruential type extensively discussed in Knuth, Seminumerical
' Algorithms, 3rd edition, section 3.2.1. The generators used here are slightly
' modified from the example in "Efficient and Portable Random Number Generators"
' by Pierre L'Ecuyer, CACM 31, 742(1988).  The changes are:
'   1) the addition of constants in the evaluations to allow 0 seeds, and to
'      avoid a series of very small values when small seeds are used
'   2) adjustment of the range of the uniform floating-point variates generated
'      from the integers to exclude the end values 0.0 and 1.0
'
' The two generators are combined by modulo subtraction. Their periods are
' mutually prime (in fact, both periods are prime) and so the combined period is
' the product of the two periods, giving about 4.61E18 nearly-31-bit integers
' until the sequence repeats (enough for 1 per nanosecond for more than 145
' years).
'
' The combined integers are in turn used to generate uniformly-distributed
' Single and Double floating-point values, and the Double values are used to
' generate non-uniform continuous and discrete pseudo-random distributions.
' Because of the combination of the two generators, and because of the fact that
' the low-order bits in the resulting sequence are discarded when making Single
' and Double values, there is no lie-on-hyperplane problem.
'
' Note that there are far fewer than 4.6E18 possible output values for the
' various sequences. Therefore, values will be repeated. However, the repeated
' values will be rare and irregularly spaced. For example, rndSng() can supply
' only about 1.7E7 distinct output values, so each value will be repeated about
' 2.7E11 times before the cycle repeats. But on average, you will have to go
' through 1.7E7 sequence values before you see a repeat. If you need fewer
' repeats, use rndDbl(), which supplies about 9.0E15 distinct output values,
' repeated 511 times in the entire sequence.
'
' The "rndXXX" routines pass most, but not all, of Marsaglia's Diehard "torture
' tests." To pass all of them, we would have to add a 4-long Bayes-Durham
' shuffle table, but the resulting improvement for very, very few problems was
' judged not to be worth the slowdown for practically all applications. And,
' nobody's perfect - just remember Don Knuth's maxim (Knuth, Seminumerical
' Algorithms, section 3.6): "Every random number generator will fail in at least
' one application." In fact, I think we can confidently extend failure to an
' infinite number of applications, forming a set of measure zero in application
' space.
'
' The present point in the generator cycle is shared, and advanced, by all the
' sequences. A call to any routine advances the sequence by one step, two steps,
' or (in some cases) more than two steps. Be careful not to disturb an order
' you depend on by calling another random-variate routine.

' ==============================================================================

' Warning! Adjust these constants only if you know what you're doing!
' The modulus values (Ra1_c and Rb1_c) must be mutually prime, and preferably
' individually prime and close to 2^31. The multipliers (Ra2_c and Rb2_c) must
' be carefully chosen to give good results on the spectral test (Knuth 3.3.4).
' The values used here are from L'Ecuyer's CACM article (see above) and have
' been extensively tested by many practitioners.
'
' The first integer generator is R = ( 40014 * R + 1693666889) mod 2147483563
' The modulus is 2^31 - 85, the 5th prime below 2^31.
' The multiplier is a bit less than the square root of the period: a < 46340
' The additive constant must obey:
'   Ra5_c * ( ( Ra1_c - 1) \ Ra4_c) - Ra1_c <= Ra3_c <= Ra2_c + Ra5_c - 1
'   -1658872609 <= Ra3_c <= 52224
' Note that 1693666889 is a prime close to the modulus times (3 + Sqr(3)) / 6,
' or 0.788675, which is an obsolete value suggested by Knuth (Knuth, 1st ed.,
' 3.3.3-39). Knuth now says any non-zero value is good if the modulus is prime.
' Here is a sequence of values produced by this generator:
' 229587626, 1478251139, 0, 1693666889, 1794282181, 1282895644, 2032875953
Private Const Ra1_c As Long = 2147483563  ' modulus
Private Const Ra2_c As Long = 40014  ' multiplier
' We subtract the modulus from the constant, making it negative but in range
Private Const Ra3_c As Long = 1693666889 - Ra1_c  ' added constant = -453816674
' The modulus is "approximately factored" as m = p * Ra2_c + q  (q < Ra2_c)
Private Const Ra4_c As Long = Ra1_c \ Ra2_c  ' p = 53668
Private Const Ra5_c As Long = Ra1_c - Ra2_c * Ra4_c  ' q = 12211

' The second integer generator is R = ( 40692 * R + 858993503) mod 2147483399
' The modulus is 2^31 - 249, the 12th prime below 2^31.
' The multiplier is a bit less than the square root of the period: a < 46340
' The additive constant must obey:
'   Ra5_c * ( ( Ra1_c - 1) \ Ra4_c) - Ra1_c <= Ra3_c <= Ra2_c + Ra5_c - 1
'   -1993220027 <= Rb3_c <= 44482
' The additive constant was chosen as a prime near 0.4 times the modulus.
' Here is a sequence of values produced by this generator:
' 136699883, 1468629129, 0, 858993503, 435332056, 800457904, 63827039
Private Const Rb1_c As Long = 2147483399  ' modulus
Private Const Rb2_c As Long = 40692  ' multiplier
' We subtract the modulus from the constant, making it negative but in range
Private Const Rb3_c As Long = 858993503 - Rb1_c  ' added constant = -1288489896
' The modulus is "approximately factored" as m = p * Rb2_c + q  (q < Rb2_c)
Private Const Rb4_c As Long = Rb1_c \ Rb2_c  ' p = 52774
Private Const Rb5_c As Long = Rb1_c - Rb2_c * Rb4_c  ' q = 3791

' A Long in the range 1 <= Long < Ra1_c is scaled to the valid Single
' range 0! < rndSng < 1! by the following values, using the formula:
'   rndSng = AddS_c + Long * MulS_c
' When Long = 0, the formula must yield that smallest Single which, when
' subtracted from 1, gives a result greater than 0. This is 2^(-24) =
' 5.960464E-08. When Long = Ra1_c - 1, the formula must yield the largest
' possible Single which, when subtracted from 1, gives a result greater than 0.
' This keeps rndSng() > 0!, rndSng() < 1!, 1! - rndSng() > 0! and
' 1! - rndSng() < 1!, which are all required by some algorithms.
' From these requirements, we have:
'   AddS_c = 2! ^ (-24)
'   AddS_c + MulS_c * (Ra1_c - 1&) = 1! - 2! ^ (-24)
' Solving these equations for AddS_c and MulS_c, we get:
Private Const AddS_c As Single = 2# ^ (-24)
Private Const MulS_c As Single = (1# - 2# * 2# ^ (-24)) / (Ra1_c - 1&)
' Note that, with these values, the rndSng() value can be exactly 0.5;
' that can be a problem for algorithms which use rndSng() - 0.5!, but do not
' want that quantity to be zero. Adding 1E-9 or less will fix the problem.

' Two Longs in the range 0 <= Long < Ra1_c are scaled to the valid Double
' range 0# < rndDbl < 1# by the following values, using the formula:
'   rndDbl = AddD_c + Long1 * MulHiD_c + Long2 * MulLoD_c
' When both Longs = 0, the formula must yield the smallest Double which, when
' subtracted from 1, gives a result greater than 0. This is 2^(-53) =
' 1.11022302462516E-16. When both Longs = Ra1_c - 1, the formula must yield the
' largest possible Double which, when subtracted from 1, gives a result greater
' than 0. This keeps rndDbl() > 0#, rndDbl() < 1#, 1# - rndDbl() > 0# and
' 1# - rndDbl() < 1#, which are all required by some algorithms. From these
' requirements, we have:
'   AddD_c = 2# ^ (-53)
'   AddD_c + (MulHiD_c + MulLoD_c) * (Ra1_c - 1&) = 1# - 2# ^ (-53)
' We also require that the values added on by MulLoD_c just "fill in" between
' steps of size MulHiD_c. There should be 2^(2*31-53) = 2^9 = 511 combinations
' of the two Longs leading to each Double value. This leads to the requirement
' that:
'   MulLoD_c = MulHiD_c / Ra1_c
' Solving these equations for AddD_c, MulHiD_c and MulLoD_c, we get:
Private Const AddD_c As Double = 2# ^ (-53)
Private Const MulHiD_c As Double = (1# - 2# * 2# ^ (-53)) / (Ra1_c - 1# / Ra1_c)
Private Const MulLoD_c As Double = MulHiD_c / Ra1_c
' Note that, with these values, the rndDbl() value can be exactly 0.5;
' that can be a problem for algorithms which use rndDbl() - 0.5, but do not
' want that quantity to be zero. Adding 1E-54 or less will fix the problem.

' Default seed values - used when generators are called before initialization
Private Const DefaultSeedA_c As Long = 543210
Private Const DefaultSeedB_c As Long = 1211109876

' Module-global variables (initialized as 0, "" or False when module starts)
Private temp_m As Long         ' combined sequence value
Private temp_m1 As Long        ' older combined sequence value
Private init_m As Boolean      ' flag to indicate if seeds have been initialized
Private normSaved_m As Boolean ' True if a rndNormal() value is being saved
Private rndA_m As Long         ' the present point in the sequence is held ...
Private rndB_m As Long         ' ... in rndA_m and rndB_m

' Variables used only if we are conducting unit tests
#If UnitTest_C Then
  Private test_m As Boolean      ' flag to indicate if this is only a test
  Private testVal_m As Double    ' value to be returned by rndDbl() if testing
  Private gammaCalls_m As Long   ' number of calls to rndDbl() in rndGamma()
  Private normalCalls_m As Long  ' number of calls to rndDbl() in rndNormal()
  Private poissonCalls_m As Long ' number of calls to rndDbl() in rndPoisson()
  Private ofi_m As Integer       ' output file index used by unit-test routine
#End If

'############################# Exported Routines ###############################

'===== rndDbl ==================================================================
Public Function rndDbl() _
As Double
Attribute rndDbl.VB_Description = "Sequence of pseudo-random Doubles that are uniformly distributed between 0 and 1, not including those values. The smallest value returned is 2^(-53) or 1.1102E-16, and the largest is 1 - 2^(-53) or 1 - 1.1102E-16."
' Returns a sequence of Double precision random variates uniformly distributed
' between zero and one, not including either. The smallest value returned is
' 2^(-53) or 1.11E-16, and the largest is 1 - 2^(-53) or 0.9999999999999998.
' There are about 9.0E15 distinct output values, but the sequence does not
' repeat until called about 4.6E18 times. The integer sequence generator is
' advanced twice in this routine.
' You ought to do "rndSeedSet seedA, seedB" before using the sequence values.
'
' Produces about 6 million variates per second when compiled and run on a
' 2 GHz Pentium. Interpreted, produces about 480,000 per second.
'
' Version of 17 May 2005 by John Trenholme
#If ExcelFunction_C Then
  Application.Volatile True  ' result changes even if input does not
#End If
Dim jTemp As Long
' check that the seeds have been set; use defaults if not
If Not init_m Then
  #If UnitTest_C Then
    ' if we are also in test mode, return special test value
    If test_m Then
      rndDbl = testVal_m
      Exit Function
    End If
  #End If
  rndSeedSet DefaultSeedA_c, DefaultSeedB_c
End If

' Do mixed congruential generator A to update integer A using the
' no-overflow signed-integer "approximate factoring" trick due to
' Wichmann & Hill, as elaborated by Schrage.
' See L. Schrage, ACM Trans. Math. Soft. 5 (1979), pp. 132-138 and Bratley,
' Fox & Schrage, "A Guide to Simulation" 2d edition, Springer-Verlag, 1987
jTemp = rndA_m \ Ra4_c
rndA_m = Ra2_c * (rndA_m - jTemp * Ra4_c) - jTemp * Ra5_c + Ra3_c
If rndA_m < 0& Then rndA_m = rndA_m + Ra1_c

' do mixed congruential generator B to update integer B (see ref. above)
jTemp = rndB_m& \ Rb4_c
rndB_m = Rb2_c * (rndB_m - jTemp * Rb4_c) - jTemp * Rb5_c + Rb3_c
If rndB_m < 0& Then rndB_m = rndB_m + Rb1_c

' combine the new integers and put into the valid range for integer A
temp_m1 = rndA_m - rndB_m
If temp_m1 < 0& Then temp_m1 = temp_m1 + Ra1_c

' do mixed congruential generator A to update integer A
jTemp = rndA_m \ Ra4_c
rndA_m = Ra2_c * (rndA_m - jTemp * Ra4_c) - jTemp * Ra5_c + Ra3_c
If rndA_m < 0& Then rndA_m = rndA_m + Ra1_c

' do mixed congruential generator B to update integer B
jTemp = rndB_m& \ Rb4_c
rndB_m = Rb2_c * (rndB_m - jTemp * Rb4_c) - jTemp * Rb5_c + Rb3_c
If rndB_m < 0& Then rndB_m = rndB_m + Rb1_c

' combine the new integers and put into the valid range for integer A
temp_m = rndA_m - rndB_m
If temp_m < 0& Then temp_m = temp_m + Ra1_c

' Scale into a Double value obeying 1.11E-16 <= rndDbl() < 1.0 - 1.11E-16.
' The term in temp_m1 adds (almost) 31 bits in the top end of the Double.
' The term in temp_m adds another 22 bits onto the bottom end of the Double
' and is adjusted in size so that the part it adds on just fits between
' the values produced by the term in temp_m1, without overlap or gap.
rndDbl = AddD_c + temp_m1 * MulHiD_c + temp_m * MulLoD_c
End Function

'===== rndExponential ==========================================================
Public Function rndExponential( _
  Optional ByVal meanValue As Double = 1#) _
As Double
Attribute rndExponential.VB_Description = "Sequence of pseudo-random Doubles that are exponentially distributed with supplied mean. When mean = 1, values obey 1.1102E-16 < rndExponential < 36.7368"
' Returns a sequence of pseudo-random variates from an exponential distribution
' of specified mean. If no mean is supplied, variates with unit mean are
' produced. With unit mean, the smallest value returned is 1.1102E-16, and the
' largest is 36.7368.
' Although "meanValue" ought to be strictly positive, it causes no harm to use
' zero or negative values - you just get back a zero or negative variate.
' You ought to do "rndSeedSet seedA, seedB" before using the sequence values.
'
' Uses rndDbl() as a source of pseudo-random uniform variates, calling it once.
'
' Version of 29 Jun 2000 by John Trenholme
#If ExcelFunction_C Then
  Application.Volatile True  ' result changes even if input does not
#End If
rndExponential = -meanValue * Log(rndDbl())
End Function

'===== rndGamma ================================================================
Public Function rndGamma( _
  Optional ByVal meanValue As Double = 1#, _
  Optional ByVal stdDev As Double = 1#) _
As Double
Attribute rndGamma.VB_Description = "Sequence of pseudo-random Doubles that are Gamma distributed with supplied mean and standard deviation."
' Returns a sequence of pseudo-random variates from a gamma distribution of
' specified mean and standard deviation. If no mean is supplied, a variate with
' unit mean is produced. If no standard deviation is supplied, a variate with
' unit standard deviation is produced.
'
' The gamma(a, b) distribution has the PDF given by (here a > 0 and b > 0):
'
'   p(x) = x ^ (a - 1) * Exp(-x / b) / (Gamma(a) * b ^ a)
'
' The mean value is m = a * b, the standard deviation is s = b * Sqr(a), and the
' mode is md = (a - 1) * b if a > 1, or md = 0 if a < 1. Thus a = (m / s) ^ 2,
' b = s ^ 2 / m, and md = m - s ^ 2 / m if m > s, or 0 if m < s.
' We can therefore write the distribution as:
'
'   p(x) = x ^ ((m / s) ^ 2 - 1) * Exp(-x * m / s ^ 2) / (Gamma((m / s) ^ 2) *
'          (s ^ 2 / m) ^ ((m / s) ^ 2))
'
' We generate a distribution with b = 1, and then scale the result by b. The
' method used is given in "A Simple Method for Generating Gamma Variables" by
' G. Marsaglia & W. Tsang, ACM Transactions on Mathematical Software, Vol. 26
' #3, September 2000, pp. 363-372.
'
' Uses rndDbl() as a source of pseudo-random uniform variates, calling it about
' 2.3 times when a >= 1 (mean >= std. dev.), and 3.4 times when a < 1.
'
' Version of 2005-08-09 by John Trenholme
Const ID_C As String = File_c & "rndGamma"
#If ExcelFunction_C Then
  Application.Volatile True  ' result changes even if input does not
#End If
Const MuMin As Double = 3.8E-306
Const OneThird As Double = 1# / 3#
#If UnitTest_C Then
  gammaCalls_m = 0&
#End If
If meanValue <= MuMin Then
  ' note: when called from Excel, Err.Raise causes Excel's #VALUE! error
  Err.Raise InvalidArg_c, ID_C, _
  "meanValue-too-small problem:" & vbLf & _
  "meanValue must be >= " & MuMin & " but got " & meanValue & vbLf & _
  "Fix: change your code to avoid or trap small or negative values" & vbLf & _
  "Problem in " & ID_C
End If
Dim a As Double, b As Double  ' the "standard" notation
Dim s2 As Double
s2 = stdDev * stdDev
b = s2 / meanValue
' requests for a really small standard deviation are made "reasonably" small
If b < meanValue * 1E-250 Then b = meanValue * 1E-250
a = meanValue / b
Dim lessThanOne As Boolean
If a >= 1# Then
  lessThanOne = False
Else
  a = a + 1#
  lessThanOne = True
End If
Dim c As Double, d As Double
d = a - OneThird
c = OneThird / Sqr(d)
Dim v As Double, x As Double
Do
  Do
    x = rndNormal(0#, 1#)
    #If UnitTest_C Then
      gammaCalls_m = gammaCalls_m + normalCalls_m
    #End If
    v = 1# + c * x
  Loop While v <= 0#  ' repeat is rare; only when Normal < -Sqr(9*a-3)
  v = v * v * v
  #If UnitTest_C Then
    gammaCalls_m = gammaCalls_m + 1&
  #End If
  If Log(rndDbl()) < 0.5 * x * x + d * (1# - v + Log(v)) Then Exit Do
Loop
If lessThanOne Then
  d = d * rndDbl() ^ (1# / a)
  #If UnitTest_C Then
    gammaCalls_m = gammaCalls_m + 1&
  #End If
End If
rndGamma = d * v * b
End Function

'===== gammaLog =============================================================
Public Function gammaLog( _
  ByVal x As Double) _
As Double
Attribute gammaLog.VB_Description = "Natural logarithm of the absolute value of the Gamma function g(x) = (x-1)! at the supplied argument. To get Gamma(x), use gammaSign(x) * Exp(gammaLog(x)), but note that this will overflow if x > 171.6"
' Absolute value of the logarithm of the Gamma function g(x) = (x-1)!.
'
' To get Gamma(x), use gammaSign(x) * Exp(gammaLog(x)), but note that this
' will overflow if x > 171.6
'
' Based on a Lanczos-type-series routine by Allen Miller of CSIRO. Has one less
' term in the series inside the log, and has the resulting "5.5" tweaked.
'
' The return value is consistently within 1E-14 of the true function (using
' absolute error when the function is < 1, relative error when the function is
' > 1). Much of the time, this error measure is within 1E-15 of the true value.
'
' Error limits: -2.93E18 < x < -5.57E-309 and 5.57E-309 < x < 2.55E305
'
' Version of 20 July 2007 by John Trenholme
Const Split As Double = 0.5
Dim u As Double
If x >= Split Then u = x Else u = 1# - x
Dim v As Double
v = (u - 0.5) * Log(u + 5.877) - u + Log(0.007026535811543 + _
  2.42425246210031 / u - 3.96218410578371 / (u + 1#) + _
  2.00272142074634 / (u + 2#) - 0.335162034406637 / (u + 3#) + _
  0.012977763638767 / (u + 4#) - 0.000027272752074 / (u + 5#))
If x >= Split Then
  gammaLog = v
Else
  gammaLog = Log(Abs(Pi_c / Sin(Pi_c * x))) - v - 6.5E-15
End If
End Function

'===== gammaSign ============================================================
Public Function gammaSign( _
  ByVal x As Double) _
As Double
Attribute gammaSign.VB_Description = "Sign of the Gamma function at the supplied argument. Will be negative for some negative arguments. Zero at poles 0, -1, -2, -3..."
' Returns the sign of the Gamma function, which is negative when the argument
' x is negative, and x also lies between an odd integer and the integer
' that is one larger. Note that Gamma has poles at x = 0, -1, -2, and so on.
' This function returns 0 at those points.
'
' To get Gamma(x), use gammaSign(x) * Exp(gammaLog(x)), but note that this
' will overflow if x > 171.6
'
' Error limits: -4.50E15 < x < 1.79E308
'
' Version of 20 July 2007 by John Trenholme
If x <= 0# Then
  x = 0.5 * x - 0.25
  x = Abs(x - Int(x) - 0.5) - 0.25
End If
gammaSign = Sgn(x)
End Function

'===== rndGumbel ===============================================================
Public Function rndGumbel( _
  Optional ByVal meanValue As Double = 0#, _
  Optional ByVal stdDev As Double = 1#) _
As Double
Attribute rndGumbel.VB_Description = "Sequence of pseudo-random Doubles that are Gumbel extreme-value distributed with supplied mean & standard deviation. With arguments (1,1) values obey -3.2599 < rndGumbel < 28.1935"
' Returns a sequence of pseudo-random variates from a Gumbel (AKA extreme-value,
' double-exponential, Fisher-Tippett type 1) distribution of specified mean and
' standard deviation. If no mean is supplied, a variate with zero mean is
' produced. If no standard deviation is supplied, a variate with unit standard
' deviation is produced. With zero mean and unit standard deviation, the
' smallest value returned is -3.2599, and the largest is 28.1935. If a negative
' standard deviation is specified, the resulting variates will have the long
' exponential tail extending to smaller values, rather than larger values,
' contrary to the usual definition (this is appropriate for extreme minima).
' You ought to do "rndSeedSet seedA, seedB" before using the sequence values.
'
' Uses rndDbl() as a source of pseudo-random uniform variates, calling it once.
'
' Version of 2005-08-10 by John Trenholme
#If ExcelFunction_C Then
  Application.Volatile True  ' result changes even if input does not
#End If
Const Root6overPi As Double = 0.7796968 + 1.2336761E-09 ' Sqr(6) / Pi
Const EulerMascheroni As Double = 0.57721566 + 4.90153286E-09
rndGumbel = meanValue - stdDev * Root6overPi * (Log(-Log(rndDbl())) + _
  EulerMascheroni)
End Function

'===== rndLng ==================================================================
Public Function rndLng() _
As Long
Attribute rndLng.VB_Description = "Sequence of pseudo-random Long integer values uniformly distributed from 0 to rndLngMax(), inclusive."
' Returns a sequence of pseudo-random long integers from a uniform distribution
' between 0 and rndLngMax(), inclusive. Note that rndLngMax() = 2147483562.
' The integer sequence generator is advanced once in this routine.
' You ought to do "rndSeedSet seedA, seedB" before using the sequence values.
'
' Version of 17 May 2005 by John Trenholme
#If ExcelFunction_C Then
  Application.Volatile True  ' result changes even if input does not
#End If
Dim jTemp As Long
' check that the seeds have been set; use defaults if not
If Not init_m Then rndSeedSet DefaultSeedA_c, DefaultSeedB_c

' do mixed congruential generator A to update integer A
' see comments in body of "rndDbl()" for an explanation of what's happening
jTemp = rndA_m \ Ra4_c
rndA_m = Ra2_c * (rndA_m - jTemp * Ra4_c) - jTemp * Ra5_c + Ra3_c
If rndA_m < 0& Then rndA_m = rndA_m + Ra1_c

' do mixed congruential generator B to update integer B
jTemp = rndB_m& \ Rb4_c
rndB_m = Rb2_c * (rndB_m - jTemp * Rb4_c) - jTemp * Rb5_c + Rb3_c
If rndB_m < 0& Then rndB_m = rndB_m + Rb1_c

' combine the new integers and put into the valid range for integer A
temp_m = rndA_m - rndB_m
If temp_m < 0& Then temp_m = temp_m + Ra1_c

' return the result
rndLng = temp_m
End Function

'===== rndLngMax ===============================================================
Public Function rndLngMax() _
As Double
Attribute rndLngMax.VB_Description = "The largest value returned by rndLng()."
' Returns the largest value that will be returned by rndLng().
' Note: the smallest value is 0.
'
' Version of 26 Jun 2003 by John Trenholme
rndLngMax = Ra1_c - 1&
End Function

'===== rndLogNormal ============================================================
Public Function rndLogNormal( _
  Optional ByVal meanValue As Double = 1#, _
  Optional ByVal stdDev As Double = 1#) _
As Double
Attribute rndLogNormal.VB_Description = "Sequence of pseudo-random Doubles that are log-normally distributed with supplied mean (must be positive) & standard deviation. With arguments (1,1) values obey 5.8189E-4 < rndLogNormal < 802.797"
' Returns a sequence of pseudo-random variates from a log-normal distribution of
' specified mean and standard deviation. If no mean is supplied, a variate with
' unit mean is produced. If no standard deviation is supplied, a variate with
' unit standard deviation is produced. With unit mean and unit standard
' deviation, the minimum value produced is 5.8189E-4 and the maximum is
' 802.797.
' The input argument "meanValue" must be strictly positive, so if it is not an
' error is raised.
' You ought to do "rndSeedSet seedA, seedB" before using the sequence values.
'
' Note that many authors define this distribution in terms of the mean and
' standard deviation of the Gaussian in the exponent, contrary to our usage.
' The values you supply will become the actual mean and standard deviation of
' the resulting distribution.
'
' See remarks in rndNormal regarding saved seeds and repeated sequences.
'
' Uses rndNormal() as source of pseudo-random normal variates, calling it once.
'
' Version of 11 Sep 2001 by John Trenholme (not a good day)
Const ID_C As String = File_c & "rndLogNormal"
#If ExcelFunction_C Then
  Application.Volatile True  ' result changes even if input does not
#End If
Const MuMin As Double = 1E-150
Dim t As Double
' try to make it unlikely that there will be overflow below
If meanValue < MuMin Then
  ' note: when called from Excel, Err.Raise causes Excel's #VALUE! error
  Err.Raise InvalidArg_c, ID_C, _
  "meanValue-too-small problem:" & vbLf & _
  "meanValue must be >= " & MuMin & " but got " & meanValue & vbLf & _
  "Fix: change your code to avoid or trap small or negative values" & vbLf & _
  "Problem in " & ID_C
End If
t = 1# + (stdDev / meanValue) ^ 2
rndLogNormal = Exp(rndNormal(Log(meanValue / Sqr(t)), Sqr(Log(t))))
End Function

'===== rndNormal ===============================================================
Public Function rndNormal( _
  Optional ByVal meanValue As Double = 0#, _
  Optional ByVal stdDev As Double = 1#) _
As Double
Attribute rndNormal.VB_Description = "Sequence of pseudo-random Doubles that are normally distributed with supplied mean & standard deviation. With arguments (0,1) values obey -8.5311 < rndNormal < 8.4495"
' Returns a sequence of pseudo-random variates from a normal (or Gaussian)
' distribution of specified mean and standard deviation. If no mean is supplied,
' a variate with zero mean is produced. If no standard deviation is supplied, a
' variate with unit standard deviation is produced. With zero mean and unit
' standard deviation, the minimum value produced is -8.5311 and the maximum is
' 8.4495, but such extreme values are extremely unlikely (the probability is a
' bit less than 1 per billion of a value above 6, and the same for a value below
' -6).
'
' You ought to do "rndSeedSet seedA, seedB" before using the sequence values.
'
' To produce two variates "v1" and "v2" with means "m1" and "m2," standard
' deviations "s1" and "s2," and correlation coefficient "c," do this (Knuth,
' "Seminumerical Algorithms" section 3.4.1(5)):
'
' v1 = rndNormal()
' v2 = m2 + s2 * (c * v1 + sqrt(1# - c * c) * rndNormal())
' v1 = m1 + s1 * v1
'
' Of course, you have to have -1.0 <= c <= 1.0. For more than two correlated
' variates, see Knuth, 3.4.1 Exercise 13.
'
' Uses rndDbl() as a source of pseudo-random uniform variates, calling it
' 4 / Pi = 1.2732 times per call on the average (2.5465 on the first call, then
' 0 on the next, then 2.5465, then 0, ...).
'
' This routine uses the Box-Muller-Marsaglia transformation. See Knuth,
' "Seminumerical Algorithms" section 3.4.1, Algorithm P, or the defining article
' in Annals Math. Stat. 28 (1958), p. 610. It's not the fastest method, but
' it produces superb accuracy with simple code.
'
' Note: Because this routine caches alternate values for the next call, and any
' cached value is discarded when the seed values are set, you may get unexpected
' behaviour. If you save the seed values at some point where an odd number of
' calls have been made to this routine, and later set the seeds to the saved
' values, the sequence past the seed-save point will be different. To avoid
' this, immediately set the seed values to what you just got from rndSeedGet()
' (or rndSeedGetA() and rndSeedGetB()) when you do the save, before proceeding
' with your code. This will discard any cached value, and make the sequence
' after the save point equal to the one after the restore point. That is:
'   rndSeedGet saveA, saveB
'   rndSeedSet saveA, saveB  ' do this to purge any cached values
'     ... use a sequence of rndNormal() or rndLogNormal() values ...
'   rndSeedSet saveA, saveB
'     ... use the same sequence of rndNormal() or rndLogNormal() values ...
'
' Version of 29 Jun 2000 by John Trenholme
#If ExcelFunction_C Then
  Application.Volatile True  ' result changes even if input does not
#End If
Static saved As Double  ' holds saved second value, if any
Dim t As Double
Dim x As Double
#If UnitTest_C Then
  normalCalls_m = 0&
#End If
If Not normSaved_m Then
  ' make up two Gaussian variates of zero mean and unit standard deviation
  ' the loop will be executed 4 / Pi = 1.2732 times on average
  Do
    x = rndDbl() - 0.5
    saved = rndDbl() - 0.5
    t = x * x + saved * saved
    #If UnitTest_C Then
      normalCalls_m = normalCalls_m + 2&
    #End If
  Loop Until (t <= 0.25) And (t > 0#) ' point is in uniform disk when we exit
  t = Sqr(-2# * Log(4# * t) / t)  ' transform uniform disk to Gaussian
  saved = saved * t  ' make saved value a standard normal variate
  rndNormal = meanValue + x * t * stdDev  ' set desired return value
  normSaved_m = True
Else
  rndNormal = meanValue + saved * stdDev  ' convert saved value
  normSaved_m = False
End If
End Function

'===== rndPoisson ==============================================================
Public Function rndPoisson( _
  Optional ByVal meanValue As Double = 1#) _
As Long
Attribute rndPoisson.VB_Description = "Sequence of pseudo-random Long values that are Poisson distributed with the supplied mean."
' Returns a sequence of pseudo-random variates from a Poisson distribution of
' specified mean. If no mean is supplied, a variate with unit mean is produced.
' If "meanValue" is negative, an error is raised.
' You ought to do "rndSeedSet seedA, seedB" before using the sequence values.
'
' Based on the routine in Numerical Recipes.
' Uses rndDbl() as a source of pseudo-random uniform variates. See comments
' below for the number of times it is called.
'
' Version of 26 Jun 2003 by John Trenholme
Const ID_C As String = File_c & "rndPoisson"
#If ExcelFunction_C Then
  Application.Volatile True  ' result changes even if input does not
#End If
Dim ex As Double
Dim g As Double
Dim logMean As Double
Dim sq As Double
Dim t As Double
Dim y As Double
Dim k As Long
#If UnitTest_C Then
  poissonCalls_m = 0&
#End If
If meanValue <= 0# Then
  ' note: when called from Excel, Err.Raise causes Excel's #VALUE! error
  Err.Raise InvalidArg_c, ID_C, _
  "meanValue-negative problem:" & vbLf & _
  "meanValue must be >= 0 but got " & meanValue & vbLf & _
  "Fix: change your code to avoid or trap small or negative values" & vbLf & _
  "Problem in " & ID_C
End If
If meanValue <= 12# Then
  ' small mean; use sum of exponential variates
  ' will call rndDbl() a total of (meanValue + 1) times on average
  ex = Exp(-meanValue)
  t = rndDbl()
  #If UnitTest_C Then
    poissonCalls_m = poissonCalls_m + 1&
  #End If
  k = 0&
  Do While t > ex
    t = t * rndDbl()
    #If UnitTest_C Then
      poissonCalls_m = poissonCalls_m + 1&
    #End If
    k = k + 1&
  Loop
  rndPoisson = k
Else
  ' large mean; use the rejection method
  ' will call rndDbl() around 3.8 to 3.9 times on average
  sq = Sqr(2# * meanValue)
  logMean = Log(meanValue)
  g = meanValue * logMean - gammaLog(meanValue + 1#)
  Do  ' try until we are in the acceptance region
    Do  ' get a positive value from the Lorentzian distribution
      y = Tan(Pi_c * rndDbl())
      #If UnitTest_C Then
        poissonCalls_m = poissonCalls_m + 1&
      #End If
      t = meanValue + sq * y  ' set mean & variance to enclose Poisson
    Loop Until t >= 0#
    t = Int(t)  ' move down to the first integer below trial value
    #If UnitTest_C Then
      poissonCalls_m = poissonCalls_m + 1&
    #End If
  Loop Until rndDbl() < _
             0.9 * (1# + y * y) * Exp(t * logMean - gammaLog(t + 1#) - g)
  rndPoisson = t
End If
End Function

'===== rndRicianInten ==========================================================
Public Function rndRicianInten( _
  ByVal constField As Double, _
  ByVal addedStdDev As Double) _
As Double
Attribute rndRicianInten.VB_Description = "Sequence of pseudo-random Doubles that are Rician-intensity distributed based on supplied constant field (positive) & standard deviation (positive) of added random phasors."
' Returns a sequence of pseudo-random variates from a Rician intensity
' distribution (also called modified Rician distribution), given normalized
' field values of the random process producing the Rician.
'
' Input values are:
'   constField  = field value of constant phasor
'   addedStdDev = field standard deviation of added random phasors
'
' If mu is the mean intensity and sd is the standard deviation, then the input
' field quantities should be specified as:
'
'   constField = Sqr(Sqr(mu * mu - sd * sd))
'   addedStdDev = Sqr((mu - Sqr(mu * mu - sd * sd)) / 2#)
'
' If mu is the mean intensity and cn is the "contrast" = std.dev. / mean, then
' the input field quantities should be specified as:
'
'   constField = Sqr(mu * Sqr(1# - cn * cn))
'   addedStdDev = Sqr((1# - Sqr(1# - cn * cn)) * mu / 2#)
'
' In terms of the constant field f & the added std. dev. s, the mean intensity,
' standard deviation and contrast are given by:
'
'    mu = f * f + 2# * s * s
'    sd = 2# * s * Sqr(f * f + s * s)
'    cn = 2# * s * Sqr(f * f + s * s) / (f * f + 2# * s * s)
'
' You ought to do "rndSeedSet seedA, seedB" before using the sequence values.
'
' Uses rndDbl() as a source of pseudo-random uniform variates, calling it
' 8 / Pi = 2.5465 times on the average.
'
' Version of 27 Jun 2000 by John Trenholme
#If ExcelFunction_C Then
  Application.Volatile True  ' result changes even if input does not
#End If
Dim t As Double
Dim x As Double
Dim y As Double
' make up two normal variates of zero mean and given standard deviation to
' represent the in-phase and quadrature values of the added random phasor sum
' the loop will be executed 4 / Pi = 1.2732 times on average
Do
  x = rndDbl() - 0.5
  y = rndDbl() - 0.5
  t = x * x + y * y
Loop Until (t <= 0.25) And (t > 0#) ' point is in circle when we exit
t = Sqr(-2# * Log(4# * t) / t)  ' transform uniform disk to Gaussian
x = x * t * addedStdDev
y = y * t * addedStdDev
' total in-phase field is sum of constant field and added in-phase component
t = constField + x
' add squares of in-phase and quadrature fields to get intensity
rndRicianInten = t * t + y * y
End Function

'===== rndSeedGet ==============================================================
Public Sub rndSeedGet( _
ByRef seedA As Long, _
ByRef seedB As Long)
Attribute rndSeedGet.VB_Description = "Sets its ByRef arguments to the two Long seeds that specify the present point in the sequence. They may be changed from values set by rndSeedSet() due to modulo arithmetic."
' Return both pseudo-random number seeds, by changing the values supplied as
' arguments. SeedA will always be 0 <= seedA < Ra1_c, and seedB will always be
' 0 <= seedB < Rb1_c, no matter what you tried to set them to with rndSeedSet.
'
' For example:
'   Dim seedA As Long, seedB As Long
'   rndSeedGet seedA, seedB
'
' Version of 18 Jun 2005 by John Trenholme
seedA = rndA_m
seedB = rndB_m
End Sub

'===== rndSeedGetA =============================================================
Public Function rndSeedGetA() _
As Long
Attribute rndSeedGetA.VB_Description = "Returns the first of the two Long seeds that specify the present point in the sequence. May be changed from value set by rndSeedSet() due to modulo arithmetic."
' Return first pseudo-random number seed. Will always be 0 <= seedA < Ra1_c, no
' matter what you tried to set it to with rndSeedSet.
'
' Version of 18 Jun 2005 by John Trenholme
rndSeedGetA = rndA_m
End Function

'===== rndSeedGetB =============================================================
Public Function rndSeedGetB() _
As Long
Attribute rndSeedGetB.VB_Description = "Returns the second of the two Long seeds that specify the present point in the sequence. May be changed from value set by rndSeedSet() due to modulo arithmetic."
' Return second pseudo-random number seed. Will always be 0 <= seedB < Rb1_c, no
' matter what you tried to set it to with rndSeedSet.
'
' Version of 18 Jun 2005 by John Trenholme
rndSeedGetB = rndB_m
End Function

'===== rndSeedSet ==============================================================
Public Sub rndSeedSet( _
Optional ByVal seedA As Long = DefaultSeedA_c, _
Optional ByVal seedB As Long = DefaultSeedB_c)
Attribute rndSeedSet.VB_Description = "Sets both pseudo-random number seeds, setting the starting (or re-starting) point in all the sequences. The optional defaults are a known-good start point."
' Initialize random number generator from two seed integers, and purge any
' cached values in the pseudo-random routines. seedA is interpreted modulo
' Ra1_c = 2147483563, and seedB is interpreted modulo Rb1_c = 2147483399.
'
' For a known-good sequence, you can set seedA = 543210 and seedB = 1211109876:
'    rndSeedSet 543210, 1211109876
' Note: these values are supplied as defaults when you call rndSeedSet without
' any arguments.
'
' For independent sequences, set seedA = 543210 (etc.) with seedB = 0, 1, 2, ...
' Note that if you set only one of the two integers, you should preserve the
' other by calling rndSeedGetA or rndSeedGetB before the change, as follows:
'    rndSeedSet 543210, rndSeedGetB()
'    rndSeedSet rndSeedGetA(), 1211109876
'
' To do the same thing that "Randomize" with no argument does for the built-in
' generator "Rnd()" you need a way to turn the time "right now" into a seed
' value for this routine. That is accomplished by use of the routine
' "time2Long" below. Do this (substituting anything you want for the "0" if
' desired):
'    rndSeedSet time2Long(), 0&
'
' Note that it is not a good idea to initialize the two seeds to values that
' always have the same difference. For example, don't do this:
'    rndSeedSet 12& + j, 654& + j
' The reason is that the two sequences are combined by subtraction, and so the
' difference of the seeds should not be constant. You get the same result by
' changing just one seed:
'    rndSeedSet 12& + j, 654&
' If you really want to change both, change them in opposite directions:
'    rndSeedSet 12& + j, 654& - j
'
' Version of 17 May 2005 by John Trenholme
' modulus arithmetic - force first integer into range 0 <= rndA_m < Ra1_c
rndA_m = seedA
Do While rndA_m < 0&
  rndA_m = rndA_m + Ra1_c
Loop
Do While rndA_m >= Ra1_c
  rndA_m = rndA_m - Ra1_c
Loop
' modulus arithmetic - force second integer into range 0 <= rndB_m < Rb1_c
rndB_m = seedB
Do While rndB_m < 0&
  rndB_m = rndB_m + Rb1_c
Loop
Do While rndB_m >= Rb1_c
  rndB_m = rndB_m - Rb1_c
Loop
' set flag to show values are initialized
init_m = True
' throw out any cached values, so next call will reflect new seeds
normSaved_m = False
End Sub

'===== rndSng ==================================================================
Public Function rndSng() _
As Single
Attribute rndSng.VB_Description = "Sequence of pseudo-random Singles that are uniformly distributed between 0 and 1, not including those values. The smallest value returned is 2^(-24) or 5.96E-8, and the largest is 1 - 2^(-24) or 0.99999994."
' Returns a sequence of pseudo-random variates uniformly distributed between
' zero and one, not including either. The smallest value returned is 2^(-24)
' or 5.96E-8, and the largest is 1 - 2^(-24) or 0.99999994. There are about
' 1.7E7 distinct output values, but the sequence does not repeat until called
' about 4.6E18 times. When you need more than 1.7E7 distinct variates, this
' routine will not be suitable. To get finer-grained values, use "rndDbl()".
' The integer sequence generator is advanced once in this routine.
' You ought to do "rndSeedSet seedA, seedB" before using the sequence values.
'
' Produces about 11 million variates per second when compiled and run on a
' 2 GHz Pentium. Interpreted, produces about 860,000 per second.
'
' Version of 2005-08-04 by John Trenholme
#If ExcelFunction_C Then
  Application.Volatile True  ' result changes even if input does not
#End If
Dim jTemp As Long
' check that the seeds have been set; use defaults if not
If Not init_m Then rndSeedSet DefaultSeedA_c, DefaultSeedB_c

' do mixed congruential generator #1 to update integer #1
' see comments in body of "rndDbl()" for an explanation of what's happening
jTemp = rndA_m \ Ra4_c
rndA_m = Ra2_c * (rndA_m - jTemp * Ra4_c) - jTemp * Ra5_c + Ra3_c
If rndA_m < 0& Then rndA_m = rndA_m + Ra1_c

' do mixed congruential generator #2 to update integer #2
jTemp = rndB_m& \ Rb4_c
rndB_m = Rb2_c * (rndB_m - jTemp * Rb4_c) - jTemp * Rb5_c + Rb3_c
If rndB_m < 0& Then rndB_m = rndB_m + Rb1_c

' combine the new integers and put into the valid range for integer A
temp_m = rndA_m - rndB_m
If temp_m < 0& Then temp_m = temp_m + Ra1_c

' scale into a Single value obeying 5.96E-8! <= rndSng() <= 1! - 5.96E-8!
rndSng = AddS_c + temp_m * MulS_c
End Function

'===== time2Long ============================================================
Public Function time2Long( _
  Optional ByVal time As Variant = 0&) _
As Long
Attribute time2Long.VB_Description = "Convert supplied time to a Long holding the number of 67ths of a second between the supplied time and the start of the year included in that time. Used with rndSeedSet() to replace Randomize. Note 0 <= time2Long <= 2112912000 (not leap) or 2118700800 (leap year)"
' Convert a time to a long that holds the number of 67ths of a second between
' the supplied time and the start of the year specified in the supplied time.
' The minimum value returned is 0, at midnight on Jan 1 (#1/1/YYYY#).
' The maximum value returned is 2,112,912,000 (or 2,118,700,800 during leap
' years) if Now() is supplied; it is 67 less if a Date is supplied. This
' routine is supplied in this Module so that the Visual Basic routine
' "Randomize" (or "Randomize Timer") can be replaced by, and improved upon,
' by (you can replace the 0& here by any other value):
'    rndSeedSet time2Long(), 0&
' or
'    rndSeedSet 0&, time2Long()
'
' If no argument is supplied, "Now()" is used as the time. On a standard PC,
' the result will change by 2 or 3 units as time proceeds because of the
' granularity of the "Timer" function. If an argument is supplied, the result
' changes only every second, as fractional seconds are not available in a Date.
'
' The apparently arbitrary choice of 67ths of a second was made so that the
' result would be close to 2^31 (the largest possible Long) at the end of
' a leap year, thus supplying almost all possible positive seed values.
'
' Version of 6 Aug 2002 by John Trenholme
Const Rate As Double = 67#
' we use the type of the default to distinguish it from a Date
If IsDate(time) Then
  ' argument was supplied; it has no fractional seconds (Dates don't)
  time2Long = Rate * (86400# * (DatePart("y", time) - 1&) + _
                          3600# * DatePart("h", time) + _
                            60# * DatePart("n", time) + _
                                  DatePart("s", time))
Else  ' no argument was supplied - use Now() with year part removed + seconds
  time2Long = Rate * (86400# * (DatePart("y", Now()) - 1&) + Timer())
End If
End Function

'===== randomVersion ========================================================
Public Function RandomVersion(Optional ByVal trigger As Variant) _
As String
Attribute RandomVersion.VB_Description = "Date of the latest revision to this file, in the format 'yyyy-mm-dd'"
' Returns the date of the latest revision to this Module as a string in the
' format "yyyy-mm-dd".
RandomVersion = Version_c
End Function

'<*><*><*><*><*><*><*><*><*><*><*><*><*><*><*><*><*><*><*><*><*><*><*><*><*><*>

#If UnitTest_C Then  ' set UnitTest_C = True to use unit-test routines

'&&&&& RandomUnitTest &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Public Sub randomUnitTest()
Attribute randomUnitTest.VB_Description = "Perform tests on the routines in this module; output to Immediate window (Ctrl-G) & file ""Random[date].UnitTest.txt""."
' This checks for proper implementation and does some simple sanity checks on
' the output of the routines. Full-up testing of pseudo-random variates is a
' lengthy and difficult process, and we rely on the authors of the references
' above to assure the quality of the algorithms.
'
' Output goes to a file, & to Immediate window if in VB6 or VBA Editor.
'
' To run this routine from VBA, put the cursor somewhere in it and hit F5.
' To run this routine from VB6, enter "randomUnitTest" in the Immediate window.
' (If the Immediate window is not open, use View... or Ctrl-G to open it.)
'
' Version of 2012-11-12 by John Trenholme
Const Module_c As String = "Random.bas"
' get path to folder where Workbook resides
Dim path As String
' VBA file opening stuff - we presume we are running under Excel
' note: in Excel, must save a new workbook at least once before path exists
path = Excel.ActiveWorkbook.path
' if there is no workbook path, use CurDir (user's MyDocuments or Documents)
If vbNullString = path Then path = FileSystem.CurDir$
If vbNullString = path Then  ' no current directory?
  VBA.Interaction.Beep
  MsgBox _
    "Unable to find current folder" & vbLf & _
    "Excel workbook has no disk location, & no CurDir" & vbLf & _
    "Save workbook to disk before proceeding because" & vbLf & _
    "we need a known location to write the file to." & vbLf & _
    "Output will be sent to Immediate window only (Ctrl-G in editor)", _
    vbOKOnly Or vbCritical Or vbMsgBoxSetForeground, _
    File_c & " ERROR - No File Path"
  ofi_m = 0  ' no output to file
Else
  ' be sure path separator is at end of path (only C:\ etc. have it already)
  Dim ps As String
  ps = Application.PathSeparator
  If Right$(path, 1&) <> ps Then path = path & ps
  ' make up full file name, with path
  Dim fileName As String
  fileName = File_c & "UnitTest.txt"
  Dim ffn As String
  ffn = path & fileName
  ofi_m = FreeFile  ' file index, module-global so teeOut can use it
  ' try to open output file, over-writing any existing file
  On Error Resume Next
  Open ffn For Output Access Write Lock Write As #ofi_m
  Dim errNum As Long
  errNum = Err.Number
  On Error GoTo 0     ' clear Err object & enable default error handling
  If 0& <> errNum Then  ' cannot open file; disk full?
    Dim errDesc As String
    errDesc = Err.Description
    VBA.Interaction.Beep
    MsgBox _
      "Unable to open output file:" & vbLf & _
      """" & fileName & """" & vbLf & _
      "in folder:" & vbLf & _
      """" & left$(path, Len(path) - 1&) & """" & vbLf & _
      "Error: " & errDesc & vbLf & _
      "Error number: " & errNum & vbLf & _
     "Output will be sent to Immediate window only (Ctrl-G in editor)", _
      vbOKOnly Or vbCritical Or vbMsgBoxSetForeground, _
      File_c & " ERROR - Can't Open File"
    ofi_m = 0  ' no output to file
  End If
End If

teeOut "########## Test of " & Module_c & " routines at " & Now()
teeOut "Using Module " & left$(File_c, Len(File_c) - 1&)

Dim N As Long, n0 As Long
N = 0&  ' error count

teeOut vbNewLine & "---------- Tests of underlying generators"
rndSeedSet 229587626, 136699883  ' two steps back from 0, 0
Dim j As Long
Dim s As Single
For j = 1& To 6&
  s = rndSng()
Next j
If rndA_m <> 2032875953 Then _
  N = N + 1&: teeOut "OOPS! generator A not working (1): " & rndA_m
If rndB_m <> 63827039 Then _
  N = N + 1&: teeOut "OOPS! generator B not working (1): " & rndB_m

rndSeedSet 1478251139, 1468629129  ' go to 0, 0 when advanced
s = rndSng()
If rndA_m <> 0& Then _
  N = N + 1&: teeOut "OOPS! generator A not working (2): " & rndA_m
If rndB_m <> 0& Then _
  N = N + 1&: teeOut "OOPS! generator B not working (2): " & rndB_m
If N = 0& Then teeOut "No errors"

teeOut vbNewLine & "---------- Tests of rndSng & rndDbl implementation"
n0 = N

' Constants used to turn Long(s) into floating point values
teeOut "Evaluation constant values:"
teeOut "AddS_c = " & AddS_c & "  MulS_c = " & MulS_c
teeOut "AddD_c = " & AddD_c & " = 1.1102230246E-16 + " & _
  CSng(AddD_c - 1.1102230246E-16)
teeOut "MulHiD_c = " & MulHiD_c & " = 4.6566130573E-10 + " & _
  CSng(MulHiD_c - 4.6566130573E-10)
teeOut "MulLoD_c = " & MulLoD_c & " = 2.1684045166E-19 + " & _
  CSng(MulLoD_c - 2.1684045166E-19)

teeOut "Extreme values:"
' check extreme values of rndSng()
rndSeedSet 1478251139, 1468629129  ' go to 0, 0 when advanced
s = rndSng()
teeOut "smallest rndSng() = " & s & " = 2^(" & CSng(Log(s) / Log(2#)) & ")"
If s <= 0! Then _
  N = N + 1&: teeOut "OOPS! smallest rndSng not strictly positive: " & s
s = 1! - s  ' evaluate rndSng min subtracted from 1, using storage values
If s >= 1! Then _
  N = N + 1&: teeOut "OOPS! Minimum value of rndSng() is too small"
rndSeedSet 1543672803, 1468629129   ' go to 2147483562, 0 when advanced
s = rndSng()
If s < 1! Then
  teeOut "largest rndSng() = 1 - " & 1! - s & _
    " = 1 - 2^(" & CSng(Log(1# - s) / Log(2#)) & ")"
  s = 1! - s
  If s <= 0! Then
    N = N + 1&
    teeOut "OOPS! largest rndSng() too large: " & s & " <= 0"
  Else
    If -Log(s) / Log(2#) < 24 Then
      N = N + 1&
      teeOut "OOPS! rndSng() not using bit 24: " & -Log(s) / Log(2#)
    End If
  End If
Else
  N = N + 1&: teeOut "OOPS! largest rndSng not < 1.0: " & s
End If

' verify that 0.5 (exactly) is a possible result (from (Ra1_c-1)/2)
rndSeedSet 1510961971, 1468629129   ' go to 1073741781, 0 when advanced
s = rndSng()
teeOut "At k = " & rndSeedGetA() & ", rndSng() = " & s & _
  " (rndSng() - 0.5! = " & s - 0.5! & ")"

' check extreme values of rndDbl()
' WARNING! not using actual code!
Dim dLo As Double
dLo = rd(0&, 0&)
If dLo <> 2# ^ (-53) Then _
  N = N + 1&: teeOut "OOPS! smallest rndDbl not = 2^(-53): " & dLo
teeOut "smallest rndDbl() = " & dLo & _
  " = 2^(" & CSng(Log(dLo) / Log(2#)) & ")"
Dim dHi As Double
dHi = rd(2147483562, 2147483562)
If dHi < 1# Then
  teeOut "largest rndDbl() = 1 - " & 1# - dHi & _
    " = 1 - 2^(" & CSng(Log(1# - dHi) / Log(2#)) & ")"
  If dHi <> 1# - 2# ^ (-53) Then _
    N = N + 1&: teeOut "OOPS! largest rndDbl not = 1 - 2^(-53): " & dHi & _
    " = 1 - 2^(" & CSng(Log(1# - dHi) / Log(2#)) & ")"
Else
  N = N + 1&: teeOut "OOPS! largest rndDbl() not < 1.0: " & dHi
End If

' verify that 0.5 (exactly) is a possible result (from both = (Ra1_c-1)/2)
' WARNING! not using actual code!
Dim k As Long
k = 1073741781
Dim d As Double
d = rd(k, k)
teeOut "At k1 = k2 = " & k & ", rndDbl() = " & d & _
  " (rndDbl() - 0.5 = " & d - 0.5 & ")"

' check that large and small parts of rndDbl() interleave properly
' WARNING! not using actual code!
Dim k1 As Long, k2 As Long, k3 As Long
' start back about 5.5 Double increments (about 511 per Double)
k1 = Ra1_c - 2&
k2 = Ra1_c - 5.5 * 511&
Dim str As String
str = "Number of rndDbl's per Double:"
d = rd(k1, k2)
For j = 0& To 12
  k3 = -1&
  Do
    k3 = k3 + 1&
    k2 = k2 + 1&
    If k2 >= Ra1_c Then
      k2 = 0&
      k1 = k1 + 1&
      str = str + " |"
    End If
  Loop While d = rd(k1, k2)
  d = rd(k1, k2)
  If j > 0 Then
    str = str & " " & k3
    ' counts may be 510, 511 or 512 without complaint
    If Abs(k3 - 511&) > 1& Then _
      N = N + 1&: teeOut "OOPS! rndDbl() not interleaving properly: " & k3
  End If
Next j
teeOut str
If N = n0 Then teeOut "No errors"

teeOut vbNewLine & "---------- Simple sanity checks on derived variates"
' check that variates with adjustable standard deviation don't vary when 0
n0 = N
d = rndGumbel(1#, 0#)
If d <> 1# Then _
  N = N + 1&: teeOut "OOPS! rndGumbel(1.0,0.0) not = 1.0 but = " & d
d = rndNormal(1#, 0#)
If d <> 1# Then _
  N = N + 1&: teeOut "OOPS! rndNormal(1.0,0.0) not = 1.0 but = " & d
d = rndLogNormal(1#, 0#)
If d <> 1# Then _
  N = N + 1&: teeOut "OOPS! rndLogNormal(1.0,0.0) not = 1.0 but = " & d
d = rndRicianInten(1#, 0#)
If d <> 1# Then _
  N = N + 1&: teeOut "OOPS! rndRicianInten(1.0,0.0) not = 1.0 but = " & d
d = rndGamma(1#, 0#)
If d <> 1# Then _
  N = N + 1&: teeOut "OOPS! rndGamma(1.0,0.0) not = 1.0 but = " & d
If N = n0 Then teeOut "No errors"

teeOut vbNewLine & "---------- Derived-variate call counts to rndDbl()"
Dim M As Long
M = 0
normSaved_m = False ' purge any saved value
For j = 1& To 10000
  If rndNormal() > -1000# Then M = M + normalCalls_m
Next j
teeOut "rndNormal:  calls = " & 0.0001 * M & " (expect 1.2732 " & _
  Chr$(177) & " 0.013)"
M = 0
For j = 1& To 10000
  If rndPoisson(1#) > 0# Then M = M + poissonCalls_m
Next j
teeOut "rndPoisson: mean = 1  calls = " & 0.0001 * M & " (expect 2)"
M = 0
For j = 1& To 10000
  If rndPoisson(11.9999) > 0# Then M = M + poissonCalls_m
Next j
teeOut "rndPoisson: mean = 11.9999  calls = " & 0.0001 * M & " (expect 13)"
M = 0
For j = 1& To 10000
  If rndPoisson(12# + 12# * rndDbl()) > 0# Then M = M + poissonCalls_m
Next j
teeOut "rndPoisson: mean from 12.0 to 24.0  calls = " & 0.0001 * M
M = 0
For j = 1& To 10000
  If rndPoisson(25# + 75# * rndDbl()) > 0# Then M = M + poissonCalls_m
Next j
teeOut "rndPoisson: mean from 25.0 to 100.0  calls = " & 0.0001 * M
M = 0
For j = 1& To 10000
  If rndPoisson(100# + 900# * rndDbl()) > 0# Then M = M + poissonCalls_m
Next j
teeOut "rndPoisson: mean from 100.0 to 1000.0  calls = " & 0.0001 * M
M = 0
For j = 1& To 10000
  If rndGamma(1#, 1# + 9# * rndDbl()) > 0# Then M = M + gammaCalls_m
Next j
teeOut "rndGamma: mean 1, s.d. from 1 to 10  calls = " & 0.0001 * M
M = 0
For j = 1& To 10000
  If rndGamma(2#, 2# * rndDbl()) > 0# Then M = M + gammaCalls_m
Next j
teeOut "rndGamma: mean 2, s.d. from 0 to 2  calls = " & 0.0001 * M

teeOut vbNewLine & "---------- Smallest and largest derived-variate values"
test_m = True
init_m = False
testVal_m = 0.5 - 2 ^ 8 * MulLoD_c
normSaved_m = False ' purge any saved value
teeOut "minimum rndNormal(0,1): " & rndNormal()
testVal_m = 0.5 + 2 ^ 9 * MulLoD_c
normSaved_m = False ' purge any saved value
teeOut "maximum rndNormal(0,1): " & rndNormal()
testVal_m = 0.5 - 2 ^ 8 * MulLoD_c
normSaved_m = False ' purge any saved value
teeOut "minimum rndLogNormal(1,1): " & rndLogNormal()
testVal_m = 0.5 + 2 ^ 9 * MulLoD_c
normSaved_m = False ' purge any saved value
teeOut "maximum rndLogNormal(1,1): " & rndLogNormal()
testVal_m = dHi
teeOut "minimum rndExponential(1): " & rndExponential()
testVal_m = dLo
teeOut "maximum rndExponential(1): " & rndExponential()
testVal_m = dLo
teeOut "minimum rndGumbel(0,1): " & rndGumbel()
testVal_m = dHi
teeOut "maximum rndGumbel(0,1): " & rndGumbel()

teeOut vbNewLine & "---------- Tests using forced values from rndDbl()"
n0 = N
testVal_m = 0.5
M = rndPoisson(1#)
If M <> 1& Then _
  N = N + 1&: teeOut "OOPS! Poisson variate failed test: wanted 1, got " & M
If N = n0 Then teeOut "No errors"

test_m = False
init_m = False  ' use default starting seeds

M = 10000
teeOut vbNewLine & "---------- mean-value tests with " _
  & Format(M, "0,0") & " samples"
rndSeedSet time2Long(), 0&   ' randomize the sequence
teeOut "Starting seeds: " & rndSeedGetA() & " " & rndSeedGetB() & _
  " (randomized with time2Long)"
Dim dMax As Double
dMax = 0#

d = 0#
For j = 1& To M
  d = d + rndDbl()
Next j
d = d / (0.5 * M) - 1#
teeOut "rndDbl: actual / expected - 1 = " & CSng(d)
If Abs(dMax) < Abs(d) Then dMax = d

d = 0#
For j = 1& To M
  d = d + rndExponential(3#)
Next j
d = d / (3# * M) - 1#
teeOut "rndExponential: actual / expected - 1 = " & CSng(d)
If Abs(dMax) < Abs(d) Then dMax = d

d = 0#
For j = 1& To M
  d = d + rndGamma(3.3, 0.6)
Next j
d = d / (3.3 * M) - 1#
teeOut "rndGamma: actual / expected - 1 = " & CSng(d)
If Abs(dMax) < Abs(d) Then dMax = d

d = 0#
For j = 1& To M
  d = d + rndGumbel(2.5, 1#)
Next j
d = d / (2.5 * M) - 1#
teeOut "rndGumbel: actual / expected - 1 = " & CSng(d)
If Abs(dMax) < Abs(d) Then dMax = d

d = 0#
For j = 1& To M
  d = d + rndLogNormal(1.5, 0.7)
Next j
d = d / (1.5 * M) - 1#
teeOut "rndLogNormal: actual / expected - 1 = " & CSng(d)
If Abs(dMax) < Abs(d) Then dMax = d

d = 0#
For j = 1& To M
  d = d + rndNormal(1.6, 1.1)
Next j
d = d / (1.6 * M) - 1#
teeOut "rndNormal: actual / expected - 1 = " & CSng(d)
If Abs(dMax) < Abs(d) Then dMax = d

d = 0#
For j = 1& To M
  d = d + rndPoisson(1.3)
Next j
d = d / (1.3 * M) - 1#
teeOut "rndPoisson: actual / expected - 1 = " & CSng(d)
If Abs(dMax) < Abs(d) Then dMax = d

d = 0#
For j = 1& To M
  d = d + rndRicianInten(1.1, 0.2)
Next j
d = d / ((1.1 ^ 2 + 2# * 0.2 ^ 2) * M) - 1#
teeOut "rndRicianInten: actual / expected - 1 = " & CSng(d)
If Abs(dMax) < Abs(d) Then dMax = d

d = 0#
For j = 1& To M
  d = d + rndSng()
Next j
d = d / (0.5 * M) - 1#
teeOut "rndSng: actual / expected - 1 = " & CSng(d)
If Abs(dMax) < Abs(d) Then dMax = d

teeOut "Worst relative deviation from expected was " & CSng(dMax)
If Abs(dMax) > 3# / Sqr(M) Then
  N = N + 1&
  teeOut "OOPS! worst deviation unusually large " & CSng(dMax) & " > " & _
    CSng(3# / Sqr(M))
Else
  teeOut "Deviations were acceptable at " & CSng(Abs(dMax) * Sqr(M)) & _
    " times 1 / Sqr(" & Format(M, "0,0") & ") = " & CSng(1# / Sqr(M))
End If

teeOut vbNewLine & "--- sequence restart test"
k1 = 2135609784
k2 = 1234567890
rndSeedSet k1, k2
teeOut "seeds reset to: " & rndSeedGetA() & " " & rndSeedGetB()
M = 9999&
For j = 1& To M
  dMax = rndNormal()
Next j
teeOut "  after " & M & " calls rndNormal() = " & dMax
rndSeedSet k1, k2
teeOut "seeds reset to: " & rndSeedGetA() & " " & rndSeedGetB()
For j = 1& To M
  d = rndNormal()
Next j
teeOut "  after " & M & " calls rndNormal() = " & d & "  difference = " & _
  dMax - d
If dMax <> d Then _
  N = N + 1&: teeOut "OOPS! sequences did not repeat " & dMax & " <> " & d

teeOut vbNewLine & "---------- helper routines"
n0 = N
d = Exp(gammaLog(1E-300))
teeOut "Exp(gammaLog(1E-300)) = " & d & " (should be near 'infinity')"
If d < 9E+299 Then _
  N = N + 1&: teeOut "OOPS! gammaLog(0) not near infinity: " & d
d = gammaLog(1#)
teeOut "gammaLog(1) = " & d & " (should be 0.0)"
If Abs(d) >= 0.000000000000015 Then _
  N = N + 1&: teeOut "OOPS! gammaLog(1) not near 0.0: abs err " & d
d = gammaLog(2#)
teeOut "gammaLog(2) = " & d & " (should be 0.0)"
If Abs(d) >= 0.00000000000001 Then _
  N = N + 1&: teeOut "OOPS! gammaLog(2) not near 0.0: abs err " & d
d = Exp(gammaLog(3#))
teeOut "Exp(gammaLog(3)) = " & d & " (should be 2)"
If Abs(d / 2# - 1#) >= 0.00000000000001 Then _
  N = N + 1&: teeOut "OOPS! Exp(gammaLog(3)) not near 2.0: rel err " & _
  d / 2# - 1#
d = Exp(gammaLog(10#))
teeOut "Exp(gammaLog(10)) = " & d & " (should be 362880)"
If Abs(d / 362880# - 1#) >= 0.00000000000002 Then
  N = N + 1&
  teeOut "OOPS! Exp(gammaLog(10)) not near 362880: rel err " & _
    d / 362880# - 1#
End If
teeOut "gammaLog(2.55E+305) = " & gammaLog(2.55E+305) & _
  " (should be near 'infinity')"

d = gammaSign(1#)
teeOut "gammaSign(1#)) = " & d & " (should be 1)"
If d <> 1# Then _
  N = N + 1&: teeOut "OOPS! gammaSign(1#) not = 1"
d = gammaSign(0#)
teeOut "gammaSign(0#)) = " & d & " (should be 0)"
If d <> 0# Then _
  N = N + 1&: teeOut "OOPS! gammaSign(0#) not = 0"
d = gammaSign(-0.5)
teeOut "gammaSign(-0.5)) = " & d & " (should be -1)"
If d <> -1# Then _
  N = N + 1&: teeOut "OOPS! gammaSign(-0.5) not = -1"
d = gammaSign(-1#)
teeOut "gammaSign(-1#)) = " & d & " (should be 0)"
If d <> 0# Then _
  N = N + 1&: teeOut "OOPS! gammaSign(-1#) not = 0"
d = gammaSign(-1.5)
teeOut "gammaSign(-1.5)) = " & d & " (should be 1)"
If d <> 1# Then _
  N = N + 1&: teeOut "OOPS! gammaSign(-1.5) not = 1"

teeOut "time2Long(#1/1/2002#) = " & time2Long(#1/1/100#) & " (base time)"
teeOut "time2Long(#12/31/2002 23:59:59#) = " & _
  time2Long(#12/31/2002 11:59:59 PM#) & " (not a leap year)"
teeOut "time2Long(#12/31/2004 23:59:59#) = " & _
  time2Long(#12/31/2004 11:59:59 PM#) & _
  " (leap year; note 2^31 = " & 2 ^ 31 & ")"
str = "time2Long granularity:  "
For j = 1& To 10
  M = time2Long()
  Do: Loop Until M <> time2Long()
  str = str & time2Long() - M & " "
Next j
teeOut str & " 67th's of a second"
If n0 = N Then teeOut "No errors" Else teeOut "OOPS! " & N - n0 & " error(s)"

teeOut vbNewLine & "---------- speed test for rndLng()"
If inDesign() Then M = 1& Else M = 15& ' set speed: interpreted or compiled?
Dim loops As Long
loops = 100000 * M
Dim elapsed As Single
elapsed = secondsP()
For k = 1& To 20& * loops  ' do 18 assignments per loop
  j = j: j = j: j = j: j = j: j = j: j = j: j = j: j = j: j = j
  j = j: j = j: j = j: j = j: j = j: j = j: j = j: j = j: j = j
Next k
elapsed = secondsP() - elapsed
k = 18& * (k - 1&)
teeOut Format(k, "0,0") & " j = j assignments in " & _
  Format(elapsed, "0.0##") & " seconds = " & _
  Format(1000000000# * elapsed / k, "#,#.##") & " nsec / assignment"
teeOut "there are 2 calls per assignment so subtract about " & _
  Format(1000000000# * elapsed / (2# * k), "#,#.##") & " nsec / call"
elapsed = secondsP()
For k = 1& To loops   ' do 18 calls per loop; avoid overflow
  j = rndLng() - rndLng(): j = rndLng() - rndLng(): j = rndLng() - rndLng()
  j = rndLng() - rndLng(): j = rndLng() - rndLng(): j = rndLng() - rndLng()
  j = rndLng() - rndLng(): j = rndLng() - rndLng(): j = rndLng() - rndLng()
  j = rndLng() - rndLng(): j = rndLng() - rndLng(): j = rndLng() - rndLng()
  j = rndLng() - rndLng(): j = rndLng() - rndLng(): j = rndLng() - rndLng()
  j = rndLng() - rndLng(): j = rndLng() - rndLng(): j = rndLng() - rndLng()
Next k
elapsed = secondsP() - elapsed
k = 18& * loops
teeOut Format(k, "0,0") & " rndLng() calls in " & _
  Format(elapsed, "0.0##") & " seconds = " & _
  Format(1000000000# * elapsed / k, "#,#.##") & " nsec / call"

teeOut vbNewLine & "---------- speed test for rndSng()"
loops = 180000 * M
Dim v As Single
elapsed = secondsP()
For k = 1& To 20& * loops  ' do 18 assignments per loop
  v = v: v = v: v = v: v = v: v = v: v = v: v = v: v = v: v = v
  v = v: v = v: v = v: v = v: v = v: v = v: v = v: v = v: v = v
Next k
elapsed = secondsP() - elapsed
k = 18& * (k - 1&)
teeOut Format(k, "0,0") & " v = v assignments in " & _
  Format(elapsed, "0.0##") & " seconds = " & _
  Format(1000000000# * elapsed / k, "#,#.##") & " nsec / assignment"
teeOut "there are 18 calls per assignment so subtract about " & _
  Format(1000000000# * elapsed / (18# * k), "#,#.##") & " nsec / call"
elapsed = secondsP()
For k = 1& To loops   ' do 18 calls per loop
  v = rndSng() - rndSng() + rndSng() - rndSng() + rndSng() - rndSng() _
    + rndSng() - rndSng() + rndSng() - rndSng() + rndSng() - rndSng() _
    + rndSng() - rndSng() + rndSng() - rndSng() + rndSng() - rndSng()
Next k
elapsed = secondsP() - elapsed
k = 18& * loops
teeOut Format(k, "0,0") & " rndSng() calls in " & _
  Format(elapsed, "0.0##") & " seconds = " & _
  Format(1000000000# * elapsed / k, "#,#.##") & " nsec / call"

teeOut vbNewLine & "---------- speed test for rndDbl()"
loops = 100000 * M
elapsed = secondsP()
For k = 1& To 20& * loops  ' do 18 assignments per loop
  d = d: d = d: d = d: d = d: d = d: d = d: d = d: d = d: d = d
  d = d: d = d: d = d: d = d: d = d: d = d: d = d: d = d: d = d
Next k
elapsed = secondsP() - elapsed
k = 18& * (k - 1&)
teeOut Format(k, "0,0") & " d = d assignments in " & _
  Format(elapsed, "0.###") & " seconds = " & _
  Format(1000000000# * elapsed / k, "#,#.##") & " nsec / assignment"
teeOut "there are 18 calls per assignment so subtract about " & _
  Format(1000000000# * elapsed / (18# * k), "#,#.##") & " nsec / call"
elapsed = secondsP()
For k = 1& To loops    ' do 18 calls per loop
  d = rndDbl() - rndDbl() + rndDbl() - rndDbl() + rndDbl() - rndDbl() _
    + rndDbl() - rndDbl() + rndDbl() - rndDbl() + rndDbl() - rndDbl() _
    + rndDbl() - rndDbl() + rndDbl() - rndDbl() + rndDbl() - rndDbl()
Next k
elapsed = secondsP() - elapsed
k = 18& * (k - 1&)
teeOut Format(k, "0,0") & " rndDbl() calls in " & _
  Format(elapsed, "0.###") & " seconds = " & _
  Format(1000000000# * elapsed / k, "#,#.##") & " nsec / call"

str = vbNewLine & "~~~~~ " & Module_c & " unit test done: " & N & " error"
If N <> 1 Then str = str & "s"
teeOut str & " ~~~~~ end of file ~~~~~"

Close ofi_m
End Sub

'&&&&& inDesign &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Private Function inDesign() _
As Boolean
' Returns True if program is running in IDE (editor) design environment, and
' False if program is running as a standalone EXE. Useful for "hooking" only
' when standalone, or adjusting for the speed difference between compiled and
' interpreted. So in your program you can say: if [Not] inDesign() Then ...
'         John Trenholme - 2009-10-21
inDesign = False
On Error Resume Next  ' set to ignore error in Assert
Debug.Assert 1& \ 0&  ' attempts this illegal feat only in IDE
If 0& <> Err.Number Then
  inDesign = True  ' comment this out to get compiled behavior while in IDE
  Err.Clear  ' do not pass divide-by-zero error back up to caller (yes, it can!)
End If
End Function

'&&&&& rd &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Private Function rd( _
  ByVal hiLng As Long, _
  loLng As Long) _
As Double
' This produces the same result as rndDbl(), but with specified integers.
' Make sure that the expression here matches the one at the end of rndDbl().
' Unit-test support routine - John Trenholme - 2002-02-04
rd = AddD_c + hiLng * MulHiD_c + loLng * MulLoD_c
End Function

'&&&&& secondsP &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Private Function secondsP() As Double
' Return number of seconds since first call to this routine. A return of -86400
' indicates an error. The granularity will be around a microsecond.
' Unit-test support routine - John Trenholme - 2007-01-23
Static s_base As Currency  ' initializes to 0
Static s_freq As Currency  ' initializes to 0
Const c_Default As Double = -86400#  ' 1 day in seconds, negated
If s_freq = 0@ Then  ' routine not initialized, or unable to read frequency
  QueryPerformanceFrequency s_freq  ' try to read frequency
  ' if frequency is good, try to read base time (else it stays at 0)
  If s_freq <> 0@ Then QueryPerformanceCounter s_base
End If
' if we have a good base time, then we must have a good frequency also
If s_base <> 0@ Then
  Dim time As Currency
  QueryPerformanceCounter time
  If time <> 0@ Then
    secondsP = (time - s_base) / s_freq
  Else
    secondsP = c_Default
  End If
Else  ' something is wrong - return error value
  secondsP = c_Default
End If
End Function

'&&&&& teeOut &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Private Sub teeOut( _
ByRef str As String)
' Send the supplied string to the output file (if it is open) and to the
' Immediate window (Ctrl-G to open) if in VB editor.
' Unit-test support routine - John Trenholme - 9 Jul 2002
Debug.Print str  ' send to Immediate window (only if in Editor; limited size)
If 0 <> ofi_m Then Print #ofi_m, str
End Sub

#End If

'-------------------------------- end of file ----------------------------------

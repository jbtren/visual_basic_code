VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RandomNumbers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "This class supplies a number of pseudo-random sequences, including  uniformly-distributed Singles & Doubles, Normal variates, and others. Devised and coded by John Trenholme."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'
'###############################################################################
'#
'# Visual Basic 6 (or VBA) class file "RandomNumbers.cls"
'#
'# Class to supply pseudo-random number sequences.
'# Devised and coded by John Trenholme. Begun 20 Jul 2004.
'#
'###############################################################################
'#
'# Example usage:
'#
'# Dim rn As RandomNumbers
'# Set rn = New RandomNumbers
'#
'# rn.Seed1 = rn.Time2Long()  ' randomize the sequence
'#
'# ... use some sequence values ...
'# speed = rn.Normal(1.1, 0.3)
'# hits = rn.Poisson(2.345)
'#
'# Set rn = Nothing  ' release memory used by RandomNumbers object
'#
'###############################################################################

Option Explicit

Private Const c_version As String = "2004-07-27"

' This class was written because the built-in pseudo-random function Rnd() in
' VB has too short a cycle length (only 2^24 or 16,777,216 numbers) to be
' useful in a serious Monte Carlo job. It also has the serious defect that
' succesive N-tuples lie on planes in the unit hypercube. Also, it has only
' single-precision resolution. A longer-period generator of known high quality
' and higher precision is needed. This class supplies that generator.
'
' Note that the random number generator in Excel is also deficient, and should
' be replaced by these routines (use the "Random.bas" module to get functions).
'
' You can now replace VB's Rnd() with Sng(), and replace Randomize, with its
' confusing and inconsistent behavior, with Seed1 and Seed2 (using Time2Long()
' as one of the arguments if you want different results on every run).
'
' You also get the double-precision pseudo-random sequence Dbl(), and
' routines to produce pseudo-random sequences drawn from a number of common
' probability distributions. The sequences provided are:
'
'   Dbl()          uniform double-precision variates  0.0 < Dbl() < 1.0
'   Exponential()  exponentially-distributed double-precision variates
'   Gumbel()       Gumbel extreme-value double-precision variates
'   LogNormal()    log-normal double-precision variates
'   Lng()          uniform long-integer values  0 <= Lng() <= 2147483562
'                       note: Lng's upper limit is returned by LngMax()
'   Normal()       normal (or Gaussian) double-precision variates
'   Poisson()      Poisson long-integer variates
'   RicianInten()  Rician intensity double-precision variates
'   Sng()          uniform single-precision variates  0.0 < Sng() < 1.0
'
' All these sequences are derived from one underlying generator whose starting
' point is based on two Long values, called seeds. If you only want to use one
' seed, set the other to any arbitrary constant value, or leave it at the
' default initial value. You set and read the seed values by use of the class
' Properties Seed1 and Seed2:
'
'   rn.Seed1 = some Long value
'   someLong = rn.Seed2()
'
' If you save the seed values before setting them, and later set them again
' with the saved values, the exact same sequence of pseudo-random variates
' will be supplied (assuming you then make the same calls in the same order).
'
' Two support routines used in the class are made publicly available. Their
' results are independent of class values (they are "free" functions):
'
'   GammaLog(x)       logarithm of the Gamma function (used by Poisson)
'   Time2Long()       count of 67th-of-a-second intervals since New Year's Day
'                     this can be used to set seed(s) and "randomize" results
'                     the way Randomize(Timer) does in VB

' ===== Details of the pseudo-random generator and routines ====================
'
' The basis for good pseudo-random number sequence generators is often a good
' integer sequence generator. Fortunately, many superior integer generators are
' available. Some even work with signed 31-bit integers (i.e., Longs), which VB
' is restricted to. In this module, we use a pair of portable pseudo-random
' integer generators for 31-bit signed integer arithmetic that work without
' overflow of intermediate results (another VB restriction). The generators are
' slightly modified from the example in "Efficient and Portable Random Number
' Generators" by Pierre L'Ecuyer, CACM 31, 742(1988).  The changes are:
'   1) the addition of constants in the evaluations to allow 0 seeds
'   2) adjustment of the range of the uniform floating-point variates generated
'      from the integers to exclude the end values 0.0 and 1.0
'   3) returning the Single and Long corresponding to the sequence integers
'      before they are updated, rather than after.
'
' The two generators are combined by modulo addition. Their periods are mutually
' prime (in fact, both periods are prime) and so the combined period is the
' product of the two periods, giving about 4.6E18 nearly-31-bit integers until
' the sequence repeats (enough for 1 per nanosecond for more than 145 years).
' The combined integers are in turn used to generate uniformly-distributed
' Single and Double floating-point values, and the Double values are used to
' generate non-uniform continuous and discrete pseudo-random distributions.
' Because of the two-generator combination, and the fact that the low-order
' bits in the resulting sequence are discarded when making Single and Double
' values, there is no lie-on-hyperplane problem.
'
' These sequences pass most, but not all, of Marsaglia's "torture tests."
' To pass all of them, we would have to add a 4-long shuffle table, but the
' resulting improvement for very, very few problems was judged to not be worth
' the slowdown for practically all applications. And, nobody's perfect - just
' remember Don Knuth's maxim: "Every random number generator will fail in at
' least one application." (Seminumerical Algorithms, section 3.6)
'
' The present point in the generator cycle is shared, and advanced, by all the
' sequences. A call to any routine advances the sequence by one step, two steps,
' or (in some cases) more than two steps. Be careful not to disturb an order
' you depend on by calling another random-variate routine. If in doubt, get
' sequences from different objects.

' ==============================================================================
'
' Warning! Adjust these constants only if you know what you're doing!
'
' The first integer generator is R = ( 40014 * R + 1693666889) mod 2147483563
' The modulus is 2^31 - 85, the 5th prime below 2^31.
Private Const c_ra1 As Long = 2147483563
Private Const c_ra2 As Long = 40014
Private Const c_ra3 As Long = 1693666889 - c_ra1
Private Const c_ra4 As Long = c_ra1 \ c_ra2
Private Const c_ra5 As Long = c_ra1 - c_ra2 * c_ra4
' The second integer generator is R = ( 40692 * R + 1) mod 2147483399
' The modulus is 2^31 - 249, the 12th prime below 2^31.
Private Const c_rb1 As Long = 2147483399
Private Const c_rb2 As Long = 40692
Private Const c_rb3 As Long = 1
Private Const c_rb4 As Long = c_rb1 \ c_rb2
Private Const c_rb5 As Long = c_rb1 - c_rb2 * c_rb4
' Integers are scaled into the Single range 0 < variate < 1 by this multiplier
Private Const c_scaleS As Single = 1! / c_ra1
' Integers are scaled into the Double range 0 < variate < 1 by these multipliers
' The peculiar multiplier is due to the fact that we supply slightly fewer than
' 31 bits (i.e., log(2147483563) / log(2) = 30.9999999428964)
Private Const c_scaleDhi As Double = 1# / c_ra1
Private Const c_scaleDlo As Double = 2# ^ (-31.0000001) / c_ra1

' Default seed values - used when generators are called before initialization
Private Const c_defaultSeed1 As Long = 543210
Private Const c_defaultSeed2 As Long = 1211109876

' Module-global variables (initialized as 0, "" or False when module starts)
Private m_rnd1 As Long         ' the present point in the sequence is held ...
Private m_rnd2 As Long         ' ... in m_rnd1 and m_rnd2
Private m_normSaved As Boolean ' True if a Normal() value is being saved
Private m_PoissonCalls As Long ' number of calls to Dbl() in Poisson()

' Variables used to support unit tests
Private m_test As Boolean      ' flag to indicate if this is only a test
Private m_testVal As Double    ' value to be returned by Dbl() if testing

' Double precision constants have to be written as a sum in Visual Basic
' (VB[A] will otherwise truncate a digit when module is saved to file & loaded)
Private Const c_Pi As Double = 3.14159265 + 3.5897931E-09

'---------- Event --------------------------------------------------------------
Private Sub Class_Initialize()
' set sequence to known-good default starting point
m_rnd1 = Me.Seed1Default()
m_rnd2 = Me.Seed2Default()
End Sub

'========== Property ===========================================================
Property Get GammaLog(ByVal x As Double) As Double
Attribute GammaLog.VB_Description = "Logarithm of the Gamma function g(x) = (x-1)! for positive arguments, good to 14 digits. To get Gamma(x), use Exp(GammaLog(x)), but that will overflow if x > 171.6"
' Logarithm of the Gamma function g(x) = (x-1)! for positive arguments.
'
' To get Gamma(x), use Exp(GammaLog(x)), but that will overflow if x > 171.6
'
' For x < 0.5, you can use the relation:
'   Gamma( x) = Pi / ( Gamma( 1.0 - x) * Sin( Pi * x))
' Of course, you still have to avoid the singularities at x = 0, -1, -2, ...
'
' Based on a Lanczos-type-series routine by Alan Miller of CSIRO.
' Coefficients tweaked to be good to 14 digits for most "reasonable" inputs.
Const c_xMin As Double = 3.8E-306

If x < c_xMin Then x = c_xMin  ' avoid negative arguments, and stop overflow
GammaLog = Log(1# + 676.520368121884 / x - 1259.13921672248 / (x + 1#) + _
  771.323428780167 / (x + 2#) - 176.615029181445 / (x + 3#) + _
  12.5073433413946 / (x + 4#) - 0.13857119516943 / (x + 5#) + _
  1.006163373865E-05 / (x + 6#) + 1.27346548454E-07 / (x + 7#)) - _
  5.58106146679531 - x + (x - 0.5) * Log(x + 6.5)
End Property

'========== Property ===========================================================
Property Get Dbl() As Double
Attribute Dbl.VB_Description = "One value from a pseudo-random sequence of Double variates, uniformly distributed between 0 and 1 (non-inclusive)."
' Returns a sequence of Double precision pseudo-random variates uniformly
' distributed between zero and one, not including either of those values. The
' smallest value returned is 2**(-62) or 2.17e-19, and the largest is
' 1 - 2**(-53) or 0.9999999999999998. There are about 9.0e15 distinct output
' values, but the sequence does not repeat until called about 4.6e18 times. The
' integer sequence generator is advanced twice in this routine.
' You ought to set Seed1 and Seed2 before using any sequence values.
Dim jTemp As Long
Dim d As Double

If m_test Then  ' we are doing tests and want to force a specific Dbl value
  Dbl = m_testVal
  Exit Property
End If

' combine the present integers and put into valid range
jTemp = m_rnd1 - m_rnd2
If jTemp < 0 Then jTemp = jTemp + c_ra1

' scale into a Double number 0.0 <= Dbl < 1.0, giving 31 random bits
' the result may be zero, but that will be fixed in the next step
' Note: c_scaleDhi and c_scaleDlo *must* be sized such that
' ( c_ra1 - 1&) * c_scaleDhi + c_ra1 * c_scaleDlo < 1!
d = jTemp * c_scaleDhi

' do mixed congruential generator #1 to update integer #1
' uses no-overflow signed-integer trick due to Schrage:
' see L. Schrage, ACM Trans. Math. Soft. 5 (1979), pp. 132-138 and
' Bratley, Fox & Schrage, "A Guide to Simulation" 2d edition,
' Springer-Verlag, 1987
jTemp = m_rnd1 \ c_ra4
m_rnd1 = c_ra2 * (m_rnd1 - jTemp * c_ra4) - jTemp * c_ra5 + c_ra3
If m_rnd1 < 0 Then m_rnd1 = m_rnd1 + c_ra1

' do mixed congruential generator #2 to update integer #2 (see ref. above)
jTemp = m_rnd2& \ c_rb4
m_rnd2 = c_rb2 * (m_rnd2 - jTemp * c_rb4) - jTemp * c_rb5 + c_rb3
If m_rnd2 < 0 Then m_rnd2 = m_rnd2 + c_rb1

' combine the present integers and put into valid range
jTemp = m_rnd1 - m_rnd2
If jTemp < 0 Then jTemp = jTemp + c_ra1

' add another 22 bits onto the bottom end of the Double
' be sure that 0 cannot occur by adding 1, and make up for the fact that there
' were slightly fewer than 31 bits added in the first step
Dbl = d + (jTemp + 1&) * c_scaleDlo

' do mixed congruential generator #1 to update integer #1 (see ref. above)
jTemp = m_rnd1 \ c_ra4
m_rnd1 = c_ra2 * (m_rnd1 - jTemp * c_ra4) - jTemp * c_ra5 + c_ra3
If m_rnd1 < 0 Then m_rnd1 = m_rnd1 + c_ra1

' do mixed congruential generator #2 to update integer #2 (see ref. above)
jTemp = m_rnd2& \ c_rb4
m_rnd2 = c_rb2 * (m_rnd2 - jTemp * c_rb4) - jTemp * c_rb5 + c_rb3
If m_rnd2 < 0 Then m_rnd2 = m_rnd2 + c_rb1
End Property

'========== Property ===========================================================
Property Get Exponential(Optional ByVal meanValue As Double = 1#) As Double
Attribute Exponential.VB_Description = "One value from a pseudo-random sequence of Double variates, exponentially distributed with the supplied mean value (default 1)."
' Returns a sequence of pseudo-random variates from an exponential distribution
' of specified mean. If no mean is supplied, variates with unit mean are
' produced. The smallest value returned is 1.11e-16, and the largest is 42.975.
' Although "meanValue" ought to be strictly positive, it causes no harm to use
' zero or negative values - you just get back a zero or negative variate.
' You ought to set Seed1 and Seed2 before using any sequence values.
' Uses Dbl() as a source of pseudo-random uniform variates, calling it once.
Exponential = -meanValue * Log(Me.Dbl())
End Property

'========== Property ===========================================================
Property Get Gumbel(Optional ByVal meanValue As Double = 0#, _
                    Optional ByVal stdDev As Double = 1#) _
             As Double
Attribute Gumbel.VB_Description = "One value from a pseudo-random sequence of Double variates, Gumbel distributed with the supplied mean (default 0) and standard deviation (default 1)."
' Returns a sequence of pseudo-random variates from a Gumbel (AKA extreme-value,
' double-exponential, Fisher-Tippett type 1) distribution of specified mean and
' standard deviation. If no mean is supplied, a variate with zero mean is
' produced. If no standard deviation is supplied, a variate with unit standard
' deviation is produced. The smallest value returned is -2.382, and the largest
' is 29.19. If a negative standard deviation is specified, the resulting
' variates will have the long exponential tail extending to smaller values,
' rather than larger values, contrary to the usual definition.
' You ought to set Seed1 and Seed2 before using any sequence values.
' Uses Dbl() as a source of pseudo-random uniform variates, calling it once.
Const c_root6overPi As Double = 0.7796968 + 1.2336761E-09 ' Sqr(6) / Pi
Const c_eulerMascheroni As Double = 0.57721566 + 4.90153286E-09

Gumbel = meanValue - stdDev * c_root6overPi * (Log(-Log(Me.Dbl())) + _
  c_eulerMascheroni)
End Property

'========== Property ===========================================================
Property Get LogNormal(Optional ByVal meanValue As Double = 1#, _
                       Optional ByVal stdDev As Double = 1#) _
             As Double
Attribute LogNormal.VB_Description = "One value from a pseudo-random sequence of Double variates, log-normally distributed with the supplied mean (default 1) and standard deviation (default 1)."
' Returns a sequence of pseudo-random variates from a log-normal distribution of
' specified mean and standard deviation. If no mean is supplied, a variate with
' unit mean is produced. If no standard deviation is supplied, a variate with
' unit standard deviation is produced. The input argument "meanValue" should be
' strictly positive, so if it is not it is silently made so (if you want the
' long tail to extend to smaller values, negate the return value).
' You ought to set Seed1 and Seed2 before using any sequence values.
'
' Note that many authors define this distribution in terms of the mean and
' standard deviation of the Gaussian in the exponent, contrary to our usage.
'
' Uses Normal() as a source of pseudo-random normal variates, calling it once.
Dim t As Double

If meanValue < 1E-101 Then meanValue = 1E-101  ' an arbitrary lower limit
t = 1# + (stdDev / meanValue) ^ 2
LogNormal = Exp(Me.Normal(Log(meanValue / Sqr(t)), Sqr(Log(t))))
End Property

'========== Property ===========================================================
Property Get Lng() As Long
Attribute Lng.VB_Description = "One value from a pseudo-random sequence of Long variates, uniformly distributed between 0 and 2147483562."
' Returns a sequence of pseudo-random Long integers from a uniform distribution
' between 0 and LngMax(), inclusive (note that LngMax() = 2147483562 here).
' The integer sequence generator is advanced once in this routine.
' You ought to set Seed1 and Seed2 before using any sequence values.
Dim jTemp As Long

' combine the present integers and put into valid range
jTemp = m_rnd1 - m_rnd2
If jTemp < 0 Then jTemp = jTemp + c_ra1

' prepare to return the result
Lng = jTemp

' do mixed congruential generator #1 to update integer #1
' see comments in body of "Dbl()" for an explanation of what's happening
jTemp = m_rnd1 \ c_ra4
m_rnd1 = c_ra2 * (m_rnd1 - jTemp * c_ra4) - jTemp * c_ra5 + c_ra3
If m_rnd1 < 0 Then m_rnd1 = m_rnd1 + c_ra1

' do mixed congruential generator #2 to update integer #2
jTemp = m_rnd2& \ c_rb4
m_rnd2 = c_rb2 * (m_rnd2 - jTemp * c_rb4) - jTemp * c_rb5 + c_rb3
If m_rnd2 < 0 Then m_rnd2 = m_rnd2 + c_rb1
End Function

'========== Property ===========================================================
Property Get LngMax() As Long
Attribute LngMax.VB_Description = "The largest value that Lng() will return (the minimum is 0)."
' The largest value that Lng() will return (the minimum is 0).
LngMax = c_ra1 - 1&
End Property

'========== Property ===========================================================
Property Get Normal(Optional ByVal meanValue As Double = 0#, _
                    Optional ByVal stdDev As Double = 1#) _
             As Double
Attribute Normal.VB_Description = "One value from a pseudo-random sequence of Double variates, normally distributed with the supplied mean (default 0) and standard deviation (default 1). Note that the defaults give a ""standard normal"" variate."
' Returns a sequence of pseudo-random variates from a normal (or Gaussian)
' distribution of specified mean and standard deviation. If no mean is supplied,
' a variate with zero mean is produced. If no standard deviation is supplied, a
' variate with unit standard deviation is produced. Note that a call with no
' arguments returns a "standard normal" variate. The minimum value produced
' is -8.531 and the maximum is 8.450, but such extreme values are extremely
' unlikely (the probability is a bit less than 1 per billion of a value above 6,
' and the same for a value below -6).
' You ought to set Seed1 and Seed2 before using any sequence values.
'
' To produce two variates "v1" and "v2" with means "m1" and "m2," standard
' deviations "s1" and "s2," and correlation coefficient "c," do this (Knuth,
' "Seminumerical Algorithms" section 3.4.1(5) & exercise 13):
'
' temp = X.Normal()
' v1 = m1 + s1 * temp
' v2 = m2 + s2 * (c * temp + Sqr(1# - c * c) * x.Normal())
'
' Of course, you have to have -1.0 <= c <= 1.0 for this to work.
'
' Uses Dbl() as a source of pseudo-random uniform variates, calling it
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
' behavior. If you save the seed values at some point where an odd number of
' calls have been made to this routine, and later set the seeds to the saved
' values, the sequence past the seed-save point will be different. To avoid
' this, immediately set a seed values to what you just got from either Seed1()
' or Seed2() when you do the save, before proceeding with your code.
' This will discard any cached value, and make the sequence after the save point
' equal to the one after the restore point. That is:
'   save1 = X.Seed1()
'   save2 = X.Seed2()
'   X.Seed1 = save1  ' do this to purge any cached values
'     ... do something with a sequence of rndNormal() values ...
'   X.Seed1 = save1
'   X.Seed2 = save2
'     ... do something with the same sequence of rndNormal() values ...
Static saved As Double  ' holds saved second value, if any
Dim n As Long
Dim t As Double
Dim x As Double

If Not m_normSaved Then
  ' make up two Gaussian variates of zero mean and unit standard deviation
  ' the loop will be executed 4 / Pi = 1.2732 times on average
  n = 0&
  Do
    x = Me.Dbl() - 0.5
    saved = Me.Dbl() - 0.5
    t = x * x + saved * saved
    n = n + 1&
  Loop Until ((t <= 0.25) And (t > 0#)) Or (n >= 100&)
  ' avoid possible infinite loops during testing (see m_testVal)
  If n = 100& Then
    x = 0.2
    saved = 0.3
    t = x * x + saved * saved
  End If
  ' point is uniformly distributed inside disc of radius 0.25 when we exit
  t = Sqr(-2# * Log(4# * t) / t)  ' transform uniform disk to Gaussian
  saved = saved * t  ' make saved value a standard normal variate
  Normal = meanValue + x * t * stdDev  ' set desired return value
  m_normSaved = True
Else
  Normal = meanValue + saved * stdDev  ' convert saved value
  m_normSaved = False
End If
End Property

'========== Property ===========================================================
Property Get Poisson(Optional ByVal meanValue As Double = 1#) As Long
Attribute Poisson.VB_Description = "One value from a pseudo-random sequence of Long variates, Poisson distributed with the supplied mean (default 1)."
' Returns a sequence of pseudo-random variates from a Poisson distribution of
' specified mean. If no mean is supplied, a variate with unit mean is produced.
' If "meanValue" is negative, it is silently set to zero, and zero is returned.
' If you want negative Poisson values, negate the result of this routine.
' You ought to set Seed1 and Seed2 before using any sequence values.
'
' Based on the routine in Numerical Recipes.
' Uses Dbl() as a source of pseudo-random uniform variates. See comments
' below for the number of times it is called.
Dim ex As Double
Dim g As Double
Dim logMean As Double
Dim sq As Double
Dim t As Double
Dim y As Double
Dim k As Long

m_PoissonCalls = 0&
If meanValue <= 0# Then  ' force a default for impossible input
  Poisson = 0&
  Exit Property
End If

If meanValue <= 12# Then
  ' small mean; use sum of exponential variates
  ' will call Dbl() a total of meanValue + 1 times on average
  ex = Exp(-meanValue)
  t = Me.Dbl()
  m_PoissonCalls = m_PoissonCalls + 1&
  k = 0
  Do While t > ex
    t = t * Me.Dbl()
    m_PoissonCalls = m_PoissonCalls + 1&
    k = k + 1
  Loop
  Poisson = k
Else
  ' large mean; use the rejection method
  ' will call Dbl() around 3.8 to 3.9 times on average
  sq = Sqr(2# * meanValue)
  logMean = Log(meanValue)
  g = meanValue * logMean - GammaLog(meanValue + 1#)
  Do  ' try until we are in the acceptance region
    Do  ' get a positive value from the Lorentzian distribution
      y = Tan(c_Pi * Me.Dbl())
      m_PoissonCalls = m_PoissonCalls + 1&
      t = meanValue + sq * y  ' set mean & variance to enclose Poisson
    Loop Until t >= 0#
    t = Int(t)  ' move down to the first integer below trial value
    m_PoissonCalls = m_PoissonCalls + 1&
  Loop Until Me.Dbl() < _
             0.9 * (1# + y * y) * Exp(t * logMean - GammaLog(t + 1#) - g)
  Poisson = t
End If
End Property

'========== Property ===========================================================
Property Get PoissonDblCalls() As Long
Attribute PoissonDblCalls.VB_Description = "Number of calls to Dbl() used during most recent call to Poisson(). Supplied for testing."
' Number of calls to Dbl() used during most recent call to Poisson().
PoissonDblCalls = m_PoissonCalls
End Property

'========== Property ===========================================================
Property Get RicianInten(ByVal constField As Double, _
                         ByVal addedStdDev As Double) _
             As Double
Attribute RicianInten.VB_Description = "One value from a pseudo-random sequence of Double variates, Rician intensity distributed based on the sum of a constant field and a stochastic phasor field with the supplied standard deviation in both in-phase & quadrature components."
' Returns a sequence of pseudo-random variates from a Rician intensity
' distribution (also called modified Rician distribution), given field values
' of random process.
'
' Input values are:
'   constField  = field value of constant phasor
'   addedStdDev = field standard deviation of added random phasors
'
' Note: if mu is the mean intensity and cn is the "contrast" = std.dev. / mean,
' then the input field quantities are given by:
'
'   constField = Sqr(mu * Sqr(1# - cn * cn))
'   addedStdDev = Sqr((1# - Sqr(1# - cn * cn)) * mu / 2#)
'
' In terms of the constant field f & the added std. dev. s, the mean intensity
' and contrast are given by:
'
'    mu = f ^ 2 + 2# * s ^ 2
'    cn = 2# * s * Sqr(f ^ 2 + s ^ 2) / (f ^ 2 + 2# * s ^ 2)
'
' You ought to set Seed1 and Seed2 before using any sequence values.
'
' Uses Normal() as a source of pseudo-random uniform variates, calling it twice
Dim x As Double
Dim y As Double

' total in-phase field is sum of constant field and added in-phase component
x = Me.Normal(Abs(constField), addedStdDev)
' quadrature field is just added quadrature component
y = Me.Normal(0#, addedStdDev)
' add squares of in-phase and quadrature fields to get intensity
RicianInten = x * x + y * y
End Property

'========== Property ===========================================================
Property Get ScaleDlo() As Double
Attribute ScaleDlo.VB_Description = "One of two scaling coefficients used in turning pseudo-random integers into Doubles. Supplied for testing."
' One of two scaling coefficients used in turning pseudo-random integers into
' Doubles.
ScaleDlo = c_scaleDlo
End Property

'========== Property ===========================================================
Property Get ScaleDhi() As Double
Attribute ScaleDhi.VB_Description = "One of two scaling coefficients used in turning pseudo-random integers into Doubles. Supplied for testing."
' One of two scaling coefficients used in turning pseudo-random integers into
' Doubles.
ScaleDhi = c_scaleDhi
End Property

'========== Property ===========================================================
Property Get ScaleS() As Double
Attribute ScaleS.VB_Description = "Scaling coefficient used in turning pseudo-random integers into Singles. Supplied for testing."
' Scaling coefficient used in turning pseudo-random integers into Singles.
ScaleS = c_scaleS
End Property

'========== Property ===========================================================
' Setting Seed1 and/or Seed2 initializes the random number generator to a known
' point in the pseudo-random sequence. It also purges any cached values in the
' pseudo-random routines. Seed1 is interpreted modulo 2147483563, and Seed2 is
' interpreted modulo 2147483399.
'
' For a known-good sequence, you can set seed1 = 543210 and seed2 = 1211109876:
'    X.Seed1 = 543210
'    X.Seed2 = 1211109876
' Note: these values are supplied as defaults when the object is created, so
' you can just do nothing and accept the defaults. To restart again, you can:
'    X.Seed1 = X.Seed1Default()
'    X.Seed2 = X.Seed2Default()
'
' For independent sequences, set Seed1 = 543210 (etc.) with Seed2 = 0, 1, 2, ...
'
' To do the same thing that "Randomize" with no argument does for the built-in
' generator "Rnd()" you need a way to turn the time "right now" into a seed
' value for these routines. That is accomplished by use of the routine
' "Time2Long" below. Do this:
'    X.Seed1 = rn.Time2Long()

Public Property Get Seed1() As Long
Attribute Seed1.VB_Description = "One of two values used to specify the present position in the underlying integer sequence. See also Seed2."
' No matter what you set Seed1 to, you get back 0 <= Seed1 < c_ra1
Seed1 = m_rnd1
End Property

Public Property Let Seed1(ByVal newSeed As Long)
' modulus arithmetic - force internal seed into range 0 <= m_rnd1 < c_ra1
m_rnd1 = newSeed
Do While m_rnd1 < 0
  m_rnd1 = m_rnd1 + c_ra1
Loop
Do While m_rnd1 >= c_ra1
  m_rnd1 = m_rnd1 - c_ra1
Loop
' throw out any cached values, so next call will reflect new seed values
m_normSaved = False
End Property

Public Property Get Seed2() As Long
Attribute Seed2.VB_Description = "One of two values used to specify the present position in the underlying integer sequence. See also Seed1."
' No matter what you set Seed2 to, you get back 0 <= Seed2 < c_ra2
Seed2 = m_rnd2
End Property

Public Property Let Seed2(ByVal newSeed As Long)
' modulus arithmetic - force internal seed into range 0 <= m_rnd2 < c_ra2
m_rnd2 = newSeed
Do While m_rnd2 < 0
  m_rnd2 = m_rnd2 + c_ra2
Loop
Do While m_rnd2 >= c_ra2
  m_rnd2 = m_rnd2 - c_ra2
Loop
' throw out any cached values, so next call will reflect new seed values
m_normSaved = False
End Property

'========== Property ===========================================================
Property Get Seed1Default() As Long
Attribute Seed1Default.VB_Description = "Default value used for Seed1 if the user has not set it."
' The default value for Seed1
Seed1Default = c_defaultSeed1
End Property

'========== Property ===========================================================
Property Get Seed2Default() As Long
Attribute Seed2Default.VB_Description = "Default value used for Seed2 if the user has not set it."
' The default value for Seed2
Seed2Default = c_defaultSeed2
End Property

'========== Property ===========================================================
Property Get Sng() As Single
Attribute Sng.VB_Description = "One value from a pseudo-random sequence of Single variates, uniformly distributed between 0 and 1 (non-inclusive)."
' Returns a sequence of Single precision pseudo-random variates uniformly
' distributed between zero and one, not including either. The smallest value
' returned is 2**(-31) or 4.65e-10, and the largest is 1 - 2**(-24) or
' 0.99999994. There are about 1.7e7 distinct output values, but the sequence
' does not repeat until called about 4.6e18 times. If you need finer-grained
' values, use "Dbl()".  The integer sequence generator is advanced once in
' this routine.
' You ought to set Seed1 and Seed2 before using any sequence values.
Dim jTemp As Long

' combine the present integers and put into valid range
jTemp = m_rnd1 - m_rnd2
If jTemp < 0 Then jTemp = jTemp + c_ra1

' scale into a Single number 0.0 < rndSng < 1.0 and prepare to return it
' Note: c_scaleS *must* be sized so c_ra1 * c_scaleS < 1!
Sng = (jTemp + 1&) * c_scaleS

' do mixed congruential generator #1 to update integer #1
' see comments in body of "Dbl()" for an explanation of what's happening
jTemp = m_rnd1 \ c_ra4
m_rnd1 = c_ra2 * (m_rnd1 - jTemp * c_ra4) - jTemp * c_ra5 + c_ra3
If m_rnd1 < 0 Then m_rnd1 = m_rnd1 + c_ra1

' do mixed congruential generator #2 to update integer #2
jTemp = m_rnd2& \ c_rb4
m_rnd2 = c_rb2 * (m_rnd2 - jTemp * c_rb4) - jTemp * c_rb5 + c_rb3
If m_rnd2 < 0 Then m_rnd2 = m_rnd2 + c_rb1
End Property

'========== Property ===========================================================
Property Get Test() As Boolean
Attribute Test.VB_Description = "Set this True to force Dbl() to always return TestVal. Supplied for testing."
' When this is True, the Dbl() routine always returns TestVal (below), and all
' routines that use Dbl() follow along.
Test = m_test
End Property

Property Let Test(ByVal newValue As Boolean)
m_test = newValue
End Property

'========== Property ===========================================================
Property Get TestVal() As Double
Attribute TestVal.VB_Description = "When Test is True, Dbl() always returns this value. Supplied for testing."
' The value that Dbl() always returns when Test (above) is True.
TestVal = m_testVal
End Property

Property Let TestVal(ByVal newValue As Double)
m_testVal = newValue
End Property

'========== Property ===========================================================
Property Get Time2Long(Optional ByVal time As Variant = 0&) As Long
Attribute Time2Long.VB_Description = "Converts a time to a Long that holds the number of 67ths of a second between that time and the start of the time's year. If no argument is supplied, ""Now()"" is used. Useful for ""randomizing"" seeds."
' Convert a time to a long that holds the number of 67ths of a second between
' the supplied time and the start of the year specified in the supplied time.
' The minimum value returned is 0, at midnight on Jan 1 (#1/1/YYYY#).
' The maximum value returned is 2,112,912,000 (or 2,118,700,800 during leap
' years). This routine is supplied in this module so that the Visual Basic
' routine "Randomize" (or "Randomize Timer") can be replaced by, and improved
' upon, by:
'    X.Seed1 = X.Time2Long()
'
' If no argument is supplied, "Now()" is used as the time. On a standard PC,
' the result will change by 2 or 3 units as time proceeds because of the
' granularity of the "Timer" function. If an argument is supplied, the result
' changes only every second, as fractional seconds are not available in a Date.
'
' The apparently peculiar choice of 67ths of a second was made so that the
' result would be close to 2^31 (the largest possible Long) at the end of
' a leap year, thus supplying almost all possible seed values.
'
Const c_rate As Double = 67#
' we use the type of the default to distinguish it from a Date
If VarType(time) <> vbDate Then  ' no argument, or caller error
  time = Now()
Else  ' Date argument was supplied, so we have no fractional seconds
  Time2Long = c_rate * (86400# * (DatePart("y", time) - 1) + _
                        3600# * DatePart("h", time) + _
                        60# * DatePart("n", time) + _
                        DatePart("s", time))
  Exit Property
End If
' time is "Now()" so we can get fractional seconds from Timer
Time2Long = c_rate * (86400# * (DatePart("y", time) - 1) + Timer)
End Property

'========== Property ===========================================================
Property Get Version() As String
Attribute Version.VB_Description = "The date of the latest revision of this class, as a string in the format YYYY-MM-DD such as 2005-03-24."
' The date of the latest revision of this class, as a string in the format
' YYYY-MM-DD such as 2005-03-24.
Version = c_version
End Property

'---------------------------- end of file --------------------------------------

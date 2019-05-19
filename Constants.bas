Attribute VB_Name = "Constants"
Attribute VB_Description = "This Module exports some useful Const values with all-bits-correct accuracy. Devised & coded by John Trenholme."
'          ____                     _                  _.
'         / ___|  ___   _ __   ___ | |_   __ _  _ __  | |_  ___
'        | |     / _ \ | '_ \ / __|| __| / _` || '_ \ | __|/ __|
'        | |___ | (_) || | | |\__ \| |_ | (_| || | | || |_ \__ \
'         \____| \___/ |_| |_||___/ \__| \__,_||_| |_| \__||___/
'
'###############################################################################
'#
'# Visual Basic 6 and VBA Module file "Constants.bas"
'#
'# Initial version 30 Nov 2003 by John Trenholme.
'#
'# This Module exports some useful Const values with all-bits-correct accuracy.
'# They are exported with names that start with a Capital letter, but no special
'# suffix. Therefore, be careful not to use any name here for your own variable.
'#
'# Exports the routines:
'#   Function infinityDbl
'#   Function nanDbl
'#   Function constantsVersion
'#
'###############################################################################

Option Base 0          ' array base value when not specified - the default
Option Compare Binary  ' string comparison based on Asc(char) - the default
Option Explicit        ' forces explicit variable declaration - changes default

Private Const Version_c As String = "2014-06-11"

' VB is remiss in not defining common useful constants, so we do it here.
' The values are as accurate as IEEE 754 64-bit floating point allows (that is
' to say, they are good to the last bit). The names start with Capital Letters
' to indicate that they are Const's.

' The sum form of the constants below is used to keep VB & VBA from truncating
' a digit or two when saved to, and then loaded from a file, since only 15
' digits are kept when that is done. This is a known defect of VB and VBA.

' Note that you can copy these values and paste them into your code as needed,
' if you don't want to include this file. You may not want the "Public" part.

' See the Maple file "Constants.mws" for the calculation of the constants.

' ===== basic mathematical constants

Public Const Catalan As Double = 0.915965594177219
Public Const Cos60 As Double = 0.5
Public Const Euler As Double = 0.57721566 + 4.9015329E-09
Public Const Pi As Double = 3.1415926 + 5.358979324E-08
Public Const PiInv As Double = 0.31830988 + 6.1837907E-09
Public Const PiOvr2x As Double = 1.5707963 + 2.67948965E-08
Public Const PiOvr3 As Double = 1.0471975 + 5.11965978E-08
Public Const Sqr2 As Double = 1.4142135 + 6.237309506E-08
Public Const Sqr3 As Double = 1.7320508 + 7.56887729E-09
Public Const Sqr5 As Double = 2.23606797749979
Public Const SqrPi As Double = 1.7724538 + 5.09055161E-08
Public Const SqrTwoPiInv As Double = 0.39894228 + 4.014327E-10

' derived constant values

Public Const Cos30 As Double = 0.5 * Sqr3
Public Const Cos45 As Double = 0.5 * Sqr2
Public Const DegToRad As Double = Pi / 180#
Public Const Golden As Double = 0.5 * (Sqr5 + 1#)  ' 1.618034  2 - 1.62 = 0.38
Public Const PiOvr2 As Double = 0.5 * Pi
Public Const PiOvr4 As Double = 0.25 * Pi
Public Const PiOvr6 As Double = 0.5 * PiOvr3
Public Const SqrTwoPi As Double = Sqr2 * SqrPi
Public Const SqrHalf As Double = 0.5 * Sqr2
Public Const RadToDeg As Double = 180# / Pi
Public Const Sin30 As Double = Cos60
Public Const Sin45 As Double = 0.5 * Sqr2
Public Const Sin60 As Double = 0.5 * Sqr3
Public Const TwoPi As Double = 2# * Pi

' Base-e and base-10 logs

Public Const Ln_2 As Double = 0.6931471 + 8.055994531E-08
Public Const Ln_5 As Double = 1.6094379 + 1.24341003E-08
Public Const Ln_10 As Double = 2.302585 + 9.29940457E-08
' multiply Log10_e times Log(x) to get Log10(x)
Public Const Log10_e As Double = 0.43429448 + 1.903251828E-09
Public Const Log10_2 As Double = 0.30102999 + 5.663981195E-09
Public Const Log10_5 As Double = 0.69897 + 4.33601883E-09

' ===== IEEE 754 floating-point constants

' largest possible Double
Public Const HugeDbl As Double = 1.79769313486231E+308 + 5.7E+293
' smallest possible normalized Double
Public Const TinyDbl As Double = 2.2250738585072E-308 + 1.48219693752374E-323
' smallest possible un-normalized Double (one-bit mantissa)
Public Const TinyTinyDbl As Double = 4.94065645841247E-324
' this is the smallest value that changes 1.0 when added to it
' you can subtract half this amount and still get a change
Public Const EpsDbl As Double = 2.22044604925031E-16 + 3E-31

' VB uses signed 32-bit values for Long variables
Public Const LongMax As Long = 2147483647
Public Const LongMin As Long = -2147483648#  ' must be a Double

' ===== conversion factors

Public Const CmToM As Double = 0.01        ' centimeters to meters
Public Const M3ToCc As Double = 1000000#   ' cubic meters to cubic centimeters
Public Const MmToM As Double = 0.001       ' millimeters to meters
Public Const ToKilo As Double = 0.001      ' unit to kilo-unit
Public Const ToMega As Double = 0.000001   ' unit to mega-unit
Public Const ToMicro As Double = 1000000#  ' unit to micro-unit
Public Const ToMilli As Double = 1000#     ' unit to milli-unit

' =====  UDT's (not exported)

'-- UDT's used in converting Doubles to-from hex, using Lset to move bits over
Private Type Long2  ' 64 bits seen as two Longs
  lng0 As Long  ' low 32 bits
  lng1 As Long  ' high 32 bits
End Type

Private Type Dbl1  ' 64 bits seen as one Double
  dbl As Double
End Type

'===============================================================================
Public Function constantsVersion() As String
Attribute constantsVersion.VB_Description = "The date of the latest revision to this file, in the format ""YYYY-MM-DD""."
' The date of the latest revision to this file, in the format "YYYY-MM-DD".
constantsVersion = Version_c
End Function

'===============================================================================
Public Function infinityDbl()
Attribute infinityDbl.VB_Description = "Returns positive ""infinity"", which prints as 1.#INF. Be cautious using this in VB6 & VBA code, since strange things may happen. For example, subtraction involving this can cause an overflow error."
' Returns positive "infinity", which prints as 1.#INF. Be cautious using this
' in VB6 & VBA code, since strange things may happen. For example, subtraction
' involving this can cause an overflow error.
Static cache As Double  ' initializes to zero, so use that as not-init flag
If cache = 0# Then cache = hexToDouble("7FF0000000000000")  ' do this only once
infinityDbl = cache
End Function

'===============================================================================
Public Function nanDbl()
Attribute nanDbl.VB_Description = "Returns a positive ""quiet"" NaN (Not-a-Number), which prints as 1.#QNAN. Be cautious using this in VB6 & VBA code, since strange things may happen. For example, comparisons using this can cause an overflow error."
' Returns a positive "quiet" NaN (Not-a-Number), which prints as 1.#QNAN.
' Apparently, this is the only type of NaN that VB6 & VBA recognize.
' Be cautious using this in VB6 & VBA code, since strange things may happen.
' For example, comparisons using this can cause an overflow error (that's why
' we can't use the simpler 'cache' trick here that we used in 'infinityDbl').
Static Init As Boolean, cache As Double
If Not Init Then
  Init = True
  cache = hexToDouble("7FF8000000000000")  ' do this only once
End If
nanDbl = cache
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function hexToDouble(ByVal arg As String) As Double
Attribute hexToDouble.VB_Description = "Internal routine to convert the hex representation of a Double, as a String, into a numeric value."
' Returns a Double with the bit pattern specified in the supplied hex digits.
' First bit is sign, then 11 bits of exponent biased by 3FF, then mantissa.
' Mantissa has hidden leading 1 bit unless number denormalized (exponent 000).
Dim src As Long2, dst As Dbl1
arg = Right$("000000000000000" & arg, 16&)  ' pad left with 0's to length 16
' get argument String into 2 Longs in a UDT, so it can be bit-blasted
src.lng0 = val("&H" & Right$(arg, 8&))
src.lng1 = val("&H" & Left$(arg, 8&))
LSet dst = src  ' cram bits from two Longs over the Double
hexToDouble = dst.dbl
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


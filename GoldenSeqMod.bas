Attribute VB_Name = "GoldenSeqMod"
'
'###############################################################################
'#     _____       _     _             _____            __  __ _
'#    / ____|     | |   | |           / ____|          |  \/  |         | |
'#   | |  __  ___ | | __| | ___ _ __ | (___   ___  __ _| \  / | ___   __| |
'#   | | |_ |/ _ \| |/ _` |/ _ \ '_ \ \___ \ / _ \/ _` | |\/| |/ _ \ / _` |
'#   | |__| | (_) | | (_| |  __/ | | |____) |  __/ (_| | |  | | (_) | (_| |
'#    \_____|\___/|_|\__,_|\___|_| |_|_____/ \___|\__, |_|  |_|\___/ \__,_|
'#                                                   | |
'#                                                   |_|
'# Visual Basic for Applications (VBA) Module file "GoldenSeqMod.bas"
'#
'# Supplies a sequence of distinct Doubles between 0 and 1 for use in selecting
'# plot colors, or other places where values are used to distinguish things.
'# The sequence is deterministic, but jumps around in a way that slowly fills in
'# between previous values. It uses the Golden Mean (or Golden Ratio), which in
'# some sense is the "most irrational" number.
'#
'# Devised and coded by John Trenholme - started 2014-05-13
'#
'###############################################################################

Option Base 0          ' array base value, when not specified - the default
Option Compare Binary  ' string comparison based on Asc(char) - the default
Option Explicit        ' forces explicit variable declaration - changes default

Private Const Version_c As String = "2014-05-15"

Private retVal_m As Double  ' return value from Function

'===============================================================================
Public Function GoldenSeqVersion(Optional ByVal trigger As Variant) As String
' Date of the latest revision to this code, as a string with format "yyyy-mm-dd"
GoldenSeqVersion = Version_c
End Function

'===============================================================================
Public Function goldenSeq() As Double
' return the present value and step to the next; with start at 0, you get:
' 0.000 0.382 0.764 0.146 0.528 0.910 0.292 0.674 0.056 0.438 0.820 0.202 0.584
Const Golden_c As Double = 0.381966011250105  ' (3 - Sqr(5)) / 2 to 15 digits
goldenSeq = retVal_m  ' return the present value
retVal_m = retVal_m + Golden_c  ' step to the next value; might be >= 1
If retVal_m >= 1# Then retVal_m = retVal_m - 1#  ' now 0 <= retVal < 1
End Function

'===============================================================================
Public Sub goldenSeqStart(ByVal startVal As Double)
' set the next return value from "GoldenSeq" and "goldenSeqValue"
retVal_m = startVal - Int(startVal)  ' fails when -1.1E-16 < startVal < 0
If retVal_m >= 1# Then retVal_m = retVal_m - 1#  ' now 0 <= retVal < 1
End Sub

'===============================================================================
Public Function goldenSeqValue() As Double
' return the present value; do not step to the next
goldenSeqValue = retVal_m
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


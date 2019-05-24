Attribute VB_Name = "TeeOutMod"
'         __       __           ______                 __         __.
'    __  / /___   / /   ___    /_  __/____ ___  ___   / /  ___   / /__ _  ___.
'   / /_/ // _ \ / _ \ / _ \    / /  / __// -_)/ _ \ / _ \/ _ \ / //  ' \/ -_)
'   \____/ \___//_//_//_//_/   /_/  /_/   \__//_//_//_//_/\___//_//_/_/_/\__/
'
'###############################################################################
'#
'# Visual Basic for Applications (VBA) Module file "TeeOutMod.bas"
'#
'# Routines for sending output to Immediate window and (optionally) file
'#
'# Devised and coded by John Trenholme - started 2014-06-04
'#
'# Some code sections could be cut and pasted into your code.
'#
'###############################################################################

Option Base 0          ' array base value, when not specified - the default
Option Compare Binary  ' string comparison based on Asc(char) - the default
Option Explicit        ' forces explicit variable declaration - changes default
'Option Private Module  ' No visibility outside this VBA Project

Private Const Version_c As String = "2014-10-04"
Private Const File_c As String = "TeeOutMod.bas[" & Version_c & "]"

' Make this visible everywhere you use teeOut
Global of_g As Integer  ' file unit for output

'########################### Exported Routines #################################

'===============================================================================
Sub teeOutExample()
' Follow this example in your code
Dim pathx As String
pathx = Environ$("UserProfile") & "\Desktop\"  ' put output on desktop
Dim fileName As String
fileName = pathx & "NPSolveUnitTest_" & TimeStamp() & ".txt"
of_g = FreeFile  ' get a free file unit (note: of_g is module global)
Open fileName For Output Access Write Lock Write As #of_g

If Rnd(-1) >= 0! Then Randomize 1  ' replace 1 with any desired seed point
Dim elapse As Single, count As Double
elapse = Timer() - Sqr(2!)

teeOut "========== Unit Tests of NP-Complete Fast Solver =========="
teeOut "File " & File_c & "   Now " & Now()
teeOut

teeOut "-- Tests of hyperdrive controller of the first kind --"
teeOut "parameter", "this code", , "Octave 60-digit result", "relative error"
Dim j As Long, s As String
For j = 0 To 5
  s = Left$(j ^ 1.0345 & Space$(13), 14)  ' force overfill of tab zone
  teeOut j, s, 10# ^ (-10 * j) * (Rnd() - 0.5), 1E-15! * (Rnd() - 0.5!)
Next j
teeOut

teeOut "-- Missing argument tests --"
teeOut , 2, 3
teeOut 1, , 3
teeOut 1, 2, Null
teeOut

elapse = Timer() - elapse + Sqr(2!)  ' remove the "+ Sqr(2!)" in your code
teeOut "~~ task complete ~~ elapsed time " & Round(elapse, 3) & " seconds"

Close #of_g  ' close the file
End Sub

'===============================================================================
Public Sub teeOut(ParamArray arguments() As Variant)
' Prints 0 or more 'arguments' to the Immediate window if in IDE (so always in
' Excel), and also to the output file set up on unit "of_g" if it is open (non-0).
' Each comma-delimited argument is sent to the next available 14-character tab
' zone, the same as Debug.Print's output (but you don't get ";" handling).
' It's best to send only strings or numbers to this Sub, to avoid evil type
' coercion. You can get away with (for example) Currency, Date, Empty (= null
' string), Null (= null string), and so forth, but be careful. Don't send an
' Array, Object, Collection, Dictionary, or non-simple Variant. Missing values
' are OK (even the first argument) and "print" as empty strings. However, VB
' does not allow the final argument in a Sub call to be Missing, so use (e.g.)
' "teeOut 1, 2, 3, null" if you need the final argument to be Missing.
Dim ret As String  ' initialize result as first argument, or null string if none
ret = vbNullString  ' in case there is no argument at all (gives a blank line)
Dim j As Long, k As Long
For j = 0 To UBound(arguments)  ' add on other arguments, tabbing 14 spaces
  If j > 0 Then  ' we need to advance to the next tab zone
    k = Len(ret)
    Const TZW As Long = 14  ' tab zone width
    ret = ret & Space$(TZW * Int(k / TZW + 1) - k)
  End If
  If Not IsMissing(arguments(j)) Then ret = ret & arguments(j)
Next j
Debug.Print ret  ' send to Immediate window (no-op if not in IDE; unlikely)
If 0 <> of_g Then Print #of_g, ret  ' try to print to file, if it has a unit number
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~




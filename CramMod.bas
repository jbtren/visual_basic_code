Attribute VB_Name = "CramMod"
'
'###############################################################################
'#
'# Visual Basic 6 & Visual Basic for Applications Module file "CramMod.bas"
'#
'# Exports a Function that turns a number into a string of specified size.
'# This is useful for printing numbers in minimal space, and for printing
'# numbers into fixed-width fields with maximum accuracy.
'#
'# Devised & coded by John Trenholme - this version started 2008-09-07
'#
'###############################################################################

Option Explicit

Const Version_c As String = "2016-08-26"

' module-global extreme quantities used in unit test reports
Private rMax_m As Double
Private cMax_m As String
Private wMax_m As Long
Private xMax_m As Double
Private rMin_m As Double
Private cMin_m As String
Private wMin_m As Long
Private xMin_m As Double

'===============================================================================
Public Function cram( _
  ByVal number As Double, _
  Optional ByVal size As Long = 6&, _
  Optional ByVal useM As Boolean = True) _
As String
Attribute cram.VB_Description = "Cram a number into the supplied size String (min 4), replacing E- in exponent field with M (unless useM=False). Size = 6 gives at least 1 digit for any Double."
' Turn the input number into a string of the specified size, keeping as much
' numeric precision as is possible. The minimum size you can specify is 4,
' which allows values from -9M9 to -9E9 (i.e., single-digit exponents). The
' default size value is 6. If the size is less than 6, numbers with large
' exponents (2 or 3 digits) may yield strings whose size is larger by 1 (if
' size = 5) or by 2 (if size = 4) than the requested value (1 more than these
' values if "useM" is False). If there are spaces in the result, they are on
' the left (that is, the result is right-justified).
'
' Negative exponents are returned as "M" rather than "E-", to give space for
' one more significant digit. Thus "cram( 1.45E-19, 6&)" yields "1.5M19".
' You can inhibit this behavior by setting the optional argument "useM" to
' False, but you will lose a digit when a negative exponent is needed, and
' you must also use a size of 7 or more to accomodate all possible numbers if
' you don't want some returned strings to be larger than requested.
'
' If you need to convert a string "s" that may have an "M" exponent character in
' it back to a number, you can use the form Val(Replace(s, "M", "E-")).
'
' Note: to get full 15-digit values for any possible exponent, set size >= 21.
'       for 6-digit accuracy, set size = 12
'       for K-digit accuracy, set size = 6 + K
'       add 1 to these values if "useM" is False
'
' Because most of the work is done in VB library formatting routines, this
' function is fast - it takes about 20 microseconds on a 3 GHz Pentium 4.

' enforce some sanity - want to support at least -9E9 and -9M9, so size >= 4
If size < 4& Then size = 4&
' allow space for minus sign to be added later - size is >= 3 after this
Dim sizM As Long
If number < 0# Then sizM = size - 1& Else sizM = size
Dim x As Double
x = Abs(number)  ' work with positive number; add sign on later

' first, see if the default conversion will work (for few-digit numbers)
Dim s As String
s = CStr(x)  ' will have up to 15 digits of precision
If Len(s) <= sizM Then
  ' may contain exponent sign as E+ or E-,; shorten if it does
  s = Replace(s, "E+", "E")  ' so positive exponent values are always signless
  ' we use M for "minus" exponents; also no sign (caller can inhibit this)
  If useM Then s = Replace(s, "E-", "M")
  ' pad on left with spaces so right-justified
  s = Right$(Space$(sizM - 1&) & s, sizM)
Else  ' default was too long, try tweaking
  ' try to fit it in by trimming digits - get min-length scientific format
  Dim digits As Long
  If sizM < 15& Then digits = sizM Else digits = 15&  ' avoid trailing zeros
  s = Format$(x, String$(digits, "0") & "E-0")  ' digits, then E part
  ' find location, and value, of power of 10
  Dim eLoc As Long
  eLoc = InStr(s, "E")  ' must exist
  Dim p10 As Long
  p10 = Val(Mid$(s, eLoc + 1&))  ' extract & convert power of 10
  Dim fmt As String
  If p10 > 0& Then  ' we must add decimal point & positive exponent
    ' start at largest possible accuracy and shrink to fit
    If sizM >= 4& Then fmt = "0." & String$(sizM - 4&, "0") Else fmt = "0"
    Do
      s = Format$(x, fmt & "E-0")
      fmt = Left$(fmt, Len(fmt) - 1&)  ' try fewer digits
    Loop Until (Len(s) <= sizM) Or (Len(fmt) = 0&)
  ElseIf p10 = 0& Then  ' will fit as an integer - no decimal point needed
    s = Left$(s, sizM)
  ElseIf p10 >= -sizM Then  ' needs decimal point but no exponent
    s = Format$(x, String$(sizM + p10, "0") & "." & String$(-p10 - 1&, "0"))
    If Len(s) > sizM Then
      ' maybe a number like 9.999 got rounded up to greater length
      Dim last As String
      last = Right$(s, 1&)
      If (last = "0") Or (last = ".") Then s = Left$(s, sizM)  ' if so, fix
    End If
  Else  ' must add decimal point & negative exponent
    Dim sizP As Long
    If useM Then sizP = 1& Else sizP = 0&
    ' start at largest possible accuracy and shrink to fit
    If sizM >= 4& Then fmt = "0." & String$(sizM - 4&, "0") Else fmt = "0"
    Do
      s = Format$(x, fmt & "E-0")
      fmt = Left$(fmt, Len(fmt) - 1&)  ' try fewer digits
    Loop Until (Len(s) <= sizM + sizP) Or (Len(fmt) = 0&)
    If useM Then
      ' make negative exponent take only one space (caller can inhibit this)
      s = Replace(s, "E-", "M")
      ' fix cases where, for example, 9.95E-10 in width 6 becomes 1.00M9
      If Len(s) < sizM Then s = Replace(s, "M", "0M")
    Else
      ' fix cases where, for example, 9.95E-10 in width 7 becomes 1.00E-9
      If Len(s) < sizM Then s = Replace(s, "E", "0E")
    End If
  End If
End If
If number < 0# Then  ' we need to restore the minus sign
  s = LTrim$(s)  ' might be leading spaces for large size; remove them first
  s = "-" & s
  If Len(s) < size Then s = Space$(size - Len(s)) & s  ' restore spaces
End If
cram = s
End Function

'===============================================================================
Public Function cramVersion() As String
cramVersion = Version_c
End Function

'===============================================================================
Public Sub cramUnitTests()
Attribute cramUnitTests.VB_Description = "Run unit tests"
Attribute cramUnitTests.VB_ProcData.VB_Invoke_Func = "e\n14"
' to run this routine, put the cursor anywhere in it and press F5
Dim tStart As Single: tStart = Timer
Dim fileName As String
fileName = "cramUnitTests_" & Format$(Now(), "yyyy-mm-dd+hh-mm-ss") & ".txt"
Dim ofi As Integer  ' output-file index

Dim pith As String ' "Path" is a reserved word
pith = Environ$("UserProfile") & "\Desktop\" ' path to this user's desktop

ofi = FreeFile
On Error Resume Next
Open pith & fileName For Output Access Write Lock Write As #ofi
If Err.number <> 0& Then
  ofi = 0  ' file did not open - don't use it
  MsgBox "ERROR - unable to open output file:" & vbLf & _
         """" & fileName & """" & vbLf & _
         "Error description: " & Err.Description & vbLf & _
         "Error number: " & Err.number & vbLf & _
         "No output will be written to file. Quitting!", _
         vbOKOnly Or vbExclamation, _
         "File Open ERROR"
  Exit Sub
End If
On Error GoTo 0
Application.StatusBar = "Working..."

Const Seed As Long = 123456789
If Rnd(-1) < 2# Then Randomize Seed

Print #ofi, "********** Unit Tests of Trenholme's 'cram' Function **********"
Print #ofi, "User: " & Environ$("UserName") & _
  "     Domain: " & Environ$("UserDomain")
Dim ts As String
ts = Replace(Environ$("Processor_Identifier"), "ping", vbNullString)
ts = Replace(ts, "GenuineIntel", "Level " & Environ$("Processor_Level") & _
  ", Rev. " & Environ$("Processor_Revision"))
Print #ofi, "Computer: " & Environ$("ComputerName") & "     " & ts
Print #ofi, "CramMod Version tested: " & cramVersion()
Print #ofi, "Started: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
Print #ofi,

rMax_m = -1E+308
rMin_m = -rMax_m

Application.StatusBar = "*** Examples (no testing) ***"
Print #ofi, "*** Examples (no testing) ***"
Dim nTrials As Long
nTrials = -1&
Dim nErrors As Long
Dim j As Long
Dim x As Double, wide As Long, c As String
Const nEx As Long = 100&
For j = 0& To nEx - 1&
  x = 9.99999999999999 * Rnd() * 10# ^ CInt(615# * (Rnd() - 0.5))
  If Rnd() > 0.5 Then x = -x  ' put on random sign
  wide = 6& + Int(16# * j / (nEx - 6&)) ' widths from 6 to 22
  c = cram(x, wide)
  check c, wide, x, nErrors, nTrials, ofi
Next j
Print #ofi,

nTrials = 0&
nErrors = 0&

Application.StatusBar = "*** Tests of simple values ***"
Print #ofi, "*** Tests of simple values:"
For wide = 4& To 23&
  check cram(0#, wide), wide, 0#, nErrors, nTrials, ofi  ' test zero
Next wide
For j = 1& To 997& Step 3&
  x = j
  For wide = 6& To 12&
    check cram(x, wide), wide, x, nErrors, nTrials, ofi
    check cram(-x, wide), wide, -x, nErrors, nTrials, ofi
    check cram(x * 1000#, wide), wide, x * 1000#, nErrors, nTrials, ofi
    check cram(-x * 1000#, wide), wide, -x * 1000#, nErrors, nTrials, ofi
    check cram(x * 9876#, wide), wide, x * 9876#, nErrors, nTrials, ofi
    check cram(-x * 9876#, wide), wide, -x * 9876#, nErrors, nTrials, ofi
  Next wide
Next j
Print #ofi, "Total errors now " & Format$(nErrors, "#,##0") & _
  " out of " & Format$(nTrials, "#,##0") & " trials"
Print #ofi,

Application.StatusBar = "*** Tests of 9.999...En rounding to 10En+1 ***"
Print #ofi, "*** Tests of 9.999...En rounding to 10En+1:"
For j = -308& To 307&
  ' force the round-to-next exponent cases
  x = 9.9999999999 * 10# ^ j
  For wide = 6& To 17&
    c = cram(x, wide)
    check c, wide, x, nErrors, nTrials, ofi  ' positive
    c = cram(-x, wide)
    check c, wide, -x, nErrors, nTrials, ofi  ' negative
  Next wide
Next j
Print #ofi, "Total errors now " & Format$(nErrors, "#,##0") & _
  " out of " & Format$(nTrials, "#,##0") & " trials"
Print #ofi,

Application.StatusBar = "*** Tests of width & accuracy of random values ***"
Print #ofi, "*** Tests of width & accuracy of randomly-generated values:"
Print #ofi, "Pseudo-random seed value: " & Seed
Const nRandom As Long = 1000000
For j = 1& To nRandom
  ' make up random mantissa & exponent
  x = 9.99999999999999 * Rnd() * 10# ^ CInt(615# * (Rnd() - 0.5))
  If Rnd() > 0.5 Then x = -x  ' put on random sign
  wide = Int(6# + 18# * Rnd())  ' random widths from 6 to 23
  c = cram(x, wide)
  check c, wide, x, nErrors, nTrials, ofi
  If nErrors = 100& Then Exit For  ' abort if more than 100 errors
Next j
Print #ofi, "  Tested " & Format$(nRandom, "#,##0") & _
  " random cases - saw " & Format$(nErrors, "#,##0") & " errors"
Print #ofi,

Print #ofi, "Total errors now " & Format$(nErrors, "#,##0") & _
  " out of " & Format$(nTrials, "#,##0") & " trials"
Print #ofi,

Print #ofi, "Max relative error = " & rMax_m
Print #ofi, "  "; xMax_m; " in width "; wMax_m; " became "; cMax_m
Print #ofi, "Min relative error = " & rMin_m
Print #ofi, "  "; xMin_m; " in width "; wMin_m; " became "; cMin_m
Print #ofi,

Print #ofi, "--- Unit tests done at " & _
  Format$(Now(), "yyyy-mm-dd hh:mm:ss") & " --- elapsed time " & _
  Format$(Timer() - tStart, "#,##0.000") & " seconds ---"
Close #ofi
ofi = 0
Application.StatusBar = "File: """ & fileName & """ on desktop"
End Sub

' ==============================================================================
Private Sub check(c As String, wide As Long, x As Double, j As Long, _
  n As Long, ofi As Integer)
Dim k As Long, m As Long, s As String, t As String, u As String, w As Long
s = Left$("x: " & x & Space$(23&), 25&)
w = wide
s = s & Left$(" w: " & w & "  ", 6&)
m = Len(c)
t = Left$("'" & c & "'" & Space$(20&), 25&)
k = 0&
u = Replace(c, "M", "E-")
If m <> w Then
  t = t & " ERROR: w = " & m & " not " & w
  k = 1&
ElseIf Not IsNumeric(u) Then
  t = t & " ERROR: value not numeric"
  k = 1&
Else
  Dim r As Double
  ' get relative error; zero is a special case
  If x <> 0# Then r = Val(u) / x - 1# Else r = Val(u)
  ' track & report the extrema
  If rMin_m > r Then
    rMin_m = r
    cMin_m = c
    wMin_m = w
    xMin_m = x
  ElseIf rMax_m < r Then
    rMax_m = r
    cMax_m = c
    wMax_m = w
    xMax_m = x
  End If
  ' kludge to narrow acceptance window as digit count rises
  ' selected by trial & error to pass all correctly crammed values
  Dim rMul As Double
  If w > 21& Then w = 21&  ' no tight error for really wide fields
  If w <= 7& Then rMul = 1# Else rMul = 0.38 * 10# ^ (w - 7&)
  r = r * rMul
  If (r < -0.3334) Or (r > 0.3334) Then  ' out of bounds (values for width = 6)
    t = t & " ERROR: value is " & Val(c)
    k = 1&
  End If
End If
If (k > 0&) Or (n < 0&) Then Print #ofi, s & " -> " & t
If k > 0& Then j = j + k&
If n >= 0& Then n = n + 1&
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

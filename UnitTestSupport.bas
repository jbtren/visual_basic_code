Attribute VB_Name = "UnitTestSupport"
Attribute VB_Description = "Routines that support VB6 & VBA unit tests with output to Immediate window and to file. Devised and coded by John Trenholme."
'
'_   _        _  _    _____           _    ___                              _.
' | | | _ _  (_)| |_ |_   _| ___  ___| |_ / __| _  _  _ __  _ __  ___  _ _ | |_.
' |_| || ' \ | ||  _|  | |  / -_)(_-<|  _|\__ \| || || '_ \| '_ \/ _ \| '_||  _|
'\___/ |_||_||_| \__|  |_|  \___|/__/ \__||___/ \_,_|| .__/| .__/\___/|_|   \__|
'                                                    |_|   |_|
'
'###############################################################################
'#
'# Visual Basic 6 & VBA code module "UnitTestSupport.bas"
'#
'# Routines that support unit tests with output to Immediate window and to file.
'#
'# Exports the routines:
'#   Function UnitTestSupportVersion
'#   Sub utCheckLimit
'#   Sub utCompareAbs
'#   Sub utCompareEqualString
'#   Sub utCompareLess
'#   Sub utCompareLessEqual
'#   Sub utCompareRel
'#   Sub utErrorCheck
'#   Sub utFileClose
'#   Sub utFileOpen
'#   Function utPackMultiLine
'#   Sub utSummarize
'#   Sub utTeeOut
'#
'# Devised and coded by John Trenholme - initial version 2006-07-24
'#
'###############################################################################

Option Base 0
Option Compare Binary
Option Explicit
Option Private Module  ' Don't allow visibility outside this Excel Project

Private Const Version_c As String = "2007-12-04"
Private Const m_c As String = "UnitTestSupport"  ' module name

' #Const VBA = True         ' set True in Excel (etc.) VBA ; False in VB6
#Const VBA = False         ' set True in Excel (etc.) VBA ; False in VB6

Private Const EOL As String = vbNewLine  ' short form; works on both PC and Mac

Private ofi_m As Integer  ' output file index
Private time_m As Single  ' start & elapsed time

'===============================================================================
Public Function UnitTestSupportVersion() As String
Attribute UnitTestSupportVersion.VB_Description = "The date of the latest revision to this module as a string in the format 'YYYY-MM-DD' such as 2004-06-18. It's a function so Excel etc. can use it."
' The date of the latest revision to this module as a string in the format
' 'YYYY-MM-DD' such as 2004-06-18. It's a function so Excel etc. can use it.
UnitTestSupportVersion = Version_c
End Function

'===============================================================================
Public Sub utCheckLimit( _
  ByRef worst As Double, _
  ByVal limit As Double, _
  ByRef nWarn As Long)
Attribute utCheckLimit.VB_Description = "If 'worst' is greater in absolute magnitude than 'limit', send a warning to utTeeOut and increment 'nWarn'.  Reset 'worst' to zero."
' If 'worst' is greater in absolute magnitude than 'limit', send a warning to
' utTeeOut and increment 'nWarn'. Reset 'worst' to zero.
If Abs(worst) <= limit Then
  utTeeOut "Worst error: " & worst & "  pass (<= " & limit & ")"
Else
  utTeeOut "Worst error: " & worst & EOL & _
    " WARNING! That's too large - should be <= " & Format$(limit, "0.0000E-0")
  nWarn = nWarn + 1&
End If
worst = 0#
End Sub

'===============================================================================
Public Sub utCompareAbs( _
  ByVal what As String, _
  ByVal approx As Double, _
  ByVal exact As Double, _
  ByRef worst As Double)
Attribute utCompareAbs.VB_Description = "Make an absolute comparison of 'approx' to 'exact', update 'worst', and send 'str' and then indented results to 'utTeeOut'."
' Make an absolute comparison of 'approx' to 'exact', update 'worst', and send
' 'what' and then indented results to 'utTeeOut'.
utTeeOut what & " {absolute error}"
Dim absErr As Double
absErr = approx - exact
If Abs(worst) < Abs(absErr) Then worst = absErr
utTeeOut "  approx " & Format(approx, "0.00000000000000E-0") & _
         "  exact " & Format(exact, "0.00000000000000E-0") & _
         "  absErr " & Format(absErr, "0.000E-0")
End Sub

'===============================================================================
Public Sub utCompareEqualString( _
  ByVal what As String, _
  ByVal str1 As String, _
  ByVal str2 As String, _
  ByRef nWarn As Long)
' Compare 'str1' to 'str2', increment 'nWarn' if unequal, and send 'what'
' and then indented results to 'utTeeOut'.
utTeeOut what & " {string equality}"
If str1 = str2 Then
  utTeeOut "  Pass - strings are both equal to:"
  utTeeOut "  """ & str1 & """"
Else
  utTeeOut "  FAIL - strings are unequal:"
  utTeeOut "  #1 = """ & str1 & """"
  utTeeOut "  #2 = """ & str2 & """"
  nWarn = nWarn + 1&
End If
End Sub

'===============================================================================
Public Sub utCompareLess( _
  ByVal what As String, _
  ByVal have As Double, _
  ByVal upperLimit As Double, _
  ByRef nWarn As Long)
' Check 'have' against 'upperLimit', increment 'nWarn' if >=, and send 'what'
' and then indented results to 'utTeeOut'.
Dim pf As String
If have < upperLimit Then
  pf = "  pass"
Else
  pf = "  FAIL!"
  nWarn = nWarn + 1&
End If
utTeeOut what & " {< comparison}"
utTeeOut "  have " & Format(have, "0.00000000000000E-0") & _
         "  upperLimit " & Format(upperLimit, "0.00000000000000E-0") & pf
End Sub

'===============================================================================
Public Sub utCompareLessEqual( _
  ByVal what As String, _
  ByVal have As Double, _
  ByVal upperLimit As Double, _
  ByRef nWarn As Long)
' Check 'have' against 'upperLimit', increment 'nWarn' if >, and send 'what'
' and then indented results to 'utTeeOut'.
Dim pf As String
If have <= upperLimit Then
  pf = "  pass"
Else
  pf = "  FAIL!"
  nWarn = nWarn + 1&
End If
utTeeOut what & " {<= comparison}"
Dim haveStr As String
If have = Int(have) Then
  haveStr = CStr(have)
Else
  haveStr = Format$(have, "0.00000000000000E-0")  ' 15-digit floating-point
End If
Dim limStr As String
If upperLimit = Int(upperLimit) Then
  limStr = CStr(upperLimit)
Else
  limStr = Format$(upperLimit, "0.00000000000000E-0")  ' 15-digit floating-point
End If
utTeeOut "  have " & haveStr & "  upperLimit " & limStr & pf
End Sub

'===============================================================================
Public Sub utCompareRel( _
  ByVal what As String, _
  ByVal approx As Double, _
  ByVal exact As Double, _
  ByRef worst As Double)
Attribute utCompareRel.VB_Description = "Make a relative comparison of 'approx' to 'exact', update 'worst', and send 'str' and then indented results to 'utTeeOut'."
' Make a relative comparison of 'approx' to 'exact', update 'worst', and send
' 'what' and then indented results to 'utTeeOut'.

Dim relErr As Double
If exact <> 0# Then
  relErr = approx / exact - 1#
Else  ' can't do a relative comparison to 0, so fake it
  If approx = 0# Then
    relErr = 0#
  Else
    relErr = 1000#  ' an arbitrary large value
  End If
End If
If Abs(worst) < Abs(relErr) Then worst = relErr
utTeeOut what & " {relative error}"
utTeeOut "  approx " & Format(approx, "0.00000000000000E-0") & _
         "  exact " & Format(exact, "0.00000000000000E-0") & _
         "  relErr " & Format(relErr, "0.000E-0")
End Sub

'===============================================================================
Public Sub utErrorCheck( _
  ByVal who As String, _
  ByVal expect As Long, _
  ByRef nWarn As Long)
Attribute utErrorCheck.VB_Description = "Determine if the error in global object Err is the one that was expected, reporting the results (good or bad) to utTeeOut."
' Determine if the error in global object Err is the one that was expected,
' reporting the results (good or bad) to utTeeOut.
' Usage:
'    nWarn = 0&
'    On Error Resume Next
'    ... other code that tests error behavior via utErrorCheck...
'    x = 709.782712893385  ' will cause overflow
'    Err.Clear  ' clear any prior error, in case of no error here
'    f = besselI0(x)  ' operation that should cause error
'    utErrorCheck "besselI0(" & x & ")", 6&, nWarn  ' did it overflow?
'    ... other code that tests error behavior via utErrorCheck...
'    On Error GoTo 0
' At this point, nWarn has the number of errors that were not as expected

Dim goodOrBad As String
If expect = Err.Number Then
  goodOrBad = " - that is the correct behavior"
Else
  goodOrBad = " - WARNING! Should have caused error " & expect
  nWarn = nWarn + 1&
End If
utTeeOut who & " caused error " & Err.Number & goodOrBad
If Err.Number <> 0& Then
  utTeeOut "----- Error source: " & Err.Source
  utTeeOut "----- Error description: -----"
  ' replace multiple EOL's with single EOL's for compact printout
  utTeeOut utPackMultiLine(Err.Description)
  utTeeOut "----- end error description --"
End If
End Sub

'===============================================================================
Public Sub utFileClose()
Attribute utFileClose.VB_Description = "Close the file with the results in it (see utFileOpen)."
' Close the file with the results in it, and mark it as closed (index = 0).
If ofi_m <> 0 Then
  time_m = Timer() - time_m  ' get elapsed time
  ' fix midnight rollover (once)
  If time_m < -0.004 Then time_m = time_m + 86400#
  If time_m < 0# Then time_m = 0#  ' ignore timing jitter
  Print #ofi_m, ""
  Print #ofi_m, "~~~~~~ end of file ~~~~~~ elapsed time " & time_m & " seconds"
  Close #ofi_m
  ofi_m = 0
End If
End Sub

'===============================================================================
Public Sub utFileOpen(ByVal fileName As String)
Attribute utFileOpen.VB_Description = "Open the results file, with the supplied name, in the directory where the code is located (EXE folder, Project folder, or Workbook folder)."
' Open the results file, with the supplied name, in the directory where the
' code is located (EXE folder, Project folder, or Workbook folder).

' get path to current folder, and prepend to file name
Dim path As String
#If VBA Then
' note: in Excel, you must save a new workbook at least once so path exists
  path = Excel.Workbooks(1).path
  If path = "" Then
    ofi_m = 0  ' don't use file
    MsgBox "Warning! Workbook has no disk location!" & EOL & _
           "Save workbook to disk before proceeding because" & EOL & _
           "Unit-test routine needs a known location to write to." & EOL & _
           "No unit-test output will be written to file.", _
           vbOKOnly Or vbExclamation, m_c
    Exit Sub
  End If
#Else  ' this is VB6
  ' note: this is the project folder if in VB6 IDE; EXE folder if stand-alone
  path = App.path
#End If
' the following is WINDOWS-ONLY
If Right$(path, 1) <> "\" Then path = path & "\"  ' only C:\ etc. have "\"
Dim ofs As String
ofs = path & fileName

' try to open the file
ofi_m = FreeFile
On Error Resume Next
Err.Clear
Open ofs For Output Access Write Lock Read Write As #ofi_m  ' output file
If Err.Number <> 0 Then
  On Error GoTo 0
  ofi_m = 0  ' don't use file
  MsgBox "ERROR - unable to open unit-test result file:" & EOL & EOL & _
    """" & ofs & """" & EOL & EOL & _
    "No unit-test output will be written to file.", _
    vbOKOnly Or vbExclamation, m_c
End If
' save the start time
time_m = Timer()
End Sub

'===============================================================================
Public Function utPackMultiLine( _
  ByRef text As String, _
  Optional ByVal lineEnd As String = EOL) _
As String
Attribute utPackMultiLine.VB_Description = "Replace multiple EOL's with single EOL's for a more compact printout."
' Replace multiple EOL's with single EOL's for a more compact printout.
' Note that we have defined 'Const EOL As String = vbNewLine' above.
Dim short As String
short = text
Dim find As String
find = lineEnd & lineEnd  ' we look for two in a row
' EOL may have 1 character or 2, depending on caller & system; adapt
Dim size1 As Long, size2 As Long
size1 = Len(lineEnd)
size2 = Len(find)
Dim multiLoc As Long
Do
  multiLoc = InStr(short, find)  ' look for two EOL's in a row
  If multiLoc = 0& Then Exit Do  ' quit if not found
  short = Left$(short, multiLoc - 1& + size1) & _
          Mid$(short, multiLoc + size2) ' replace by 1 EOL & repeat search
Loop
utPackMultiLine = short
End Function

'===============================================================================
Public Sub utSummarize( _
  ByVal nWarn As Long)
utTeeOut
If nWarn = 0& Then
  utTeeOut "Unit test SUCCESS - no warnings"
ElseIf nWarn = 1& Then
  utTeeOut "Unit test FAILURE - 1 warning"
Else
  utTeeOut "Unit test FAILURE - " & nWarn & " warnings"
End If
End Sub

'===============================================================================
Public Sub utTeeOut(Optional ByRef str As String = "")
Attribute utTeeOut.VB_Description = "Send 'str' to Immediate window (if in IDE) and to result file (if open)."
' Send 'str' to Immediate window (if in IDE) and to result file (if open).

Debug.Print str  ' works only if in VB[6,A] IDE editor environment
If ofi_m <> 0 Then Print #ofi_m, str
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


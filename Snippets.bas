' File "Snippets.bas" contains useful code snippets in the VBA-VB6 language
' Devised and coded by John Trenholme

Option Base 0          ' array base value when not specified     - default
Option Compare Binary  ' string comparison based on Asc(char)    - default
Option Explicit        ' force explicit declaration of variables - not default
Option Private Module  ' No visibility outside this VBA Project (no VB6 effect)

################################################################################

'********** Constants **********
Private Const Version_c As String = "2008-12-30"  ' date of latest revision
Private Const FN_c As String = """MainModule[" & Version_c & "]."  ' file name
Private Const N_c As String = "MyName"   ' Module, Form or Class name
Private Const EOL As String = vbNewLine  ' shorthand for end-of-line

' This Project Global is set False to halt at errors, True to use handlers
' It should only exist in one *.bas file; preferably the "main" one
' Set it True in the release version, so users get useful information
' Set it False while developing, so you stop right at the error location
Public Const RelVer_c As Boolean = True  ' this is the "Release Version"
' Public Const RelVer_c As Boolean = False  ' this is not the "Release Version"
Public Const BugVer_c As Boolean = Not RelVer_c  ' this is the "Debug Version"

#Const UnitTest_c = True
' #Const UnitTest_c = False

#If UnitTest_c Then  ' tell unit-test routine who it is we work for
  #Const ThisIsVBA_c = True  ' we are using VBA (probably in Excel)
  ' #Const ThisIsVBA_c = False
#End If

Private Const Pi_c As Double = 3.1415926 + 5.358979324E-08
Private Const TwoPi_c As Double = 2# * Pi_c
Private Const HugeDbl_c As Double = 1.79769313486231E+308 + 5.7E+293
Private Const TinyDbl As Double = 2.2250738585072E-308 + 1.48219693752374E-323
Private Const EpsDbl_c As Double = 2.22044604925031E-16 + 3E-31  ' add changes 1
Private Const MaxLong_c As Long = 2147483647  ' 2^31 - 1
' smallest possible Long - due to VB quirk, must be specified as a Double
Private Const MinLong_c As Long = -2147483648#  ' -( 2^31)

################################################################################

'********** Call counter that keeps on ticking **********
Static calls_s As Double  ' number of times this routine has been called
calls_s = calls_s + 1#    ' stops adding at 9,007,199,254,740,992 calls

################################################################################

'********** Timing with midnight-rollover correction **********
' start timing
Dim startDate As Date
startDate = Date
Dim elapsed As Single
elapsed = Timer()

' -- do something --

' stop timing
elapsed = Timer() - elapsed + 86400! * DateDiff("d", startDate, Date)

################################################################################

'********** Timing with microsecond resolution **********
' Windows API access to precision timing (microsecond resolution)
Private Declare Function QueryPerformanceFrequency _
  Lib "kernel32" (f As Currency) As Boolean
Private Declare Function QueryPerformanceCounter _
  Lib "kernel32" (p As Currency) As Boolean

' start precision timing
Dim timerFreq As Currency, timerStart As Currency, timerEnd As Currency
Call QueryPerformanceFrequency(timerFreq)
Call QueryPerformanceCounter(timerStart)

' -- do something --

' stop precision timing - get result in seconds
Call QueryPerformanceCounter(timerEnd)
Dim elapsed As Double
elapsed = CDbl(timerEnd - timerStart) / CDbl(timerFreq)

################################################################################

'********** Running estimate of mean and variance *********
' used in running mean and variance calculation (Welford 1962)
Dim mean As Double, variance As Double, diffMean As Double
mean = 0#
variance = 0#
For j = 1& To N  ' must have N >= 2
  x = Rnd()  ' calculate random variate 'x'
  diffMean = x - mean
  mean = mean + diffMean / j
  variance = variance + diffMean * (x - mean)
Next j
standardDeviation = Sqr(variance / (N - 1&))

################################################################################

dim nSamples as long
nSamples = 100000000
Const N_incr As Long = 200000  ' report count increment
Dim j As Long, k As Long
' set up to do part by part, reporting progress every N_incr cases
j = 1&  ' first index
Do
  k = j + N_incr - 1&  ' last index, for this part
  If k > nSamples Then k = nSamples  ' final part may be shorter than others
  For j = j To k
    '----------- do the calculation for case j ---------------------------------
    '---------------------------------------------------------------------------
  Next j
  '----- report that j cases are complete
  DoEvents
Loop While j <= nSamples

################################################################################

'********** Current directory in VB6 & Excel VBA **********
' Note: alternatively, path = ".\" works; you may have to search for the file
#Const ThisIsVBA_c = True         ' set True in Excel (etc.) VBA ; False in VB6

' get path to current directory, and prepend to file name
Dim ofs As String
Dim pathx As String
#If ThisIsVBA_c Then
' note: in Excel, save workbook at least once so path exists
  pathx = ThisWorkbook.Path
  If pathx = "" Then
    MsgBox "Warning! Workbook has no disk location!" & vbNewLine & _
           "Save workbook to disk before proceeding because" & vbNewLine & _
           "routine needs a known location to write to." & vbNewLine & _
           "No unit test output will be written to file.", _
           vbOKOnly Or vbCritical, _
           c_Mod & " Unit Test"
    Exit Sub
  End If
#Else
  ' note: this is the project folder if in VB6 IDE; EXE folder if stand-alone
  pathx = App.path
#End If
If Right$(pathx, 1) <> "\" Then pathx = pathx & "\"  ' only C:\ etc. have "\"

################################################################################

'********** Get sequenced file name **********
Dim j As Long, fileName As String
j = 0&
Do
  fileName = "MyFile_" & Format$(j, "000") & ".txt"
  j = j + 1&
Loop While Len(Dir$(pathx & fileName)) > 0&   ' True if file already exists

################################################################################

'********** Open output file on desktop **********
Dim fileName As String
fileName = "OutputFile.txt"
Private ofi_m As Integer  ' output-file index

Dim fn As String
fn = Environ$("UserProfile") & "\Desktop\" & fileName

ofi_m = FreeFile
If ofi_m = 0& Then
  MsgBox "ERROR - can't get unit number for output file:" & EOL & EOL & _
         """" & fileName & """" & EOL & EOL & _
         "No output will be written to file.", _
         vbOKOnly Or vbExclamation, _
         "File Open ERROR"
  'Exit Sub  ' bail out if there's no output file
Else
  On Error Resume Next
  Open fn For Output Access Write Lock Read Write As #ofi_m  ' output file
  If Err.Number <> 0& Then
    ofi_m = 0  ' file did not open - don't use it
    MsgBox "ERROR - unable to open output file:" & EOL & EOL & _
           """" & fileName & """" & EOL & EOL & _
           "No output will be written to file.", _
           vbOKOnly Or vbExclamation, _
           "File Open ERROR"
  End If
  On Error GoTo 0
  'Exit Sub  ' bail out if there's no output file
End If
Print #ofi, "Something"
Close #ofi

################################################################################

'********** Return valid file unit number **********
Private Function fu(Optional ByVal reset As Boolean = False)
' Return a unit number available for use by a file. Use with only one file at a
' time. Call with optional argument set True when you close the file.
' Usage when writing an output file:
'   Open "MyFile.txt" For Output As #fu()
'   Print #fu(), thing1
'   Print #fu(), thing2
'   Close #fu(True)
Static fileUnit As Integer
If reset Then
  fu = fileUnit
  fileUnit = 0
Else
  If fileUnit = 0 Then fileUnit = FreeFile
  fu = fileUnit
End If
End Function

################################################################################

'===============================================================================
Public Function inDesign() As Boolean
' Returns True if program is running in IDE (editor) design environment, and
' False if program is running as a standalone EXE. Useful for "hooking" only
' when standalone, or adjusting for the speed difference between compiled and
' interpreted. So in your program you can say: if [Not] inDesign() Then ...
'         John Trenholme - 2009-10-21
Attribute inDesign.VB_Description = "Returns True if program is running interpreted in IDE (editor) design environment, and False if running as a compiled standalone EXE"
inDesign = False
On Error Resume Next  ' set to ignore error in Assert
Debug.Assert 1& \ 0&  ' attempts this illegal feat only in IDE
If 0& <> Err.Number Then
  inDesign = True  ' comment this out to get compiled behavior while in IDE
  Err.Clear  ' do not pass divide-by-zero error back up to caller (yes, it can!)
End If
End Function

################################################################################

'********** Send to Immediate window & file **********
'===============================================================================
Private Sub teeOut(Optional ByRef str As String = "")
' Sends 'str' to Immediate window (if in IDE) and to output file (if open).
Debug.Print str  ' works only if in VBA or VB6 editor environment (IDE)
If ofi_m <> 0 Then Print #ofi_m, str
End Sub

################################################################################

'********** Improve error message and pass on up (top-level routine) **********
' Note: line numbers can be inserted in this file using MZ-Tools
' Get them from the web site http://www.mztools.com, then donate $20

Const ID_c As String = FN_c & "MyRoutine"""  ' name of this file & routine
Static calls_s As Double  ' number of times this routine has been called
calls_s = calls_s + 1#    ' stops adding at 9.007199254740992E15 calls
' clear Err Object, and set to restore state and notify user on errors
If RelVer_c Then On Error GoTo RestoreState

' ---- routine code ----

RestoreState:  '----- state restoration and error handling -----

' release any resources that are being held( open files, allocated arrays, etc.)
Erase bigArray  ' free up memory (no error if not allocated)

' save properties of Err object & Erl; "On Error GoTo 0" will erase them
Dim errNum As Long, errDsc As String, errLin As Long
errNum = Err.Number
errDsc = Err.Description
errLin = Erl  ' unsupported Erl holds line number of error

' turn on recalculation & screen updating; tell user we're Ready
If Not noExcelVbaSwitch_g Then switchToExcel  ' also erases Err & Erl

' handle error if one exists; if no error, fall through to Exit
If errNum <> 0& Then  ' there was an error
  On Error GoTo 0  ' avoid recursion
  ' if you have an error condition in this routine's code, you can use an
  ' Err.Raise statement with a custom message to restore state & then come
  ' here if RelVer_c is True; if you do that, use .Number values that VBA
  ' does not use (say 2000 to 29999), and set them negative
  If InStr(errDsc, "Error in") > 0& Then  ' error came from below here
    ' add traceback text
    errDsc = errDsc & vbLf & "called from " & ID_c & " call " & calls_s
  Else  ' error was in this routine
    errDsc = errDsc & vbLf & "Error in " & ID_c & " call " & calls_s
  End If
  ' if line number is available, add it on
  If errLin <> 0& Then errDsc = errDsc & " line " & errLin
  ' this is the top level routine, so log error to file before halting
  Dim ofun As Integer
  ofun = FreeFile(0)
  If ofun = 0 Then ofun = FreeFile(1)
  If ofun > 0 Then  ' append all of today's errors onto one file
    ChDir localPath  ' routine assumed to be present
    Dim errorLogName As String  ' put today's date into file name
    ' may want to use another prefix here, instead of xlsFileName()
    errorLogName = Left(xlsFileName(), Len(xlsFileName()) - 4&) & _
      "_Errors_" & Format(Now(), "yyyy-mm-dd") & ".txt"
    On Error Resume Next
    Open errorLogName For Append As #ofun
    If Err.Number = 0& Then
      On Error GoTo 0
      Print #ofun, "###### Error Report from Workbook """ & _
        xlsFileName() & """"
      Print #ofun, "File: """ & fileNameOnly(errorLogName) & """"
      Print #ofun, "Folder: """ & CurDir() & """"
      Print #ofun, "It is now " & Format(Now(), "yyyy-mm-dd hh:mm:ss")
      If ThisIsVBA_c Then
        Print #ofun, "Excel version: " & xlVersion()
        Print #ofun, "Operating system: " & opSys()
        Print #ofun, "User name: " & userName()
      End If
      Print #ofun, "Error caught by " & ID_c
      Print #ofun, "Run-time error '" & Abs(errNum) & "':"
      If Left(Error(Abs(errNum)), 19&) <> "Application-defined" Then
        Print #ofun, "VBA error description: " & Error(Abs(errNum))
      End If
      Print #ofun, Replace(errDsc, vbLf, vbNewLine)  ' vbLf bad in files
      Print #ofun, "------ end of report"
      Print #ofun,
      Close #ofun
    Else
      On Error GoTo 0
      errDsc = errDsc & vbLf & "Could not log to error file!"
    End If
  End If
  If errNum < 0& Then  ' it's an error produced by this code; no Help
    errDsc = errDsc & vbLf & _
      "Help is not available for this error - sorry! Click ""End"""
  End If
  ' force an unhandled error so user sees it on screen
  Err.Raise Abs(errNum), ID_c, errDsc  ' strip sign from coded errors
End If
End Sub

################################################################################

'********** Improve error message and pass on up (lower-level routines) ********
' Note: line numbers can be inserted in this file using MZ-Tools
' Get them from the web site http://www.mztools.com, then donate $20

Const ID_c As String = FN_c & "Sub level2"  ' name of this file & routine
Static calls_s As Double  ' number of times this routine has been called
calls_s = calls_s + 1#    ' stops adding at 9.007199254740992E15 calls
' clear Err Object, and set to restore state and notify user on errors
If RelVer_c Then On Error GoTo RestoreState

' ---- routine code ----

RestoreState:
' ----- clean up any stuff left by the routine -----
' save properties of Err object & Erl; "On Error GoTo 0" will erase them
Dim errNum As Long, errDsc As String, errLin As Long
errNum = Err.Number
' handle error if one exists; if no error, fall through to Exit
If errNum <> 0& Then  ' there was an error
  errDsc = Err.Description & vbLf
  errLin = Erl  ' unsupported Erl holds line number of error
  On Error GoTo 0  ' avoid recursion
  If InStr(errDsc, "Problem in") > 0& Then  ' error below here; add traceback
    errDsc = errDsc & "Called from " & ID_c & " call " & calls_s
  Else  ' error was in this routine (could add more info here)
    errDsc = errDsc & "Problem in " & ID_c & " call " & calls_s
  End If
  ' if line number is available, add it on
  If errLin <> 0& Then errDsc = errDsc & ", line " & errLin
  Err.Raise errNum, ID_c, errDsc ' send on up the call chain
  Resume  ' set Next Statement here & hit F8 to go to error line
End If
End Sub

################################################################################

'********** Force the same random sequence from Rnd() **********
If Rnd(-1) >= 0! Then Randomize 1  ' replace 1 with any desired sequence index
If Rnd(-1) >= 0! Then Randomize Timer  ' force a different sequence each time

################################################################################

'********** Use brute-force bisection to find a zero ***********
x = 1#  ' low end of interval
d = 2# - x  ' high end of interval, minus low end
Do While d > 0.000001  ' compare to absolute X error desired (don't use 0)
  d = 0.5 * d
  ' If Cos(x + d) <= 0 Then x = x + d  ' function call - negative to positive
  If Cos(x + d) >= 0 Then x = x + d  ' function call - positive to negative
Loop
x = x + 0.5 * d  ' zero was somewhere between x and x+d

################################################################################

'===== calculate factorials exactly up to 20, then approximately
Dim fact As Double, j As Long, k As Long
fact = 1#
For j = 1 To 30  ' max here is 170
  fact = fact * j
  If j <= 21 Then
    Debug.Print j; Tab(6); Format(fact, "#,#")
  Else
    Debug.Print j; Tab(6); fact
  End If
Next j

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

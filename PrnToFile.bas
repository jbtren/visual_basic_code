Attribute VB_Name = "PrnToFile"
Attribute VB_Description = "Module supporting simple printing to an output text file, with or without a microsecond timestamp. Try ""setPrnFileName getFirstUnusedPrnFile(""Debug_"", ""txt"")"", then (repeatedly) ""prnFile some-string"" or ""prnFileT xxx""."
'
'###############################################################################
'# Visual Basic 6 source file "PrnToFile.bas"
'#
'# Support for simple printing to an output file.
'#
'# Initial version 16 Mar 2004 by John Trenholme.
'###############################################################################

Option Explicit

Private Const c_Version As String = "2004-06-01"

'*******************************************************************************
'*
'* Declarations and module-global quantities
'*
'*******************************************************************************

' precision timing (good to a few microseconds)
Private Declare Function QueryPerformanceFrequency _
  Lib "kernel32" (f As Currency) As Boolean
Private Declare Function QueryPerformanceCounter _
  Lib "kernel32" (p As Currency) As Boolean

Private m_db As Boolean  ' flag to indicate whether output takes place
' if the output file name is "", there will be no output to file
Private m_fn As String   ' output file name with no path information
Private m_fnp As String  ' output file name with path prepended
Private m_ou As Integer  ' output file output unit (0 -> no file)

'*******************************************************************************
'*
'* Routines
'*
'*******************************************************************************

'===============================================================================
Public Function getFirstUnusedPrnFile(base As String, ext As String, _
  Optional ByVal nDigits As Integer = 3)
Attribute getFirstUnusedPrnFile.VB_Description = "Makes up sequenced file names of the form ""base000.ext"", ""base001.ext"", and so on. Returns the first unused one. Set optional argument for other than 3 digits in sequence #. Sequence uses more digits if necessary."
Dim path As String
path = localDirPath()
If nDigits < 1 Then nDigits = 1  ' force at least one sequence digit
Dim j As Integer, testName As String
j = 0
Do
  testName = base & Format$(j, String$(nDigits, "0")) & "." & ext
  j = j + 1
Loop While Len(Dir$(path & testName)) > 0&  ' True if file already exists
getFirstUnusedPrnFile = testName
End Function '------------------------------------------------------------------

'===============================================================================
Public Function getPrnFileName() As String
Attribute getPrnFileName.VB_Description = "Name of file (without path portion) that output will be written to. Null string ("""") indicates no output file (perhaps setPrnFileName not yet called?)."
getPrnFileName = m_fn
End Function

'===============================================================================
Public Function getPrnFileUnit() As Integer
Attribute getPrnFileUnit.VB_Description = "File unit number of output file, for use in ""Print #unit,..."" statements. If zero, there is no output file (perhaps setPrnFileName not yet called?)."
getPrnFileUnit = m_ou
End Function

'===============================================================================
Public Function getPrintingEnabled() As Boolean
Attribute getPrintingEnabled.VB_Description = "True if print-to-file is enabled, False if not. See setPrintingEnabled"
getPrintingEnabled = m_db
End Function

'-------------------------------------------------------------------------------
Private Function localDirPath() As String
Dim path As String
path = App.path  ' only "C:\" etc. have "\", so add it on for others
If Right$(path, 1) <> "\" Then path = path & "\"
localDirPath = path
End Function

'===============================================================================
' Write supplied text to output file, opening it if necessary.
Public Sub prnFile(text As String)
Attribute prnFile.VB_Description = "If printing is enabled, print supplied text string to file. String may contain ""vbNewLine"" to go to a new line."
If m_db Then
  If m_ou = 0 Then testForOpenFile
  If m_ou <> 0 Then Print #m_ou, text
End If
End Sub

'===============================================================================
' Write supplied text to output file, opening it if necessary and adding a
' microsecond timestamp using the seconds() function.
Public Sub prnFileT(text As String)
Attribute prnFileT.VB_Description = "If printing is enabled, print supplied text string to file, followed by microsecond timestamp. String may contain ""vbNewLine"" to go to a new line."
If m_db Then
  If m_ou = 0 Then testForOpenFile
  ' the ChrW$ value below is Unicode Greek mu
  If m_ou <> 0 Then Print #m_ou, text & _
    " " & Format$(1000000# * seconds(), "#,0") & " " & ChrW$(&H3BC) & "s"
End If
End Sub

'===============================================================================
Public Function prnToFileVersion() As String
Attribute prnToFileVersion.VB_Description = "Version of the PrnToFile module as a string in the format ""YYYY-MM-DD"" such as ""2005-05-24""."
prnToFileVersion = c_Version
End Function

'-------------------------------------------------------------------------------
' Return time since first call to this routine, in seconds.
Private Function seconds() As Double
Static s_base As Currency  ' initializes to 0
Static s_freq As Currency  ' initializes to 0
If s_freq = 0@ Then  ' routine not initialized, or unable to read frequency
  QueryPerformanceFrequency s_freq  ' try to read frequency
  ' if frequency is good, try to read base time (else it stays at 0)
  If s_freq <> 0@ Then QueryPerformanceCounter s_base
End If
' if we have a good base time, then we must have a good frequency also
If s_base <> 0@ Then
  Dim time As Currency
  QueryPerformanceCounter time
  seconds = (time - s_base) / s_freq
Else  ' something is wrong - return error value
  seconds = -1#
End If
End Function

'===============================================================================
Public Sub setPrnFileName(newName As String)
Attribute setPrnFileName.VB_Description = "Name of the output file. If """", no printing to file takes place. File normally is in Project directory (if in IDE), or in EXE directory (if compiled), but paths relative to that can be used (""..\\File.txt"")."
If m_fn <> newName Then  ' name is changing; close old file (if any)
  Close m_ou
  m_ou = 0
  m_fn = newName
End If
m_db = (m_fn <> "")  ' turn output on if there is an output file, else off
End Sub

'===============================================================================
Public Sub setPrintingEnabled(ByVal newState As Boolean)
Attribute setPrintingEnabled.VB_Description = "Specify whether printing to file takes place (True) or not (False). Automatically set True when a file name is specified via setPrnFileName."
m_db = newState  ' if file name is "" (null), you still get no output
End Sub

'-------------------------------------------------------------------------------
Private Sub testForOpenFile()
If m_fn <> "" Then  ' auto-open file on first use (End auto-closes)
  m_fnp = localDirPath() & m_fn
  m_ou = FreeFile
  If m_ou = 0 Then
    MsgBox "Unable to get unit number for file" & vbLf & _
           """" & m_fnp & """" & vbLf & vbLf & _
           "No file will be opened & no file output will be produced.", _
           vbOKOnly Or vbExclamation, "PrnToFile.bas: testForOpenFile"
  Else
    Open m_fnp For Output As #m_ou
  End If
End If
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


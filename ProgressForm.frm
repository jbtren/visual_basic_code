VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressForm 
   Caption         =   "Task Progress"
   ClientHeight    =   240
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10200
   OleObjectBlob   =   "ProgressForm.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "ProgressForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'
'###############################################################################
'#  _____                                          ______
'# |  __ \                                        |  ____|
'# | |__) |_ __  ___    __ _  _ __  ___  ___  ___ | |__  ___   _ __  _ __ ___
'# |  ___/| '__|/ _ \  / _` || '__|/ _ \/ __|/ __||  __|/ _ \ | '__|| '_ ` _ \
'# | |    | |  | (_) || (_| || |  |  __/\__ \\__ \| |  | (_) || |   | | | | | |
'# |_|    |_|   \___/  \__, ||_|   \___||___/|___/|_|   \___/ |_|   |_| |_| |_|
'#                      __/ |
'#                     |___/
'# VBA UserForm file "ProgressForm.frm"
'#
'# Supplies a bar so Excel macros can show their progress.
'#
'# Usage:
'#   ProgressForm.Startup "Progress in Task 37"
'#   For jBigLoop = 1& to LoopCount
'#     -- carry out some lengthy task
'#     ProgressForm.Update jBigLoop / LoopCount
'#   Next
'#   Unload ProgressForm
'#
'# Started 2014-05-15 by John Trenholme
'#
'###############################################################################

Option Base 0          ' array base value, when not specified - the default
Option Compare Binary  ' string comparison based on Asc(char) - the default
Option Explicit        ' forces explicit variable declaration - changes default

' class-global Const values (convention: start with upper-case; suffix "_c")
Private Const Version_c As String = "2014-05-19"
Private Const File_c As String = "ProgressForm[" & Version_c & "]."

Private Const PctFmt_c As String = "00.0%"  ' to show percent completion

Private pctCap_m As String  ' the most recent percent string
Private pctWide_m As Long   ' width of the percent-text label
Private dStart_m As Single  ' date we started a new progress display
Private tStart_m As Single  ' time we started a new progress display

'############################ Events ###########################################

'===============================================================================
Private Sub UserForm_Initialize()
' Actions carried out when UserForm is created. This happens the first time
' it is used, or when it is used after End, Reset or the like has erased all
' variables & objects.
With Me  ' set sensible initial values
  pctCap_m = Format$(0#, PctFmt_c)
  pctWide_m = .PercentLBL.Width
End With
End Sub

'===============================================================================
Private Sub UserForm_Terminate()
' Actions carried out when code says "Unload "ProgressForm"
' check if "several" seconds have elapsed; if so, inform the user
Const Several_ As Single = 10!
If Timer() + 86400! * DateDiff("d", dStart_m, Date) > tStart_m + Several_ Then
  VBA.Interaction.Beep  ' default beep sound; must be defined (Control Panel)
  Application.Wait Now() + TimeSerial(0!, 0!, 1!)
  Application.Speech.Speak "ok! task complete", True
End If
End Sub

'############################ Methods ##########################################

'===============================================================================
Public Sub Startup(Optional ByVal topMsg As String = "Task Progress Fraction", _
  Optional ByVal leftRightFrac As Double = 0.5, _
  Optional ByVal topFrac As Double = 0.015)
' Set the initial state (and, optionally, position) of the ProgressForm
' "topMsg" is what appears in the ProgressForm's title bar
' "leftRightFrac" is position of center of bar, with respect to Excel
' "topFrac" is position of top of bar, with respect to Excel
With Me
  .Caption = topMsg
  ' the UserForm position can be wrong when there are multiple
  ' monitors, so set position relative to Excel's window
  .Left = leftRightFrac * (Application.Width - .Width)
  ' put near top, so worksheet is not obscured (even without ribbon)
  .Top = Application.Top + (topFrac * Application.Height)
  pctCap_m = Format$(0#, PctFmt_c)
  .PercentLBL.Caption = pctCap_m
  .Show
End With
dStart_m = Date
tStart_m = Timer()
End Sub

'===============================================================================
Public Sub Update(ByVal fraction As Double)
' Change the ProgressForm to show the new fraction of completion, if it makes a
' difference to the text display of the fraction.
Dim pctCap As String
pctCap = Format$(fraction, PctFmt_c)
If pctCap_m <> pctCap Then  ' fraction changed by enough to make a difference
  pctCap_m = pctCap
  With Me
    .PercentLBL.Caption = pctCap
    .BarFrontLBL.Width = fraction * (.InsideWidth - pctWide_m)
    .Repaint  ' force display of new state
  End With
End If
End Sub

'############################ Properties #######################################

'===============================================================================
Public Property Get Version() As String
' Date of the latest revision to this code, as a string with format "YYYY-MM-DD"
Version = Version_c
End Property

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

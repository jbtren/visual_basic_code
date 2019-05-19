Attribute VB_Name = "Priority"
Attribute VB_Description = "Functions to get and set the priority of the program."
'
'###############################################################################
'#              ____         _               _  __
'#             / __ \ _____ (_)____   _____ (_)/ /_ __  __
'#            / /_/ // ___// // __ \ / ___// // __// / / /
'#           / ____// /   / // /_/ // /   / // /_ / /_/ /
'#          /_/    /_/   /_/ \____//_/   /_/ \__/ \__, /
'#                                               /____/
'# Visual Basic 6 source file "Priority.bas"
'#
'# Get and Set priority of this process.
'#
'# Initial version 4 Jun 2004 by John Trenholme.
'###############################################################################

Option Explicit

Public Const PriorityVersion As String = "2004-06-04"

' Windows API definitions
Public Enum PRIORITY_CLASS
  IDLE_PRIORITY_CLASS = &H40
  NORMAL_PRIORITY_CLASS = &H20
  HIGH_PRIORITY_CLASS = &H80
  REALTIME_PRIORITY_CLASS = &H100  ' use with care, for short periods only!
End Enum

Private Const PROCESS_DUP_HANDLE = &H40

Private Declare Function OpenProcess Lib "kernel32" ( _
  ByVal dwDesiredAccess As Long, _
  ByVal bInheritHandle As Long, _
  ByVal dwProcessId As Long) As Long
    
Private Declare Function CloseHandle Lib "kernel32" ( _
  ByVal hObject As Long) As Long
   
Private Declare Function GetCurrentProcessId Lib "kernel32.dll" () As Long
   
Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long
Private Declare Function GetPriorityClass Lib "kernel32" ( _
  ByVal hProcess As Long) As Long
Private Declare Function SetPriorityClass Lib "kernel32.dll" ( _
  ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long

'===============================================================================
Public Function getPriority() As PRIORITY_CLASS
Attribute getPriority.VB_Description = "Return the present priority of the program."
getPriority = GetPriorityClass(GetCurrentProcess())
End Function

'===============================================================================
Public Sub setPriority(ByVal newPriority As PRIORITY_CLASS)
Attribute setPriority.VB_Description = "Set the program's priority to the new value. Use REALTIME_PRIORITY_CLASS very cautiously, for brief periods."
Static answer1 As Long, answer2 As Long
Dim pid As Long
pid = GetCurrentProcessId()
Dim hProcess As Long
hProcess = OpenProcess(PROCESS_DUP_HANDLE, True, pid)
If (hProcess = 0&) And (answer1 <> vbNo) Then
  answer1 = MsgBox( _
    "Priority: setPriority" & vbLf & _
    "Could not get handle to current process." & vbLf & _
    "Priority not changed." & vbLf & vbLf & _
    "Show this error again?", _
    vbYesNo Or vbExclamation, "setPriority ERROR")
Else
  Dim priorityNow As Long
  priorityNow = GetPriorityClass(hProcess)
  If priorityNow <> newPriority Then
    Dim errCode As Long
    errCode = SetPriorityClass(hProcess, newPriority)
    If (errCode <> 0&) And (answer2 <> vbNo) Then
      answer2 = MsgBox( _
        "Priority: setPriority" & vbLf & _
        "Could not change priority of current process." & vbLf & _
        "Present value: &H" & Hex$(priorityNow) & vbLf & _
        "Desired new value: &H" & Hex$(newPriority) & vbLf & vbLf & _
        "Show this error again?", _
        vbYesNo Or vbExclamation, "setPriority ERROR")
    End If
  End If
End If
CloseHandle hProcess  ' release the handle
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


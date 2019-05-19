Attribute VB_Name = "WriteChartToFileMod"
Attribute VB_Description = "Allows PC Excel user to write any embedded chart or chart sheet to a disk file in a choice of formats (EMF, GIF, JPEG, PDF, PNG, or XPS). Files will be locked until Excel exits. Devised and coded by John Trenholme."
'
'###############################################################################
'# __        __    _ _        ____ _                _  _____     _____ _ _.
'# \ \      / / __(_) |_ ___ / ___| |__   __ _ _ __| ||_   _|__ |  ___(_) | ___
'#  \ \ /\ / / '__| | __/ _ \ |   | '_ \ / _` | '__| __|| |/ _ \| |_  | | |/ _ \
'#   \ V  V /| |  | | ||  __/ |___| | | | (_| | |  | |_ | | (_) |  _| | | |  __/
'#    \_/\_/ |_|  |_|\__\___|\____|_| |_|\__,_|_|   \__||_|\___/|_|   |_|_|\___|
'#
'# Visual Basic for Applications (VBA) Module file "WriteChartToFileMod.bas"
'#
'# Write the selected Excel embedded chart, or chart sheet, in one of several
'# formats, to a sequence-numbered disk file in the directory where the workbook
'# is located. Uses the first unused file name in the sequence "Wbook_00.xxx",
'# "Wbook_01.xxx", ..., where "xxx" is "emf", "gif", "jpeg", "pdf", "png", or
'# "xps". If there are already more than 100 charts, uses Wbook_100.xxx, then
'# Wbook_1000.xxx etc.
'#
'# Here "Wbook" is the workbook's name, without its suffix (no .xls, .xlsm, ...)
'#
'# Note: the files will be locked until Excel exits (can't move, rename, ...).
'#
'# Devised and coded by John Trenholme - Started 17 Aug 2001
'#
'# Exports the routines:
'#   Sub WriteChartToEMFfile  ' WARNING! EMF version replaces clipboard data!
'#   Sub WriteChartToFile
'#   Function WriteChartToFileVersion
'#   Sub WriteChartToGIFfile
'#   Sub WriteChartToJPEGfile
'#   Sub WriteChartToPDFfile
'#   Sub WriteChartToPNGfile
'#   Sub WriteChartToXPSfile
'#   Function WroteChartToFile
'#   Function WroteChartToPath
'#
'###############################################################################

Option Base 0          ' array base value, when not specified - the default
Option Compare Binary  ' string comparison based on Asc(char) - the default
Option Explicit        ' forces explicit variable declaration - changes default
'Option Private Module  ' no effect in VB6; globals project-only in VBA

' Module-global Const values (convention: starts with upper-case; suffix "_c")

Private Const Version_c As String = "2014-04-02"
Private Const File_c As String = "WriteChartToFileMod[" & Version_c & "]"

Public wroteChartToFile_m As String  ' file we wrote to
Public wroteChartToPath_m As String  ' path we wrote the file to

'*******************************************************************************
'
' Win32 API declarations (needed only for EMF file format)
'
'*******************************************************************************

Private Const CF_ENHMETAFILE As Long = 14&

Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function CopyEnhMetaFile Lib "gdi32" Alias "CopyEnhMetaFileA" _
  (ByVal hemfSrc As Long, ByVal lpszFile As String) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) _
  As Long
Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal _
  wFormat As Long) As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) _
  As Long

'############################ Exported routines ################################

'===============================================================================
Public Sub WriteChartToEMFfile()  ' WARNING! EMF version replaces clipboard data
Attribute WriteChartToEMFfile.VB_Description = "Write selected embedded chart or chart sheet to disk as WorkbookName_NN.EMF in the directory the workbook started from. Shortcut: Ctrl-e"
Attribute WriteChartToEMFfile.VB_ProcData.VB_Invoke_Func = "e\n14"
' Send the selected embedded chart or chart sheet to a disk file in EMF format.
' The file name will be WorkbookName_NN.EMF where NN is a sequence number.
' Can be selected and run by the user from the Macro window (Alt-F8)
' Coupled to Ctrl-e so Excel users can do a one-key chart dump-to-file.
commonCode "EMF"
End Sub

'===============================================================================
Public Sub WriteChartToFile(ByRef chartRef As Object, ByVal fileType As String)
Attribute WriteChartToFile.VB_Description = "Write embedded chart or chart sheet (first argument) to disk as WorkbookName_NN.XXX in the directory the workbook started from, where XXX (second argument) is EMF, GIF, JPEG, PDF, PNG, or XPS"
' This allows VBA code to specify an embedded chart, or chart sheet, and send it
' to a file, without disturbing the state of the present Selection. The file
' name will be WorkbookName_NN.XXX where NN is a sequence number and XXX is one
' of the file types given below. Couple this routine to a click-able object on
' a sheet to enable one-click chart-to-file functionality.
' Given the Excel "name" of a Chart (see usage below), send it to a file with
' the file type specified by the supplied 'fileType' argument, which should be
' one of EMF, GIF, JPEG, PDF, PNG, or XPS. If the Chart is on a Worksheet, the
' call to this routine will be of the form:
'   WriteChartToFile Worksheets("Main").ChartObjects("Chart 1"), "EMF"
' or perhaps:
'   WriteChartToFile Sheet1.ChartObjects("Chart 2"), "PNG"
' If the Chart is on its own Chart sheet, the call looks like:
'   WriteChartToFile Charts("Chart 3"), "GIF"
' Note: Excel default Chart names have a space between "Chart" & the number.
Const ID_c As String = File_c & ".WriteChartToFile"

fileType = UCase$(fileType)  ' normalize to UPPER CASE
If "EMF" <> fileType And "GIF" <> fileType And "JPEG" <> fileType And _
   "PDF" <> fileType And "PNG" <> fileType And "XPS" <> fileType Then
  Application.ScreenUpdating = True  ' make sure user can see the MsgBox
  MsgBox _
    "ERROR in " & ID_c & vbLf & vbLf & _
    "Expected 'fileType' to be one of:" & vbLf & _
    "EMF, GIF, JPEG, PDF, PNG, or XPS, but it is: " & fileType & vbLf & _
    "Please specify an allowed chart type and try again." & vbLf & _
    "*** No chart written. ***", _
    vbOKOnly Or vbCritical Or vbMsgBoxSetForeground, _
    "Excel macro: Write Chart to File"
  Exit Sub
End If

Dim screenUpdate As Boolean
screenUpdate = Application.ScreenUpdating  ' save state
Application.ScreenUpdating = False         ' minimize screen flicker
' save the user's active (selected) Cell (and therefore Sheet and WorkBook)
' if no cell is selected, Application.ActiveWindow.ActiveCell returns Nothing
Dim userCell As Range
Set userCell = Application.ActiveWindow.ActiveCell

chartRef.Select      ' Select the supplied object
commonCode fileType  ' run the Selection-to-file code

' put display back at the correct Workbook, Worksheet and Cell
' note: if the Selection was not a Cell, this may "jump back" to some previously
' selected Cell. C'est la vie.
If Not (userCell Is Nothing) Then
  With userCell
    .Worksheet.Parent.Activate  ' the Workbook
    .Worksheet.Activate         ' the Worksheet
    .Activate                   ' the Cell
  End With
  Set userCell = Nothing        ' release object memory (just being careful)
End If
Application.ScreenUpdating = screenUpdate  ' restore state
End Sub

'===============================================================================
Public Function WriteChartToFileVersion( _
  Optional ByVal trigger As Variant) _
As String
Attribute WriteChartToFileVersion.VB_Description = "Date of the latest revision to this code, as a string with format 'yyyy-mm-dd'"
' Date of the latest revision to this code, as a string with format 'yyyy-mm-dd'
' Call from an Excel cell by (e.g.) WriteChartToFileVersion(NOW()) so we don't
' have to use "Application.Volatile True" here
WriteChartToFileVersion = Version_c
End Function

'===============================================================================
Public Sub WriteChartToGIFfile()
Attribute WriteChartToGIFfile.VB_Description = "Write selected embedded chart or chart sheet to disk as WorkbookName_NN.GIF in the directory the workbook started from."
Attribute WriteChartToGIFfile.VB_ProcData.VB_Invoke_Func = " \n14"
' Send the selected embedded chart or chart sheet to a disk file in GIF format.
' The file name will be WorkbookName_NN.GIF where NN is a sequence number.
' Can be selected and run by the user from the Macro window (Alt-F8)
commonCode "GIF"
End Sub

'===============================================================================
Public Sub WriteChartToJPEGfile()
Attribute WriteChartToJPEGfile.VB_Description = "Write selected embedded chart or chart sheet to disk as WorkbookName_NN.JPEG in the directory the workbook started from."
Attribute WriteChartToJPEGfile.VB_ProcData.VB_Invoke_Func = " \n14"
' Send the selected embedded chart or chart sheet to a disk file in JPEG format.
' The file name will be WorkbookName_NN.JPEG where NN is a sequence number.
' Can be selected and run by the user from the Macro window (Alt-F8)
commonCode "JPEG"
End Sub

'===============================================================================
Public Sub WriteChartToPDFfile()
Attribute WriteChartToPDFfile.VB_Description = "Write selected embedded chart or chart sheet to disk as WorkbookName_NN.PDF in the directory the workbook started from."
Attribute WriteChartToPDFfile.VB_ProcData.VB_Invoke_Func = " \n14"
' Send the selected embedded chart or chart sheet to a disk file in PDF format.
' The file name will be WorkbookName_NN.PDF where NN is a sequence number.
' Can be selected and run by the user from the Macro window (Alt-F8)
' note: before use, select chart and use Print Preview to set output options
commonCode "PDF"
End Sub

'===============================================================================
Public Sub WriteChartToPNGfile()
Attribute WriteChartToPNGfile.VB_Description = "Write selected embedded chart or chart sheet to disk as WorkbookName_NN.PNG in the directory the workbook started from."
Attribute WriteChartToPNGfile.VB_ProcData.VB_Invoke_Func = " \n14"
' Send the selected embedded chart or chart sheet to a disk file in PNG format.
' The file name will be WorkbookName_NN.PNG where NN is a sequence number.
' Can be selected and run by the user from the Macro window (Alt-F8)
commonCode "PNG"
End Sub

'===============================================================================
Public Sub WriteChartToXPSfile()
Attribute WriteChartToXPSfile.VB_Description = "Write selected embedded chart or chart sheet to disk as WorkbookName_NN.XPS in the directory the workbook started from."
Attribute WriteChartToXPSfile.VB_ProcData.VB_Invoke_Func = " \n14"
' Send the selected embedded chart or chart sheet to a disk file in XPS format.
' The file name will be WorkbookName_NN.XPS where NN is a sequence number.
' XPS is Microsoft's lame attempt to impose a common file format like PDF.
' Can be selected and run by the user from the Macro window (Alt-F8)
' note: before use, select chart and use Print Preview to set output options
commonCode "XPS"
End Sub

'===============================================================================
Public Function WroteChartToFile(Optional ByVal trigger As Variant) As String
Attribute WroteChartToFile.VB_Description = "Return the name of the most recently written Chart file. If the write failed, return <NoFile>. If no file-write try, or VBA reset, return <Uninitialized>."
' Return the name of the most recently written Chart file. If the write failed,
' return <NoFile>. If no file-write try, or VBA reset, return <Uninitialized>.
If 0& = Len(wroteChartToFile_m) Then wroteChartToFile_m = "<Uninitialized>"
WroteChartToFile = wroteChartToFile_m
End Function

'===============================================================================
Public Function WroteChartToPath(Optional ByVal trigger As Variant) As String
Attribute WroteChartToPath.VB_Description = "Return the path of the most recently written Chart file. If the write failed, return <NoPath>. If no file-write try, or VBA reset, return <Uninitialized>."
' Return the path of the most recently written Chart file. If the write failed,
' return <NoPath>. If no file-write try, or VBA reset, return <Uninitialized>.
If 0& = Len(wroteChartToPath_m) Then wroteChartToPath_m = "<Uninitialized>"
WroteChartToPath = wroteChartToPath_m
End Function

'############################ Private routines #################################

'===============================================================================
Private Sub commonCode(suffix As String)
' Send the selected embedded chart or chart sheet to a disk file in the format
' specified by the argument.
' This common code is used for all the disk file formats.
Const ID_c As String = File_c & ".commonCode"

Dim sfx As String
sfx = UCase$(suffix)  ' change disk-file suffix to UPPER CASE for MsgBox's

Const ChangeStatusBar_c As Boolean = True  ' switch all on
'Const ChangeStatusBar_c As Boolean = False  ' switch all off
If ChangeStatusBar_c Then  ' ----- change the StatusBar -----
  Application.StatusBar = "Writing chart to " & sfx & " file"
End If

wroteChartToFile_m = "<NoFile>"  ' assume failure
wroteChartToPath_m = "<NoPath>"  ' assume failure

' try to find a Chart associated with the present Selection; abort if none
Dim cht As Chart
Set cht = Application.ActiveChart  ' if user selected something, Chart is active
If (cht Is Nothing) Or ("Chart" <> TypeName(cht)) Then
  Application.ScreenUpdating = True  ' make sure user can see the MsgBox
  MsgBox _
    "ERROR in " & ID_c & vbLf & vbLf & _
    "Expected all or part of a chart to be selected, but" & vbLf & _
    "no chart appears to be active." & vbLf & vbLf & _
    "Please select all or part of a chart and try again." & vbLf & _
    "*** No chart written. ***", _
    vbOKOnly Or vbCritical Or vbMsgBoxSetForeground, _
    "Excel macro: Write Chart to " & sfx & " File"
  GoTo Wrapup_L
End If

' put file in directory where active workbook is located (if there is one)
Dim path As String  ' get path to workbook
path = Excel.ActiveWorkbook.path
If 0& = Len(path) Then
  Application.ScreenUpdating = True  ' make sure user can see the MsgBox
  MsgBox _
    "ERROR in " & ID_c & "[" & sfx & "]" & vbLf & _
    "Active workbook is unsaved and has no disk location!" & vbLf & _
    "Save workbook to disk before proceeding so that" & vbLf & _
    "chart can be written to workbook's location." & vbLf & _
    "*** No chart written. ***", _
    vbOKOnly Or vbCritical Or vbMsgBoxSetForeground, _
    "Excel macro: Write Chart to " & sfx & " File"
  GoTo Wrapup_L
End If

' only 'C:\" etc. have trailing path separator - Windows-specific code
If Application.PathSeparator <> Right$(path, 1&) Then
  path = path & Application.PathSeparator
End If
Dim pathTrim As String  ' path without a trailing Application.PathSeparator
pathTrim = Left$(path, Len(path) - 1&)  ' without the trailing backslash

' make a sequence prefix that is equal to the workbook name (without its suffix)
Dim prefix As String
prefix = Excel.ActiveWorkbook.Name
prefix = Left$(prefix, InStrRev(prefix, ".") - 1&) & "_"

' find an unused file name in the sequence prefix_00.xxx, prefix_01.xxx, ...
' if there are already 100 or more charts, this goes to prefix_100.xxx etc.
Dim seqNum As Long
seqNum = -1&  ' number before initial value of trial sequence number
Dim fileName As String, fullName As String
Do
  seqNum = seqNum + 1&
  fileName = prefix & Format$(seqNum, "00") & "." & LCase$(sfx)
  fullName = path & fileName
  ' Len > 0 if file exists; 7 = vbNormal + vbReadOnly + vbHidden + vbSystem
Loop While Len(Dir$(fullName, 7&)) > 0&

' write out the file
If "PDF" = sfx Then
  ' note: before use, select chart and use Print Preview to set output options
  ' arguments: PDForXPS, fileName, quality, addDocProperties, _
  '   ignorePrintArea, firstPageNumber, lastPageNumber, openAfterWrite
  On Error Resume Next
  cht.ExportAsFixedFormat xlTypePDF, fullName, xlQualityStandard, False, _
    True, , , False
  GoTo CheckIt_L  ' handle success or failure, then GoTo Wrapup_L
ElseIf "XPS" = sfx Then
  ' note: before use, select chart and use Print Preview to set output options
  ' arguments: PDForXPS, fileName, quality, addDocProperties, _
  '   ignorePrintArea, firstPageNumber, lastPageNumber, openAfterWrite
  On Error Resume Next
  cht.ExportAsFixedFormat xlTypeXPS, fullName, xlQualityStandard, False, _
    True, , , False
  GoTo CheckIt_L  ' handle success or failure, then GoTo Wrapup_L
ElseIf "EMF" = sfx Then  ' have to get data via clipboard for EMF files
  Dim res32 As Long  ' Win32 result code
  ' empty the clipboard (not usually necessary)
  res32 = OpenClipboard(0&)
  res32 = res32 Or EmptyClipboard()
  res32 = res32 Or CloseClipboard()
  If 0& = res32 Then ' failed to empty the clipboard
    Application.ScreenUpdating = True  ' make sure user can see the MsgBox
    MsgBox _
      "ERROR in " & ID_c & "[EMF]" & vbLf & _
      "Could not empty clipboard!" & vbLf & _
      "Something is horribly wrong." & vbLf & _
      "*** Trying to write in PNG format instead. ***", _
      vbOKOnly Or vbCritical Or vbMsgBoxSetForeground, _
      "Excel macro: Write Chart to " & sfx & " File"
    commonCode "PNG"
    GoTo Wrapup_L
  End If
  ' copy chart to clipboard, replacing previous contents - appear as on screen
  ' it is EVIL to clobber the clipboard this way, but what's a guy to do?
  'cht.CopyPicture xlScreen, xlPicture, xlScreen
  On Error Resume Next
  cht.ChartArea.Copy
  If 0& <> Err.Number Then  ' attempt to copy chart to clipboard failed - punt
    Application.ScreenUpdating = True  ' make sure user can see the MsgBox
    MsgBox _
      "ERROR in " & ID_c & "[EMF]" & vbLf & _
      "Could not copy chart to clipboard!" & vbLf & _
      "Something is horribly wrong." & vbLf & _
      "*** Trying to write in PNG format instead. ***", _
      vbOKOnly Or vbCritical Or vbMsgBoxSetForeground, _
      "Excel macro: Write Chart to " & sfx & " File"
    On Error GoTo 0
    commonCode "PNG"
    GoTo Wrapup_L
  End If
  res32 = IsClipboardFormatAvailable(CF_ENHMETAFILE)
  If 0& = res32 Then ' data not available, even though we just put it there
    Application.ScreenUpdating = True  ' make sure user can see the MsgBox
    MsgBox _
      "ERROR in " & ID_c & "[EMF]" & vbLf & _
      "No EMF data in clipboard!" & vbLf & _
      "Since we just put EMF data there," & vbLf & _
      "something is horribly wrong." & vbLf & _
      "*** Trying to write in PNG format instead. ***", _
      vbOKOnly Or vbCritical Or vbMsgBoxSetForeground, _
      "Excel macro: Write Chart to " & sfx & " File"
    commonCode "PNG"
    GoTo Wrapup_L
  End If
  ' open the clipboard
  Do
    res32 = OpenClipboard(0&)  ' open & lock clipboard - associate w/ this task
    If 0& = res32 Then ' open failed - some other program is using clipboard?
      Dim res As Integer
      Application.ScreenUpdating = True  ' make sure user can see the MsgBox
      res = MsgBox( _
        "ERROR in " & ID_c & "[EMF]" & vbLf & _
        "Could not open clipboard with EMF data!" & vbLf & _
        "Some other program may be using it." & vbLf & _
        "Check for such a program before continuing." & vbLf & _
        "*** No chart written. ***", _
        vbRetryCancel Or vbExclamation Or vbMsgBoxSetForeground, _
        "Excel macro: Write Chart to " & sfx & " File")
      ' note: we get vbCancel from either Cancel button or dialog's close box
      If vbCancel = res Then
        MsgBox _
          "*** Trying to write in PNG format instead of " & sfx & ". ***", _
          vbOKOnly Or vbInformation Or vbMsgBoxSetForeground, _
          "Excel macro: Write Chart to " & sfx & " File"
        commonCode "PNG"
        GoTo Wrapup_L
      End If
    End If
  Loop While 0& = res32  ' keep trying if we couldn't open it & user says Retry
  ' get clipboard data
  Dim handle As Long
  handle = GetClipboardData(CF_ENHMETAFILE)
  If 0& = handle Then ' could not get handle to data
    res32 = CloseClipboard()
    Application.ScreenUpdating = True  ' make sure user can see the MsgBox
    MsgBox _
      "ERROR in " & ID_c & "[EMF]" & vbLf & _
      "Can't get access to EMF data in clipboard!" & vbLf & _
      "Since we just put EMF data there," & vbLf & _
      "something is horribly wrong." & vbLf & _
      "*** Trying to write in PNG format instead. ***", _
      vbOKOnly Or vbCritical Or vbMsgBoxSetForeground, _
      "Excel macro: Write Chart to " & sfx & " File"
    commonCode "PNG"
    GoTo Wrapup_L
  End If
  ' play enhanced metafile into disk file
  res32 = CopyEnhMetaFile(handle, fullName)
  If 0& = res32 Then ' could not write to disk
    res32 = CloseClipboard()  ' no point in checking for error here
    Application.ScreenUpdating = True  ' make sure user can see the MsgBox
    MsgBox _
      "ERROR in " & ID_c & "[EMF]" & vbLf & _
      "Could not write chart EMF data to:" & vbLf & _
      "Path: """ & pathTrim & """" & vbLf & _
      "File: """ & fileName & """" & vbLf & _
      "Something is horribly wrong. Maybe the disk is full?" & vbLf & _
      "*** Trying to write in PNG format instead. ***", _
      vbOKOnly Or vbCritical Or vbMsgBoxSetForeground, _
      "Excel macro: Write Chart to " & sfx & " File"
    commonCode "PNG"
    GoTo Wrapup_L
  End If
  ' close clipboard (if this does not work, there's nothing we can do about it)
  res32 = CloseClipboard()
  wroteChartToPath_m = pathTrim
  wroteChartToFile_m = fileName
Else  ' not PDF or XPS or EMF; use Excel's built-in conversion
  ' do GIF, JPEG & PNG files using Export method of Chart
  On Error Resume Next
  ' arguments: fileName, filterName, interactive
  cht.Export fullName, sfx, True  ' that was simple
  GoTo CheckIt_L  ' handle success or failure, then GoTo Wrapup_L
End If

' single exit point; breakpoint here while debugging
Wrapup_L:  ' jump-target label
On Error GoTo 0  ' in case of lingering "On Error Resume Next"

If ChangeStatusBar_c Then  ' ----- change the StatusBar -----
  If "<" <> Left$(wroteChartToFile_m, 1&) Then
    Application.StatusBar = "Chart File=""" & fileName & """ Path=""" & _
      pathTrim & """"
  End If
End If
Exit Sub

'-------------------------------------------------------------------------------
CheckIt_L:
' Internal sub to handle success or failure, then GoTo Wrapup_L
If 0& = Err.Number Then  ' there was no error
  wroteChartToPath_m = pathTrim
  wroteChartToFile_m = fileName
Else
  Application.ScreenUpdating = True  ' make sure user can see the MsgBox
  MsgBox _
    "ERROR in " & ID_c & vbLf & vbLf & _
    "Tried to write chart to file in " & sfx & " format, but " & _
    "got error " & Err.Number & vbLf & _
    """" & Error$(Err.Number) & """" & vbLf & _
    "Path: """ & pathTrim & """" & vbLf & _
    "File: """ & fileName & """" & vbLf & _
    "*** No chart written. ***", _
    vbOKOnly Or vbCritical Or vbMsgBoxSetForeground, _
    "Excel macro: Write Chart to " & sfx & " File"
End If
GoTo Wrapup_L
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

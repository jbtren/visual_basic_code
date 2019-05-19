Attribute VB_Name = "Support"
Attribute VB_Description = "Support routines for use with Excel VBA code"
'
'###############################################################################
'#
'# VBA Module file "Support.bas"
'#
'# Support routines for use with Excel VBA code.
'# This Module is full of Windows-specific code.
'# Errors are formatted in the 'standard' way that allows adding traceback info.
'#
'# Started 2010-01-19 by John Trenholme
'#
'# This module exports the routines:
'#   Sub checkIfNameExists
'#   Function clearToEnd
'#   Function CNum
'#   Function CNumPos
'#   Function compName
'#   Function getDataBlock
'#   Function getDataCol
'#   Function getDataRow
'#   Function getComment
'#   Function getLastCell
'#   Function getNameRefersTo
'#   Function hms
'#   Sub interpH3P
'#   Function IsDimArray
'#   Function localPath
'#   Function nameToBool
'#   Function nameToDouble
'#   Function nameToLong
'#   Function normalizedName
'#   Function objID
'#   Function opSys
'#   Sub progressBar
'#   Function seconds
'#   Sub selectionToIrfanView
'#   Sub setComment
'#   Sub setNameTo
'#   Function stringToBool
'#   Function stringToDouble
'#   Function stringToLong
'#   Function supportVersion
'#   Sub switchToExcel
'#   Sub switchToVba
'#   Function systemEnv
'#   Sub textToClipboard
'#   Function timeStamp
'#   Function userDocs
'#   Function userName
'#   Function uniqueID
'#   Function vbaSpeed
'#   Function windowsVersion
'#   Function xlsFileName
'#   Function xlVersion
'#
'# This module requires the files:
'#   Formats.bas
'#
'# This module requires reference to:
'#   Microsoft Forms 2.0 Object Library
'#
'###############################################################################

Option Base 0          ' array base value when not specified - the default
Option Compare Binary  ' string comparison based on Asc(char) - the default
Option Explicit        ' forces explicit variable declaration - changes default
'Option Private Module  ' makes globals invisible outside this project

' Module-global Const values (convention: start with upper-case; suffix "_c")

Private Const Version_c As String = "2013-06-12"
Private Const File_c As String = "Support[" & Version_c & "]."  ' Module name

' Module-global quantities (convention: suffix "_m")
' Retained between calls; initialize as 0, "" or False

Private userCell_m As Range
Private vbeWindow_m As Boolean

'############################## Win32 API Declarations #########################

' CSIDL (constant special item ID list) values provide a unique system-
' independent way to identify special folders used frequently by applications,
' but which may not have the same name or location on any given system. For
' example, the system folder may be "C:\Windows" on one system and "C:\Winnt"
' on another. As of Windows Vista, these values have been replaced by
' KNOWNFOLDERID values, so use them only with WinXP (but they seem to work fine
' with Win7).
Private Enum CSIDL_VALUES
  CSIDL_PROGRAM_FILES = &H26
  CSIDL_PERSONAL = &H5  ' "My Documents" or "Documents"
  CSIDL_FLAG_PER_USER_INIT = &H800
  CSIDL_FLAG_NO_ALIAS = &H1000
  CSIDL_FLAG_DONT_VERIFY = &H4000
  CSIDL_FLAG_CREATE = &H8000
  CSIDL_FLAG_MASK = &HFF00
End Enum

Private Const SHGFP_TYPE_CURRENT = &H0 'current value for user, verify it exists
Private Const SHGFP_TYPE_DEFAULT = &H1

Private Const MAX_LENGTH = 260  ' maximum return buffer size
Private Const S_OK = 0     ' API return code
Private Const S_FALSE = 1  ' API return code

Private Declare Function BringWindowToTop Lib "user32" _
  (ByVal hwnd As Long) As Long

Private Declare Function GetComputerName Lib "kernel32" Alias _
  "GetComputerNameA" (ByVal b As String, ByRef n As Long) As Long

Private Declare Function QueryPerformanceFrequency _
  Lib "kernel32" (f As Currency) As Boolean
Private Declare Function QueryPerformanceCounter _
  Lib "kernel32" (p As Currency) As Boolean

Private Declare Function SHGetFolderPath Lib "shfolder.dll" _
  Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, ByVal nFolder As Long, _
  ByVal hToken As Long, ByVal dwReserved As Long, ByVal lpszPath As String) _
  As Long

Private Declare Function strLenW Lib "kernel32" Alias "lstrlenW" _
  (ByVal lpString As Long) As Long

'############################## Exported Routines ##############################

'===============================================================================
Public Sub checkIfNameExists(ByRef theName As String)
Attribute checkIfNameExists.VB_Description = "Raise error 5 if the supplied name is not among Excel's defined names in the active workbook."
' Check to be sure the supplied name exists somewhere in the active workbook.
' Note: workbooks can contain global names, and also names local to specific
' sheets. Sheet-specific names have the form "Sheet1!theName" (i.e., the sheet
' name is prefixed, separated by an exclamation point), whereas workbook global
' names have no prefix. Names can refer to a Range, a constant, or a formula.
Const ID_C As String = File_c & "checkIfNameExists"
Dim dummy As Long
On Error Resume Next
dummy = ActiveWorkbook.Names(theName).index  ' attempt to access a Property
If 0& <> Err.Number Then  ' supplied name does not exist; raise verbose error
  On Error GoTo 0
  Err.Raise 5&, ID_C, _
    "The name """ & theName & """ does not exist in the " & vbLf & _
    "active workbook """ & ActiveWorkbook.Name & """" & vbLf & _
    "Problem in " & ID_C
End If
End Sub

'===============================================================================
Public Sub clearToEnd( _
  ByRef theRange As Range, _
  Optional ByVal direction As XlDirection = xlDown, _
  Optional ByVal includeBlanks As Boolean = False)
Attribute clearToEnd.VB_Description = "Clear contents of cells in specified area on worksheet"
' Clear contents of cells at and below top (left) row of supplied Range down
' (right) to the first empty cell in column (row) one of the Range. The items in
' parentheses apply if the optional argument "direction" is xlRight. If the
' Optional argument "includeBlanks" is True, clear to the last filled cell(s),
' including any blank cells.
' TODO write correct code and run unit tests
Dim rng As Range  ' the top-left cell of the supplied range
Set rng = theRange(1)
' if top-left cell is blank, do nothing
If 0 = Application.WorksheetFunction.CountA(rng) Then Exit Sub
Dim last As Long
If includeBlanks Then
  last = Cells(Rows.Count, rng(1).Column).Resize(rng.Columns, 1&).End(xlUp).Row
Else
  ' note: End(xlDown) gives same result as manual Ctrl+(Down-arrow)
  last = rng.End(xlDown).Row
End If
Dim nClear As Long
nClear = last - rng.Row - 1&
If nClear > 0& Then rng.Resize(nClear, rng.Columns).ClearContents
End Sub

'===============================================================================
Public Function CNum(ByRef v As Variant) As Double
Attribute CNum.VB_Description = "Convert the Variant value of an Excel Cell to a Double with more error checking than CDbl supplies."
' Custom conversion routine to convert the Value of a Cell to a Double with
' more error checking than CDbl supplies. We can do this because Value returns
' a Variant, so we know what we're dealing with. Note that the argument is
' a Variant passed by reference, and may be modified here. Any error is passed
' up for the caller to handle. Note: Error 13 is "Type mismatch"
Const ID_C As String = File_c & "CNum"
On Error GoTo ErrorHandler
If IsEmpty(v) Then v = "<Empty cell>": Err.Raise 13&  ' CDbl(Empty) = 0#
If IsArray(v) Then v = "<Array>": Err.Raise 13&  ' CDbl(Array) = Type mismatch
If vbString = VarType(v) Then
  If 0& = Len(v) Then v = "<Null string>": Err.Raise 13& ' explain to user
End If
CNum = stringToDouble(v)  ' could raise error 13, or 6, or ...
Exit Function  '----------------------------------------------------------------

ErrorHandler:
Dim errDes As String
errDes = Err.Description
'If BugVer_C Then Stop: Resume  ' for debug before Err object clear: 2X F8
' supplement text; did error come from a called routine or this routine?
errDes = errDes & vbLf & _
  "Input cell contains: " & toStr(v) & EOL_ & _
  IIf(0& < InStr(errDes, "Problem in"), "Called by ", "Problem in ") & ID_C & _
  IIf(0& <> Erl, " line " & Erl, vbNullString)
' re-raise error with this routine's ID as Source, and appended to Description
Err.Raise Err.Number, ID_C, errDes
Resume  ' if debugging, set Next Statement here and F8 back to error point
End Function

'===============================================================================
Public Function CNumPos(ByRef v As Variant, _
  Optional ByVal allowZero As Boolean = False) As Double
Attribute CNumPos.VB_Description = "Convert the Variant value of an Excel Cell to a positive (optionally non-negative) Double with more error checking than CDbl supplies."
' Custom conversion routine to convert the Value of a Cell to a Double with
' more error checking than CDbl supplies. We can do this because Value returns
' a Variant, so we know what we're dealing with. Note that the argument is
' a Variant passed by reference, and may be mofified here.
' This version requires the result to be positive or (optionally) non-negative.
' Errors are passed back up to the caller. Note: Error 13 is "Type mismatch"
Const ID_C As String = File_c & "CNumPos"
Dim res As Double
res = CNum(v)  ' could raise error 13, or 6, or ...
Const NoCanDo_c As Long = 17&  ' Can't perform requested operation
If allowZero Then
  If res < 0# Then Err.Raise NoCanDo_c, ID_C, "Can't convert " & toStr(v) & _
    " to a non-negative Double" & vbLf & "Problem in " & ID_C
Else
  If res <= 0# Then Err.Raise NoCanDo_c, ID_C, "Can't convert " & toStr(v) & _
    " to a strictly positive Double" & vbLf & "Problem in " & ID_C
End If
CNumPos = res
End Function

'===============================================================================
Public Function compName(Optional ByVal trigger As Variant) As String
Attribute compName.VB_Description = "Return the NetBIOS name of the Windows computer this program is running on."
' Return the NetBIOS name of the Windows computer this program is running on.
' This needs WMI CORE download on older OS versions such as Win 98 & NT 4.0
Dim os As Object
' "winmgmts:" gives access to Windows Management Instrumentation service
' InstancesOf returns the set of "Win32_OperatingSystem" entries in the registry
For Each os In GetObject("winmgmts:").InstancesOf("Win32_OperatingSystem")
  compName = os.CSName
Next os
Set os = Nothing
End Function

'===============================================================================
Public Function getComment(ByRef theRange As Range) As String
Attribute getComment.VB_Description = "Return the comment in the upper-left cell of the supplied Range; null string if none."
' Return the text from a cell's comment. If no comment, return the null string.
' Note: Excel does not think a cell has changed if only the comment changes
Application.Volatile True  ' so that we always return the present value
Dim topLeft As Range
' get top left Cell, in case 'theRange' is a multi-cell Range
Set topLeft = theRange.item(1&, 1&)
With topLeft
  ' check for an existing comment
  If .Comment Is Nothing Then  ' there is no comment; return null string
    getComment = vbNullString
  Else  ' get comment text; strip out any non-printing characters using Clean
    getComment = Application.WorksheetFunction.Clean( _
      .Comment.Shape.TextFrame.Characters.Text)
  End If
End With
End Function

'===============================================================================
Public Function getDataBlock(ByRef rng As Range) As Variant
' Return a square array of Variants containing the values of all cells starting
' at the top left of "rng" and extending to the last column and row below and
' to the right of "rng" that contain values or formulas. If any cell contains a
' formula, return the value gotten by formula evaluation.  If there is no data
' at all in any cell, return Empty, which can be tested for by using
' IsEmpty(myReturnedArray).
Dim topLeft As Range
Set topLeft = rng.item(1&, 1&)  ' base things on the top left cell of the Range
Dim leftColumn As Variant  ' data region in the leftmost column
leftColumn = getDataCol(topLeft)  ' get entire data region in leftmost column
If IsEmpty(leftColumn) Then
  getDataBlock = Empty  ' Empty return indicates failure
Else
  Dim j As Long, k As Long, m As Long, n As Long, roe As Variant
  m = UBound(leftColumn)
  n = 0&  ' length of the longest row
  For j = 1& To m  ' step down the rows
    roe = getDataRow(rng.item(j, 1&))
    If Not IsEmpty(roe) Then
      k = UBound(roe, 2&)  ' length of this row
      If n < k Then n = k  ' at least one row has k >= 1
    Else
      Stop  ' impossible-condition halt; logic error
    End If
  Next j
  getDataBlock = Range(topLeft, topLeft.item(m, n))  ' sheet data into array
End If
End Function

'===============================================================================
Public Function getDataCol(ByRef rng As Range) As Variant
' TODO use same only-one-item logic as in getDataRow
' Return an array of Variants containing the data in the column starting at the
' top left cell of the supplied Range, and extending to the last non-empty cell
' in that column. If any cell contains a formula, return the value gotten by
' formula evaluation. If there is no data at all in any cell, return Empty,
' which can be tested for by using IsEmpty(myReturnedArray). The array you get
' back will be (somewhat perversely) dimensioned as (1 To N, 1 To 1) where N is
' the number of cells returned. Therefore, scan with myArray(j, 1) with j going
' from 1 to UBound(myArray). Individual empty cells will return Empty, which
' can be tested for by using IsEmpty(myArray(j, 1)).
' Usage:
' Dim myArray as Variant, j As Long
' myArray = getDataCol(Range("Somewhere"))
' If Not IsEmpty(myArray) Then  ' might get Type Mismatch if you don't test
'   For j = 1 to UBound(myArray)
'     ... do something with myArray(j, 1) if not IsEmpty(myArray(j, 1)) ...
'   Next j
' End If
Dim topLeft As Range
Set topLeft = rng.item(1&, 1&)  ' base things on the top left cell of the Range
Dim lastCell As Range  ' look for the last data cell in the column
' start at bottom of column
With topLeft.Parent  ' note: topLeft.Parent is the sheet that "rng" is on
  Set lastCell = .Cells(.Rows.Count, topLeft.Column)
End With
' if last cell has data, return from topLeft to bottom; if not...
If IsEmpty(lastCell) Then  ' search upward to find data, formula, ...
  ' this Find would fail if the last cell had data
  Set lastCell = Range(topLeft, lastCell). _
    Find("*", topLeft, xlFormulas, xlWhole, xlByColumns, xlPrevious)
  If lastCell Is Nothing Then  ' no data anywhere in specified region
    getDataCol = Empty         ' let the caller deal with it
    Exit Function
  End If
End If
getDataCol = Range(topLeft, lastCell).Value  ' transfer Range to array
End Function

'-------------------------------------------------------------------------------
Public Sub test_getDataCol()
' put cursor in this Sub and hit F5
Dim myArray As Variant, j As Long, var As Variant
myArray = getDataCol(Range("Scratch!F6"))  ' point to your test data
If Not IsEmpty(myArray) Then
  For j = 1 To UBound(myArray)
    var = myArray(j, 1)
    If IsEmpty(var) Then Debug.Print j; "= Empty" Else Debug.Print j; "= "; var
  Next j
End If
End Sub

'===============================================================================
Public Function getDataRow(ByRef rng As Range) As Variant
' Return an array of Variants containing the data in the row starting at the
' top left cell of the supplied Range, and extending to the last non-empty cell
' in that row. If any cell contains a formula, return the value resulting from
' formula evaluation. If there is no data at all in any cell, return Empty,
' which can be tested for by using IsEmpty(myReturnedArray). The array you get
' back will be (somewhat perversely) dimensioned as (1 To 1, 1 To N) where N is
' the number of cells returned. Therefore, scan with myArray(1, j) with j going
' from 1 to UBound(myArray, 2). Individual empty cells will return Empty, which
' can be tested for by using IsEmpty(myArray(1, j)).
' Usage:
' Dim myArray as Variant, j As Long
' myArray = getDataRow(Range("Somewhere"))
' If Not IsEmpty(myArray) Then  ' might get Type Mismatch if you don't test
'   For j = 1 to UBound(myArray, 2)
'     ... do something with myArray(1, j) if not IsEmpty(myArray(1, j)) ...
'   Next j
' End If
Dim topLeft As Range
Set topLeft = rng.item(1&, 1&)  ' base things on the top left cell of the Range
Dim lastCell As Range  ' look for the last data cell in the row
' start at end of row (note: topLeft.Parent is the sheet that "rng" is on)
With topLeft.Parent
  Set lastCell = .Cells(topLeft.Row, .Columns.Count)
End With
If IsEmpty(lastCell) Then
  ' we need to search backward to find the last cell that has something
  Set lastCell = Range(topLeft, lastCell). _
    Find("*", topLeft, xlFormulas, xlWhole, xlByRows, xlPrevious)
  If lastCell Is Nothing Then  ' no data anywhere in specified region
    getDataRow = Empty         ' let the caller deal with it
    Exit Function
  End If
  If lastCell.Column = topLeft.Column Then  ' there is only one cell with data
    Dim ret As Variant  ' we need to make up a Variant with the required shape
    ReDim ret(1& To 1&, 1& To 1&)  ' force 2D array, else get scalar
    ret(1&, 1&) = lastCell.Value
    getDataRow = ret
  Else  ' there are 2 or more entries, so we will get an array back
    getDataRow = Range(topLeft, lastCell).Value
  End If
Else  ' last cell has something, so return entire row to end
  getDataRow = Range(topLeft, lastCell).Value
End If
End Function

'-------------------------------------------------------------------------------
Public Sub test_getDataRow()
' put cursor in this Sub and hit F5
Dim myArray As Variant, j As Long, var As Variant
myArray = getDataRow(Range("Scratch!B4"))  ' point to your test data
If Not IsEmpty(myArray) Then
  For j = 1 To UBound(myArray, 2)
    var = myArray(1, j)
    If IsEmpty(var) Then Debug.Print j; "= Empty" Else Debug.Print j; "= "; var
  Next j
  Debug.Print
End If
End Sub

'===============================================================================
Public Function getLastCell(ByRef wks As Worksheet) As Range
' Return the last cell on the worksheet, defined as the cell at the intersection
' of the last row with data and the last column with data. Thus the cell itself
' may be empty, but it defines the end of the square block with data somewhere
' in it. This is more difficult than it needs to be because Excel is lax about
' updating the UsedRange property, that otherwise would do the job in one call.
Dim colMax As Long, rowMax As Long
colMax = wks.UsedRange.Columns.Count  ' update UsedRange.Columns
rowMax = wks.UsedRange.Rows.Count     ' update UsedRange.Rows (prob. un-needed)
' now that UsedRange is correctly updated, we can use LastCell
Set getLastCell = wks.Cells(rowMax, colMax)
End Function

'===============================================================================
Public Function getNameRefersTo( _
  ByVal theName As String, _
  Optional ByVal offsetRow As Long = 0&, _
  Optional ByVal offsetCol As Long = 0&) _
As String
Attribute getNameRefersTo.VB_Description = "Return a String containing the contents (Range, Constant or Formula) of a defined Name in this workbook."
' Return a String containing the contents of a defined Name. The Name may refer
' to a Range, a constant, or a formula. If it refers to a Range, the contents of
' the upper-left Cell of the Range are returned (as text). If the Optional
' offset arguments are supplied, and if the name refers to a Range, the contents
' of the Cell offset by those values from the upper-left Cell is returned.
' This is based on code by Chip Pearson; see his web site.
Const ID_C As String = File_c & "getNameRefersTo"
Dim nam As Name
On Error Resume Next
Set nam = ActiveWorkbook.Names(theName)  ' localize the Name
Dim errNum As Long
errNum = Err.Number
If Err.Number <> 0& Then
  On Error GoTo 0
  Err.Raise errNum, ID_C, _
    Error$(errNum) & " for Name """ & theName & """" & vbLf & _
      "Name is not defined in Workbook """ & ThisWorkbook.Name & """" & vbLf & _
      "Problem in " & ID_C
End If
Dim rng As Range
Set rng = nam.RefersToRange  ' try to return a Range
errNum = Err.Number
On Error GoTo 0
Dim str As String
If 0& = errNum Then  ' the Name refers to a Range
  ' content of upper left Cell of Range, perhaps offset, as text
  On Error Resume Next
  str = CStr(rng.item(1&, 1&).offset(offsetRow, offsetCol).Value)
  errNum = Err.Number
  On Error GoTo 0
  If errNum <> 0& Then
    Err.Raise errNum, ID_C, _
      Error$(errNum) & " for Name """ & theName & """" & vbLf & _
        "Could not get text from the named cell" & vbLf & _
        "Problem in " & ID_C
  End If
Else  ' the Name refers to a constant or a formula or an array
  str = nam.RefersTo  ' gets a Variant (containing a String, in this case)
  If Mid$(str, 2&, 1&) = """" Then  ' second character is a quote
    ' Name contains a quoted text constant, in the form ="text constant"
    str = Mid$(str, 3&, Len(str) - 3&)  ' strip off head (=") and tail (")
  Else
    ' Name contains a constant, in the form =42, or a formula like =COS(0.1)
    ' You may wish to use Application.Evaluate to turn it into a number (etc.)
    str = Mid$(str, 2&)  ' strip off leading equals sign
  End If
End If
getNameRefersTo = str
End Function

'===============================================================================
Public Function hms(Optional ByRef separator As String = ":") As String
Attribute hms.VB_Description = "Return the time of day to within 100 microseconds relative accuracy, as a string with the format 'HH:MM:SS.SSSS'"
' Return the time of day to within 100 microseconds relative accuracy, as a
' string with the format 'HH:MM:SS.SSSS' (absolute accuracy only good in the
' seconds and up, and then only as accurate as the system's clock).
' This function itself takes more than 500 microseconds on a 3 GHz processor,
' so you probably won't get the same value on successive calls. There can be
' a jump in the fractional seconds when the whole seconds change, so don't
' use this for precision increment timing (use the seconds() function instead).
' The time is guaranteed to always increase, unless you change the separator.
' The optional argument changes the separator from ":" to the supplied text.
' TODO sync the fraction with the seconds
Static oldPartA_s As String, secBase_s As Double  ' prior-call values
Dim secs As Double
secs = seconds()  ' use precision timer, good to about a microsecond
' Now() and microsecond timer may not roll over seconds at the same time, so
' force the result to always increase
Dim partA As String
partA = Format$(Now(), "hh" & separator & "mm" & separator & "ss.")
If oldPartA_s <> partA Then
  ' time has changed, so restart milliseconds
  ' this can cause an apparent "time jump up" when seconds change; so be it
  ' can also be triggered if user changes separator within a second; so be it
  oldPartA_s = partA
  secBase_s = secs  ' new microsecond base
End If
secs = secs - secBase_s  ' begins at zero every time Now() changes seconds
secs = (secs - Int(secs)) * 9999.99  ' milliseconds 0 <= ms < 1000
hms = partA & Format$(Int(secs), "0000")
End Function

'===============================================================================
Public Sub interpH3P(ByRef rng As Range, ByRef nPts As Long, ByVal nMax As Long)
Attribute interpH3P.VB_Description = "Insert extra X,Y points into a specified list, interpolating them by cubic Hermite interpolation"
' Given a Range 'rng' starting a total of 'nPts' X, Y pairs in adjacent
' columns extending downward and to the right from rng.Item(1, 1), add an
' equal number of points between each pair of existing points, using cubic
' Hermite interpolation with the slopes found from parabolic interpolation of
' three points around each interval-end point. Add enough more points so that
' there will be as close to 'nMax' total points as possible, without exceeding
' that value. The input X values should be in strictly increasing or decreasing
' order, or very strange things will happen. Note that 'rng' does not have to
' include all the points, it can just indicate the top left X value.
' Upon return, 'nPts' may be updated, and the range of cells holding the old
' X, Y pairs may be extended. There is no check to see if the result will
' overwrite other sheet data, so be sure there are at least 'nMax' rows
' available for the results. Results are better if the input points are more
' or less equally spaced, unless there is large variation in Y in just a small
' part of the X domain. Because the input data may be overwritten, save it in
' a safe spot if you want to keep it for later use.
Const ID_C As String = File_c & "interpH3P"
Dim topLeft As Range
Set topLeft = rng(1&, 1&)  ' localize the (top left of the) input Range

If nPts <= 0& Then  ' caller wants us to count the data rows
  nPts = topLeft.End(xlDown).Row - topLeft.Row + 1&
End If

If nPts < 3& Then Exit Sub  ' can't do anything useful with only 1 or 2 points

Dim nAdd As Long  ' number of added points between original points
nAdd = Int((nMax - nPts) / (nPts - 1&))  ' note: Int = floor
If nAdd <= 0& Then Exit Sub  ' not enough room for even 1 extra point

' informative alias values for X and Y indices
Const JX As Long = 1&
Const JY As Long = JX + 1&

' read in the X and Y vectors from the specified Range
Dim xyVals() As Variant  ' must be dynamically allocated to eat a Range
' after this assign, xyVals(j, 1) = X[j], xyVals(j, 2) = Y[j], j = 1 .. nPts
' despite the fact that Option Base 0 is in effect
xyVals = topLeft.Resize(nPts, JY).Value

' do a sanity check on the X,Y values
Const NoCanDo_c As Long = 17&  ' Can't perform requested operation
Dim dxSgn As Double, j As Long, oldX As Double, x As Double
For j = 1& To nPts
  ' note: IsNumeric perversely returns True for an Empty Variant
  If (Not IsNumeric(xyVals(j, JX))) Or IsEmpty(xyVals(j, JX)) Then
    Err.Raise NoCanDo_c, ID_C, _
      "Can't convert X(" & j & ") = """ & toStr(xyVals(j, JX)) & _
      """ to a numeric value" & vbLf & _
      "Problem in " & ID_C
  ElseIf (Not IsNumeric(xyVals(j, JY))) Or IsEmpty(xyVals(j, JY)) Then
    Err.Raise NoCanDo_c, ID_C, _
      "Can't convert Y(" & j & ") = """ & toStr(xyVals(j, JY)) & _
      """ to a numeric value" & vbLf & _
      "Problem in " & ID_C
  End If
  x = CDbl(xyVals(j, JX))
  If j > 1& Then
    If x = oldX Then
      Err.Raise 5&, ID_C, _
        "X(" & j & ") = " & x & " duplicates previous X" & vbLf & _
        "Problem in " & ID_C
    End If
    If dxSgn <> Sgn(x - oldX) Then
      Err.Raise 5&, ID_C, _
        "X(" & j & ") = " & x & " reverses direction from initial dX" & vbLf & _
        "X values need to steadily increase, or steadily decrease" & vbLf & _
        "Problem in " & ID_C
    End If
  Else  ' first point; j = 1
    dxSgn = Sgn(CDbl(xyVals(j + 1&, JX)) - x)  ' is X going up or down?
  End If
  oldX = x
Next j

Dim nOut As Long  ' number of output points
nOut = nPts + nAdd * (nPts - 1&)  ' original points, plus nAdd in each gap

' determine approximate slopes at X values from parabola through 3 points
' TODO apply the coarse-mesh slope constraints from Hyman's 1983 paper
Dim slope() As Double
ReDim slope(1& To nPts)
' move three points (called A, B, and C) along input array
' xB is center X etc., xBA = xB - xA etc., yB is center Y, yBA = yB - yA
Dim xB As Double, xC As Double, xBA As Double, xCB As Double
xB = xyVals(2&, JX)
xC = xyVals(3&, JX)
xBA = xB - xyVals(1&, JX)
xCB = xC - xB
Dim yB As Double, yC As Double, yBA As Double, yCB As Double
yB = xyVals(2&, JY)
yC = xyVals(3&, JY)
yBA = yB - xyVals(1&, JY)
yCB = yC - yB
' first slope is from start of parabola through points 1, 2, 3, at xA
' this parabola will be re-used, with its center slope going to point 2
slope(1&) = (yBA * (xCB + xBA + xBA) / xBA - yCB * xBA / xCB) / (xCB + xBA)
For j = 4& To nPts + 1&
  ' inside slopes are from center of parabola, at xB
  ' note this simplifies to (yC - yA) / (xC - xA) for equal spacing
  slope(j - 2&) = (yBA * xCB / xBA + yCB * xBA / xCB) / (xCB + xBA)
  If j > nPts Then Exit For  ' don't change final parabola
  ' slide over to next set of three points, minimizing array accesses
  xB = xC  ' first do the X values
  xC = xyVals(j, JX)
  xBA = xCB
  xCB = xC - xB
  yB = yC  ' then do the Y values
  yC = xyVals(j, JY)
  yBA = yCB
  yCB = yC - yB
Next j
' last slope is from end of parabola through points N-2, N-1, N, at xC
' we are re-using the parabola whose center slope went to point N-1
slope(nPts) = (yCB * (xCB + xCB + xBA) / xCB - yBA * xCB / xBA) / (xCB + xBA)

' make up Hermite basis functions at interpolation points; we pre-calculate
' because each interval between points uses the same basis-function values
Dim herm() As Double
ReDim herm(1& To nAdd, 1& To 4&)
Dim temp As Double, u As Double, u2 As Double, u3 As Double
For j = 1& To nAdd
  u = j / (nAdd + 1&)  ' fraction of distance from x1 to x2; with 0 < u < 1
  u2 = u * u
  u3 = u2 * u
  temp = 3# * u2 - 2# * u3
  herm(j, 1&) = 1# - temp ' h1(x1)=1 h1(x2)=0 s1(x1)=0 s1(x2)=0 (where s=slope)
  herm(j, 2&) = temp  ' h2(x1)=0 h2(x2)=1 s2(x1)=0 s2(x2)=0
  herm(j, 3&) = u - 2# * u2 + u3  ' h3(x1)=0 h3(x2)=0 s3(x1)=1 s3(x2)=0
  herm(j, 4&) = u3 - u2  ' h4(x1)=0 h4(x2)=0 s4(x1)=0 s4(x2)=1
Next j

' make up the output array
Dim res() As Double
ReDim res(1& To nOut, JX To JY)
Dim x1 As Double, x2 As Double
x2 = CDbl(xyVals(1&, JX))
Dim y1 As Double, y2 As Double
y2 = CDbl(xyVals(1&, JY))
Dim s1 As Double, s2 As Double  ' slope values
s2 = slope(1&)
Dim dx As Double
Dim i As Long, k As Long
k = LBound(res, 1&)  ' index of slot that result goes in
For j = 2& To nPts + 1&
  ' insert original point at bottom of this section (or top of last section)
  x1 = x2
  res(k, JX) = x1
  y1 = y2
  res(k, JY) = y1
  If j > nPts Then Exit For  ' that was the end point; we are done
  x2 = CDbl(xyVals(j, JX))
  y2 = CDbl(xyVals(j, JY))
  s1 = s2
  s2 = slope(j)
  dx = x2 - x1  ' basis spline slopes are dh/du = 1; convert to dh/dx = 1
  For i = 1& To nAdd
    u = i / (nAdd + 1&)  ' fraction of distance from x1 to x2; with 0 < u < 1
    x = x1 + u * (x2 - x1)
    k = k + 1&
    res(k, JX) = x
    res(k, JY) = y1 * herm(i, 1&) + y2 * herm(i, 2&) + _
      (s1 * herm(i, 3&) + s2 * herm(i, 4&)) * dx
  Next i
  k = k + 1&
Next j

' reset the input point count
nPts = nOut
' put the results onto the worksheet
topLeft.Resize(nOut, 2&).Value = res
' dump the array memory (supposed to happen automatically; sometimes doesn't)
Erase herm, res, slope
End Sub

'===============================================================================
Public Function IsDimArray( _
  ByRef arg As Variant, _
  Optional ByVal numIndices As Long = 1&) _
As Boolean
Attribute IsDimArray.VB_Description = "Return True if the supplied item is a properly Dim'd N-dimensional array (default N=1)."
' Return True if the supplied item is a properly Dim'd 1-dimensional array.
' If the Optional argument is present, check for that many index slots.
' Example:
'   Dim a(1 To 2, 3 To 4) as Double
'   IsDimArray(a) = IsDimArray(a, 1) -> False
'   IsDimArray(a, 2) -> True
'   IsDimArray(a, 1) -> False
Dim res As Boolean
If 0& = (VarType(arg) And vbArray) Then
  res = False  ' argument is not an array
Else  ' test for numIndices index slots, but no more
  Dim dummy As Long
  On Error Resume Next
  dummy = LBound(arg, numIndices)  ' fails if given index slot not Dim'd
  res = (0& = Err.Number)
  Err.Clear
  dummy = LBound(arg, numIndices + 1&)  ' fails if next index slot not Dim'd
  res = res And (0& <> Err.Number)
  Err.Clear  ' don't pass Err properties back up to caller (yes!)
End If
IsDimArray = res
End Function

'===============================================================================
Public Function localPath(Optional ByVal trigger As Variant) As String
Attribute localPath.VB_Description = "Return the path to the directory the active Workbook is in."
' Return the path to the directory the active Workbook is in.
' If the Workbook is new and has not been saved, there is no path; return 'C:\'
' This is the same as the INFO("directory") worksheet function
Dim path As String  ' get path to Active Workbook
path = ActiveWorkbook.path
If 0& < Len(path) Then
  ' Windows-specific code - only "C:\" etc. have trailing separator
  If Application.PathSeparator <> Right(path, 1&) Then
    path = path & Application.PathSeparator
  End If
  localPath = path
Else
  localPath = "C:\"  ' Windows-specific code - supply root as default
End If
End Function

'===============================================================================
Public Function nameToBoolean( _
  ByVal theName As String, _
  Optional ByVal offsetRow As Long = 0&, _
  Optional ByVal offsetCol As Long = 0&) _
As Boolean
Attribute nameToBoolean.VB_Description = "Get the text associated with a defined Name & convert it to a Boolean, or raise an error"
' Get the text associated with a defined Name & convert it to a Boolean, or die.
' If the Optional arguments are supplied, and if the supplied Name refers to a
' Range, offset from the top left of the Range by the supplied offsets.
Const ID_C As String = File_c & "nameToBool"
On Error GoTo ErrorHandler

Dim theText As String
' might raise "no such name" error
theText = getNameRefersTo(theName, offsetRow, offsetCol)
nameToBoolean = stringToBoolean(theText)  ' can raise "can't convert that" error
Exit Function  '----------------------------------------------------------------

ErrorHandler:
Dim errDes As String
errDes = Err.Description
' supplement text; did error come from a called routine or this routine?
errDes = errDes & vbLf & _
  IIf(0& < InStr(errDes, "Problem in"), "Called by ", "Problem in ") & ID_C & _
  IIf(0& <> Erl, " line " & Erl, vbNullString)
' re-raise error with this routine's ID as Source, and appended to Description
Err.Raise Err.Number, ID_C, errDes
Resume  ' if debugging, set Next Statement here and F8 back to error point
End Function

'===============================================================================
Public Function nameToDouble( _
  ByVal theName As String, _
  Optional ByVal offsetRow As Long = 0&, _
  Optional ByVal offsetCol As Long = 0&) _
As Double
Attribute nameToDouble.VB_Description = "Get the text associated with a defined Name & convert it to a Double, or raise an error"
' Get the text associated with a defined Name & convert it to a Double, or die.
' If the Optional arguments are supplied, and if the supplied Name refers to a
' Range, offset from the top left of the Range by the supplied offsets.
Const ID_C As String = File_c & "nameToDouble"
On Error GoTo ErrorHandler

Dim Text As String
' might raise "no such name" error
Text = getNameRefersTo(theName, offsetRow, offsetCol)
nameToDouble = stringToDouble(Text)  ' might raise "can't convert that" error
Exit Function  '----------------------------------------------------------------

ErrorHandler:
Dim errDes As String
errDes = Err.Description
' supplement text; did error come from a called routine or this routine?
errDes = errDes & vbLf & _
  IIf(0& < InStr(errDes, "Problem in"), "Called by ", "Problem in ") & ID_C & _
  IIf(0& <> Erl, " line " & Erl, vbNullString)
' re-raise error with this routine's ID as Source, and appended to Description
Err.Raise Err.Number, ID_C, errDes
Resume  ' if debugging, set Next Statement here and F8 back to error point
End Function

'===============================================================================
Public Function nameToLong( _
  ByVal theName As String, _
  Optional ByVal offsetRow As Long = 0&, _
  Optional ByVal offsetCol As Long = 0&) _
As Long
Attribute nameToLong.VB_Description = "Get the text associated with a defined Name & convert it to a Long, or raise an error"
' Get the text associated with a defined Name & convert it to a Long, or die.
' If the Optional arguments are supplied, and if the supplied Name refers to a
' Range, offset from the top left of the Range by the supplied offsets.
Const ID_C As String = File_c & "nameToLong"
On Error GoTo ErrorHandler

Dim Text As String
' might raise "no such name" error
Text = getNameRefersTo(theName, offsetRow, offsetCol)
nameToLong = stringToLong(Text)  ' might raise "can't convert that" error
Exit Function  '----------------------------------------------------------------

ErrorHandler:
Dim errDes As String
errDes = Err.Description
' supplement text; did error come from a called routine or this routine?
errDes = errDes & vbLf & _
  IIf(0& < InStr(errDes, "Problem in"), "Called by ", "Problem in ") & ID_C & _
  IIf(0& <> Erl, " line " & Erl, vbNullString)
' re-raise error with this routine's ID as Source, and appended to Description
Err.Raise Err.Number, ID_C, errDes
Resume  ' if debugging, set Next Statement here and F8 back to error point
End Function

'===============================================================================
Public Function normalizedName(ByVal theName As String) As String
Attribute normalizedName.VB_Description = "Convert the supplied string to 'normalized' form, switched to lower case with all spaces and dashes removed"
' Convert the supplied string to "normalized" form, switched to lower case with
' all spaces and dashes ("-") removed. E.g., "Lg-750 " -> "lg750". This is
' provided to allow name comparison without heartache.
normalizedName = _
  Replace$(Replace$(LCase$(theName), " ", vbNullString), "-", vbNullString)
End Function

'===============================================================================
Public Function objID(ByRef arg As Variant) As String
Attribute objID.VB_Description = "Return the TypeName of the argument, followed by '@' and its address. If it's not an Object, say so"
' Return the TypeName of the argument, followed by "@" and its address
' If it's not an Object, say so
If IsObject(arg) Then
  If 0& <> ObjPtr(arg) Then
    objID = TypeName(arg) & "@" & Right("0000000" & Hex(ObjPtr(arg)), 8&)
  Else
    objID = "Nothing"  ' null pointer is called Nothing by VB
  End If
Else
  objID = "Not an Object, but a " & TypeName(arg)
End If
End Function

'===============================================================================
Public Function opSys(Optional ByVal trigger As Variant) As String
Attribute opSys.VB_Description = "Return the operating system the machine is running & Office 32 or 64 bit info"
' Return the operating system the machine is running & Office 32 or 64 bit info.
' See 'windowsVersion' for an alternative.
Dim ret As String
ret = Application.OperatingSystem
' this is the same as the result of the INFO("osversion") worksheet function
' you get back "Windows (BB-bit) VER" where BB=32 unless host is 64-bit office
' running under 64-bit Windows, when it is 64. The VER string depends on the OS:
' Windows 95: VER = "4.00"
' Windows 98: VER = "4.10"
' Windows Me: VER = "4.90"
' Windows 2000: VER = "NT 5.00"
' Windows XP: VER = "NT 5.01"
' Windows XP 64-bit: VER = "NT 5.02"
' Windows Vista: VER = "NT 6.00"
' Windows 7: VER = "NT 6.01"
' Windows 8: VER = "NT 6.02"
Dim ver As String
ver = Mid$(ret, 18&)
' make a version the user has actually heard of; no change if unknown to us
If "4.00" = ver Then ret = Left$(ret, 17&) & "95"
If "4.10" = ver Then ret = Left$(ret, 17&) & "98"
If "4.90" = ver Then ret = Left$(ret, 17&) & "Me"
If "NT 5.00" = ver Then ret = Left$(ret, 17&) & "2000"
If "NT 5.01" = ver Then ret = Left$(ret, 17&) & "XP"
If "NT 5.02" = ver Then ret = Left$(ret, 17&) & "XP 64-bit"
If "NT 6.00" = ver Then ret = Left$(ret, 17&) & "Vista"
If "NT 6.01" = ver Then ret = Left$(ret, 17&) & "7"
If "NT 6.02" = ver Then ret = Left$(ret, 17&) & "8"
opSys = ret
End Function

'===============================================================================
Public Sub progressBar(Optional ByVal fraction As Double = 0#)
Attribute progressBar.VB_Description = "Show a progress bar in the Excel status-bar area using text characters"
' Show a progress bar in the Excel status-bar area using text characters
Static oldBlocks As Long
If fraction < 0# Then fraction = 0# Else If fraction > 1# Then fraction = 1#
Dim blocks As Long
Const BlockMax_c As Long = 60&  ' maximum number of blocks in bar
blocks = BlockMax_c * fraction  ' round to nearest integer
If oldBlocks <> blocks Then  ' appearance has changed, so update it
  oldBlocks = blocks
  Dim s As String
  ' default StatusBar font is Segoe UI - use Unicode block characters
  ' ChrW(9632) is a filled square block, ChrW(9633) is an unfilled square
  ' the multiplier in the second String$() call keeps the length ~constant
  ' note: this may look bad if user has changed Menu font in windows settings
  s = Format$(fraction, "00% ") & String$(blocks, ChrW(9632)) & _
    String$(1.5 * (BlockMax_c - blocks), ChrW(9633))
  Application.StatusBar = s
End If
End Sub
'===============================================================================
Public Function seconds(Optional ByVal resetBase As Boolean = False) As Double
Attribute seconds.VB_Description = "Return number of seconds since first call to this routine. Good to a few microseconds. A return of -86400 indicates an error."
' Return elapsed wall-clock time since first call to this routine, in seconds.
' If the optional argument is True, reset the base time to "now"
' The granularity will be around a microsecond on a standard PC.
' Usage:
' Timing a single action (for multiples, use "elapsedTime1", elapsedTime2", ...)
'   Dim elapsedTime as Double
'   elapsedTime = seconds()  ' start time
'      --- do the action ---
'   elapsedTime = seconds() - elapsedTime
'
' To accumulate "split" timings:
'   Dim elapsedTime as Double
'   elapsedTime = seconds()  ' base time; say, 100 sec's
'      --- do the first timed action ---
'   elapsedTime = seconds() - elapsedTime  ' say, 110 - 100 = 10 secs
'      --- do an untimed action ---
'   elapsedTime = seconds() - elapsedTime  ' say, 130 - 10 = 120
'      --- do the second timed action ---
'   elapsedTime = seconds() - elapsedTime  ' say, 160 - 120 = 40
Static base_s As Currency  ' base time; initializes to 0
Static freq_s As Currency  ' clock frequency; initializes to 0
Const ErrorFlag_c As Double = -86400#  ' error flag; normally impossible
If resetBase Then freq_s = 0@
If 0@ = freq_s Then  ' routine not initialized, or could not read frequency
  QueryPerformanceFrequency freq_s  ' try to read frequency
  ' if frequency is good, try to read base time (else it stays at 0)
  If 0@ <> freq_s Then QueryPerformanceCounter base_s
End If
' if we have a good base time, then we must have a good frequency also
If 0@ <> base_s Then
  Dim pTime As Currency
  QueryPerformanceCounter pTime
  If 0@ <> pTime Then  ' values are good
    seconds = (pTime - base_s) / freq_s
  Else  ' error when getting the present time - return special error value
    seconds = ErrorFlag_c
  End If
Else  ' error when getting the base time - return special error value
  seconds = ErrorFlag_c
End If
End Function

' ==============================================================================
Public Sub selectionToIrfanView()
Attribute selectionToIrfanView.VB_Description = "Copy the selection to IrfanView as a bitmap"
' Copy the selection to IrfanView as a bitmap. IrfanView can then convert it
' to almost any desired format, with optional picture edits first.
' The "crawling dots" copy border will appear; remove it with the Esc key;
' this will also remove the selection's bitmap image from the clipboard.
' Note that you can get the same result by hand. Select a region, press Ctrl-C,
' open IrfanView and be sure it is the active window, press Ctrl-V, and save to
' your desired format.
' WARNING! this routine replaces all clipboard data
Selection.copy
' we hard-code the path to IrfanView (below "Program Files") into this routine
' TODO check that this works correctly with "Program Files (x86)" in Win 7+
Const Viewer_c As String = "\IrfanView\i_view32.exe"
Dim thePath As String
thePath = getSysFolderPath(CSIDL_PROGRAM_FILES, SHGFP_TYPE_CURRENT) & Viewer_c
Shell thePath, vbNormalFocus
Application.SendKeys "^v"  ' paste
End Sub

' ==============================================================================
Public Sub setComment( _
  ByRef theRange As Range, _
  ByVal commentText As String, _
  Optional ByVal textSize As Single = 8!, _
  Optional ByVal textBold As Boolean = False, _
  Optional ByVal lineWeight As Single = 1!)
Attribute setComment.VB_Description = "Set the comment in the upper left cell of the supplied range to the supplied text"
' Set the comment in the upper left cell of the supplied range to the supplied
' text. Use 'vbLf' to force a new line (i.e., text1 & vbLf & text2).
' Note: Excel does not consider a cell to be changed when its comment changes
Dim topLeft As Range
' get top left Cell, even if 'theRange' is a big Range
Set topLeft = theRange.item(1&, 1&)
With topLeft
  ' remove any existing comment
  If Not .Comment Is Nothing Then .ClearComments
  ' add comment object & comment text
  .AddComment commentText
  ' set size & boldness
  .Comment.Shape.TextFrame.Characters.Font.size = textSize
  .Comment.Shape.TextFrame.Characters.Font.Bold = textBold
  ' set the line width
  .Comment.Shape.Line.Weight = lineWeight
  ' fit the enclosure to the text
  .Comment.Shape.TextFrame.AutoSize = True  ' fit to text
  ' square up large comments
  If Len(commentText) > 25& Then
    Dim boxArea As Single
    boxArea = .Comment.Shape.Height * .Comment.Shape.Width  ' get text area
    .Comment.Shape.Height = 1.1! * Sqr(boxArea)  ' force to be roughly square
    .Comment.Shape.Width = 1.1! * Sqr(boxArea)
  End If
End With
End Sub

' ==============================================================================
Public Sub setNameTo( _
  ByVal theName As String, _
  ByVal val As Variant, _
  Optional ByVal offsetRow As Long = 0&, _
  Optional ByVal offsetCol As Long = 0&)
Attribute setNameTo.VB_Description = "Set the cell in an Excel workbook that has the supplied name to the supplied value"
' Set the cell in an Excel workbook that has the supplied name to the supplied
' value. Raise an error if the name does not exist, or if it refers to a
' constant or a formula instead of a cell.
Const ID_C As String = File_c & "setNameTo"
On Error GoTo ErrorHandler

checkIfNameExists theName  ' there might not be any such Name
Dim nam As Name
Set nam = ActiveWorkbook.Names(theName)
On Error Resume Next  ' theName might not refer to a Range
Dim rng As Range
Set rng = nam.RefersToRange ' try to get the associated Range
Dim errNum As Long
errNum = Err.Number
On Error GoTo ErrorHandler
If 0& = errNum Then  ' theName refers to a Range
  ' set content of upper left Cell of Range, perhaps offset
  rng.item(1&, 1&).offset(offsetRow, offsetCol).Value = val
Else  ' the Name refers to a constant or a formula
  Dim str As String
  str = nam.RefersTo  ' gets a Variant (containing a String, in this case)
  If """" = Mid$(str, 2&, 1&) Then ' second character is quote
    ' Name contains a quoted text constant, in the form ="text constant"
    Err.Raise 5&, ID_C, "Name not Range; refers to text constant " & str
  Else
    ' Name contains a contant, in the form =42, or a formula like =COS(0.1)
    str = Mid$(str, 2&)  ' strip off leading equals sign
    Err.Raise 5&, ID_C, "Name not Range; refers to constant or formula " & str
  End If
End If
Exit Sub  '---------------------------------------------------------------------

ErrorHandler:  '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
Dim errDes As String
errDes = Err.Description
' supplement text; did error come from a called routine or this routine?
errDes = errDes & vbLf & _
  IIf(0& < InStr(errDes, "Problem in"), "Called by ", "Problem in ") & ID_C & _
  IIf(0& <> Erl, " line " & Erl, vbNullString)
' re-raise error with this routine's ID as Source, and appended to Description
Err.Raise Err.Number, ID_C, errDes
Resume  ' if debugging, set Next Statement here and F8 back to error point
End Sub

' ==============================================================================
Public Function stringToBoolean(ByVal theString As String) As Boolean
Attribute stringToBoolean.VB_Description = "Convert String text to a Boolean. Recognizes True/False, T/F, Yes/No, Y/N, On/Off, 1/0, -1/0. Error 13 if no-can-do"
' Convert String text to a Boolean; raise informative error if it can't be done.
' Recognizes True/False, T/F, Yes/No, Y/N, On/Off, 1/0, -1/0
Const ID_C As String = File_c & "stringToBoolean"
Dim lcArg As String
lcArg = Trim$(LCase$(theString))  ' convert to lower case & trim off blanks
If ("true" = lcArg) Or ("t" = lcArg) Or ("yes" = lcArg) Or ("y" = lcArg) _
   Or ("on" = lcArg) Or ("1" = lcArg) Or ("-1" = lcArg) Then
  stringToBoolean = True
ElseIf ("false" = lcArg) Or ("f" = lcArg) Or ("no" = lcArg) Or ("n" = lcArg) _
       Or ("off" = lcArg) Or ("0" = lcArg) Then
  stringToBoolean = False
Else
  Const TypeMismatch_c As Long = 13&
  Err.Raise TypeMismatch_c, ID_C, _
    Error$(TypeMismatch_c) & ": could not convert string" & vbLf & _
      """" & theString & """ to a Boolean" & vbLf & _
      "Expected True/False, T/F, Yes/No, Y/N, On/Off, 1/0, -1/0" & vbLf & _
      "Problem in " & ID_C
End If
End Function

' ==============================================================================
Public Function stringToDouble(ByVal theString As String) As Double
Attribute stringToDouble.VB_Description = "Convert String text to a Double. Error 13 or 6 if no-can-do"
' Convert String text to a Double; raise informative error if it can't be done.
' Note: absolute values less than "2.48E-324" will return zero with no warning.
Const ID_C As String = File_c & "stringToDouble"
On Error Resume Next
stringToDouble = CDbl(theString)
Dim errNum As Long
errNum = Err.Number
On Error GoTo 0
If 0& <> errNum Then  ' CDbl failed
  ' most likely Err.Number is 13 = "Type mismatch"; might be 6 = "Overflow"
  Err.Raise errNum, ID_C, _
    Error$(errNum) & ": could not convert string" & vbLf & _
      """" & theString & """ to a Double" & vbLf & _
      "Problem in " & ID_C
End If
End Function

' ==============================================================================
Public Function stringToLong(ByVal theString As String) As Long
Attribute stringToLong.VB_Description = "Convert String text to a Long. Error 13 or 6 if no-can-do"
' Convert String text to a Long; raise informative error if it can't be done.
' There is no error if a fractional part exists; the value is integer-rounded.
' Note that CLng uses "unbiased rounding" (odd rounds up, even rounds down, the
' IEEE 754 standard); thus 1.5 -> 2 and 2.5 -> 2
Const ID_C As String = File_c & "stringToLong"
On Error Resume Next
stringToLong = CLng(theString)
Dim errNum As Long
errNum = Err.Number
On Error GoTo 0
If 0& <> errNum Then  ' CLng failed
  ' most likely Err.Number is 13 = "Type mismatch"; might be 6 = "Overflow"
  Err.Raise errNum, ID_C, _
    Error$(errNum) & ": could not convert string" & vbLf & _
      """" & theString & """ to a Long" & vbLf & _
      "Problem in " & ID_C
End If
End Function

' ==============================================================================
Public Function supportVersion(Optional ByVal trigger As Variant) As String
Attribute supportVersion.VB_Description = "Return the date of the latest revision to this code, as a string in the format 'yyyy-mm-dd'"
' Return the date of the latest revision to this code, as a string in the
' format "yyyy-mm-dd"
supportVersion = Version_c
End Function

'===============================================================================
Public Sub switchToExcel(Optional ByVal recalc As Boolean = True)
Attribute switchToExcel.VB_Description = "Restore Excel back to its state before 'SwitchToVba' was called"
' Restore Excel back to its state before 'SwitchToVba' was called.
' WARNING! do not recalculate before raising an error; End will loop forever

On Error Resume Next  ' we don't want to stop on errors here
With Application
  ' allow user responses
  .DisplayAlerts = True
  ' turn firing of all Events back on
  .EnableEvents = True
  ' let Excel handle the StatusBar
  .StatusBar = False
  ' restore the VBE window (if it was previously Visible)
  .VBE.MainWindow.Visible = vbeWindow_m
  ' this has the unpleasant side effect of bringing it to the top, even if it
  ' was previously in the back, so force Excel to the top no matter what
  BringWindowToTop .hwnd
  ' Excel is supposed to do this when it gets control; we do it here to be sure
  .ScreenUpdating = True
  ' restore the cursor
  .Cursor = xlDefault

  ' put display back at the correct Workbook, Worksheet and Cell
  If Not (userCell_m Is Nothing) Then
    With userCell_m
      .Worksheet.Parent.Activate  ' the Workbook
      .Worksheet.Activate         ' the Worksheet
      .Activate                   ' the Cell
    End With
    Set userCell_m = Nothing  ' release object memory
  End If
  ' make Excel's calculation automatic (this may force a Calculation)
  If recalc Then .Calculation = xlCalculationAutomatic
End With
On Error GoTo 0  ' don't return live error back to caller (yes!)
End Sub

'===============================================================================
Public Sub switchToVba(Optional ByRef notice As String = "Busy...")
Attribute switchToVba.VB_Description = "Save Excel's state for restoration later by 'SwitchToExcel'"
' Save Excel's state for restoration later by 'SwitchToExcel'. Put the supplied
' message into Excel's StatusBar.
On Error Resume Next  ' we don't want to stop on errors here
With Application
  ' set the cursor to 'busy'
  .Cursor = xlWait
  ' put a message in the StatusBar
  .DisplayStatusBar = True
  .StatusBar = notice
  ' save the user's active Cell (and therefore Sheet and WorkBook)
  ' if no cell is active, Application.ActiveWindow.ActiveCell returns Nothing
  Set userCell_m = .ActiveWindow.activeCell
  ' turn off calculation
  .Calculation = xlCalculationManual
  ' turn off ScreenUpdating to speed up processing and avoid screen flicker
  .ScreenUpdating = False
  ' we do this in case there are many WorkSheet.Change Events, or the like
  ' also, with no Events firing, users can't affect the code while it is running
  .EnableEvents = False
  ' calculation can be much faster if the Visual Basic Editor (VBE) is closed
  vbeWindow_m = .VBE.MainWindow.Visible
  .VBE.MainWindow.Visible = False
End With
On Error GoTo 0  ' don't return live error back to caller (yes!)
End Sub

'===============================================================================
Public Function systemEnv(Optional ByVal trigger As Variant) As String
Attribute systemEnv.VB_Description = "Return 'Windows' if running under Windows, or 'Macintosh' if on a Mac"
' Return "Windows" if running under Windows, or "Macintosh" if on a Mac.
' This is similar to the INFO("system") worksheet function.
' Note that you may have to use conditional blocks with "Mac" to isolate code
' from the two environments, since incompatibilities cause errors. For instance:
' #If Mac Then  ' running under MacOS
'   Macintosh-specific code goes here
' #Else  ' running under Windows
'   Windows-specific code goes here
' #End If
If Application.OperatingSystem Like "*Mac*" Then
  systemEnv = "Macintosh"
Else
  systemEnv = "Windows"
End If
End Function

'===============================================================================
Public Sub textToClipboard(ByRef theText As String)
Attribute textToClipboard.VB_Description = "Put the supplied text into the clipboard. If the text has multiple lines, separate them with vbCrLf's, changing vbLf's to vbCrLf's if needed"
' Put the supplied text into the clipboard. If the text has multiple lines,
' separate them with vbCrLf's, changing vbLf's to vbCrLf's if needed.
' nb: if you know the CLSID associated with a COM object you can use the
' CreateObject function directly on the CLSID, without needing a Reference
' to the associated library (FM20.dll here), by using the following syntax:
' Set myObj = CreateObject("new:{<the CLSID>}")
' note that this usage gives late (runtime) binding, so no compile checking
' The CLSID for the Microsoft Forms DataObject is:
' 1C3B4210-F441-11CE-B9EA-00AA006B1A69
Const DataObject_Binding_c As String = _
  "new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}"
On Error Resume Next  ' we don't want to stop on errors here
' clipboard users may expect CR-LF at line end, not just LF; fix this up
' TODO fix this so it works if both EOL markers are present at the same time
Dim needsCr As Boolean  ' True if there is >= 1 vbLf, but no vbCrLf's
If 0& < InStr(theText, vbLf) Then needsCr = (0& = InStr(theText, vbCrLf))
With CreateObject(DataObject_Binding_c)  ' make & use anonymous DataObject
  If needsCr Then
    .SetText Replace(theText, vbLf, vbCrLf)  ' change all vbLf to vbCrLf
  Else
    .SetText theText  ' no isolated vbLf's, so no need to change text
  End If
  .PutInClipboard
End With  ' DataObject goes out of scope here, and is destroyed
On Error GoTo 0  ' don't return live error back to caller (yes!)
End Sub

'===============================================================================
Public Function timeStamp(Optional ByRef separator As String = "~") As String
Attribute timeStamp.VB_Description = "Return date & time of day to within 100 microseconds relative accuracy, as a string with the format 'YYYY-MM-DD_HH~MM~SS.SSSS'"
' Return date & time of day to within 100 microseconds relative accuracy, as a
' string with the format "YYYY-MM-DD_HH~MM~SS.SSSS" (absolute accuracy only good
' in the seconds and up, and then only as accurate as the system's clock).
' This function itself takes more than 500 microseconds on a 3 GHz processor,
' so you probably won't get the same value on successive calls.
' The optional argument changes the HMS separator from "~" to the supplied text.
' See comments in hms() function regarding jumpy fractional seconds.
timeStamp = Format$(Now(), "yyyy-mm-dd_") & hms(separator)
End Function

'===============================================================================
Public Function userDocs(Optional ByVal trigger As Variant) As String
Attribute userDocs.VB_Description = "Return the path to the user's 'My Documents' or 'Documents' folder"
' Return the path to the user's "My Documents" or "Documents" folder.
userDocs = getSysFolderPath(CSIDL_PERSONAL, SHGFP_TYPE_CURRENT)
End Function

'===============================================================================
Public Function userName(Optional ByVal trigger As Variant) As String
Attribute userName.VB_Description = "Return the OS record of the user name"
' Return the OS record of the user name.
userName = Application.userName
End Function

'===============================================================================
Public Function uniqueID(Optional ByVal setTo As Variant) As Double
Attribute uniqueID.VB_Description = "Return a sequence of increasing integer values, in a Double. On each call, the return is one more than on the previous call"
' Return a sequence of increasing integer values, in a Double. On each call,
' the return is one more than on the previous call. First value returned is 0.
' After 9.007199254740992E+15 values, the sequence quits with an error. If the
' optional argument is supplied, the sequence is set to the supplied value and
' that value is returned.
Static id_s As Double   ' ID number that will be issued - initializes to 0
Const maxCount_c As Double = 9.00719925474099E+15 + 2#
If Not IsMissing(setTo) Then  ' user has supplied a starting value
  ' note: IsNumeric perversely returns True for an Empty Variant
  If IsNumeric(setTo) And (Not IsEmpty(setTo)) Then
    setTo = CDbl(setTo)
  Else
    setTo = 0#  ' safety
  End If
  If -maxCount_c > setTo Then setTo = -maxCount_c
  If maxCount_c - 1& < setTo Then setTo = maxCount_c - 1&  ' allow 1 return
  id_s = setTo
End If
uniqueID = id_s
If maxCount_c > id_s Then  ' we are not too big
  id_s = id_s + 1#  ' stops adding at 9,007,199,254,740,992
Else  ' we are so big that adding 1 makes no change, so fail
  Const Oops_c As Long = 6&  ' Overflow
  Const ID_C As String = File_c & "uniqueID"
  Err.Raise Oops_c, ID_C, _
    Error$(Oops_c) & vbLf & _
    "Return value >= maximum value of " & maxCount_c & vbLf & _
    "Try starting with uniqueID(-" & maxCount_c & ")" & vbLf & _
    "to get twice as many values." & vbLf & _
    "Problem in " & ID_C
End If
End Function

'===============================================================================
Public Function vbaSpeed(Optional ByVal trigger As Variant) As Double
Attribute vbaSpeed.VB_Description = "Return relative speed of this CPU & operating system when running VBA code"
' Return relative speed of this CPU & operating system when running VBA code.
' Most modern systems, especially portables, run the CPU clock at a slower
' speed while idle, and speed up under load. Thus the reported speed may not
' reflect the system's maximum performance unless this function is called just
' after a heavy task, or during a heavy task on a multi-core system.
' Note that VBA is single-threaded, so multiple threads or cores can't increase
' this speed, except (slightly) by handling other tasks while VBA is running.
Dim j As Long, k As Long, t As Double, tLo As Double, x As Double
tLo = 1E+99
For k = 1& To 30&  ' do several runs and report the fastest (see above)
  t = seconds()
  For j = 1& To 10000&  ' repeat a bunch of stuff that uses lots of DLL code
    If (j Mod 2&) Then  ' alternate between two equivalent tasks
      x = Cos(Sin(x + 1.2345)) - 0.01 * Exp(Log(Abs(x) + 0.1))
    Else
      x = 0.01 * Exp(Log(Abs(x) + 0.1)) - Cos(Sin(x + 1.2345))
    End If
  Next j
  t = seconds() - t
  If tLo > t Then tLo = t
Next k
vbaSpeed = Round(0.01 / tLo, 3&)  ' scale result to somewhere near unity
End Function

'===============================================================================
Public Function windowsVersion(Optional ByVal trigger As Variant) As String
Attribute windowsVersion.VB_Description = "Return the name of the operating system VBA is running this code under"
' Return the name of the operating system VBA is running this code under.
' This needs WMI CORE download on older OS versions such as Win 98 & NT 4.0
' See 'opSys' for an alternative.
Dim os As Object
' "winmgmts:" gives access to Windows Management Instrumentation service
' InstancesOf returns the set of "Win32_OperatingSystem" entries in the registry
For Each os In GetObject("winmgmts:").InstancesOf("Win32_OperatingSystem")
  ' strip off the "Microsoft Windows" part; it's a waste of space
  windowsVersion = Replace( _
    os.Caption & " " & os.CSDVersion & " build " & os.BuildNumber, _
    "Microsoft Windows ", vbNullString)
Next os
Set os = Nothing
End Function

'===============================================================================
Public Function xlsFileName(Optional ByVal trigger As Variant) As String
Attribute xlsFileName.VB_Description = "Returns the name of the file the active workbook was read from (i.e., 'Sumthin.xlsm')"
' Returns the name of the file the active workbook was read from (i.e.,
' "Sumthin.xlsm"). For this to work, a new workbook must first be saved to disk.
If 0& = Len(Excel.ActiveWorkbook.path) Then  ' workbook has never been saved
  xlsFileName = "<not saved to disk>"
Else
  xlsFileName = Excel.ActiveWorkbook.Name
End If
End Function

'===============================================================================
Public Function xlVersion(Optional ByVal trigger As Variant) As String
Attribute xlVersion.VB_Description = "Return major Excel version and revision number as 'MM.RRRR Build NNNN'"
' Return major Excel version and revision number as "MM.RRRR Build NNNN"
' The Excel version is the same as the Office version. Versions include:
' XL 5 is major version 5
' XL 95 is major version 7
' XL 97 is major version 8
' XL 2000 is major version 9
' XL 2002 is major version 10
' XL 2003 is major version 11
' XL 2007 is major version 12
' XL 2010 is major version 14
' XL 2013 is major version 15
xlVersion = CStr(0.0001 * Application.CalculationVersion) & _
  "  Build " & Application.Build
End Function

'########## Private Support Routine ############################################

'-------------------------------------------------------------------------------
Private Function getSysFolderPath( _
  csidl As Long, _
  SHGFP_TYPE As Long, _
  Optional perUserInit As Boolean = False, _
  Optional forceNonAlias As Boolean = False, _
  Optional unverifiedPath As Boolean = False) _
As String
' Return the path to the specified system folder
Dim buff As String
Dim dwFlags As Long
'fill buffer with the specified folder item
buff = Space$(MAX_LENGTH)
If perUserInit Then dwFlags = dwFlags Or CSIDL_FLAG_PER_USER_INIT
If forceNonAlias Then dwFlags = dwFlags Or CSIDL_FLAG_NO_ALIAS
If unverifiedPath Then dwFlags = dwFlags Or CSIDL_FLAG_DONT_VERIFY
If SHGetFolderPath(vbNull, csidl Or dwFlags, -1, SHGFP_TYPE, buff) = S_OK Then
  getSysFolderPath = Left$(buff, strLenW(StrPtr(buff)))
End If
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


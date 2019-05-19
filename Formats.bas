Attribute VB_Name = "Formats"
Attribute VB_Description = "Exports routines useful in formatting output strings. Devised and coded by John Trenholme."
'
'###############################################################################
'#
'# VBA Module file "Formats.bas"
'#
'# VB routines useful in formatting output strings.
'#
'# Started 2010-01-27 by John Trenholme (using older routines)
'#
'# Exports the routines:
'#   Function cram
'#   Function digitRound
'#   Function eng <under construction>
'#   Function engToDbl <under construction>
'#   Function fmt
'#   Function fSng
'#   Function ss
'#   Function strConst
'#   Function strFitL
'#   Function strFitR
'#   Function strSplit
'#   Function toStr
'#
'###############################################################################

Option Base 0          ' array base value when not specified     - default
Option Compare Binary  ' string comparison based on Asc(char)    - default
Option Explicit        ' force explicit declaration of variables - not default
'Option Private Module  ' no effect in VB6; globals project-only in VBA

' Module-global Const values (convention: start with upper-case; suffix "_c")

Private Const Version_c As String = "2012-12-04"
Private Const File_c As String = "Formats[" & Version_c & "]."

Private Const BadArg_c As Long = 5&  ' = "Invalid procedure call or argument"

'===============================================================================
Public Function FormatsVersion(Optional ByVal trigger As Variant) As String
Attribute FormatsVersion.VB_Description = "Date of the latest revision to this code, as a string with format 'YYYY-MM-DD'"
' Date of the latest revision to this code, as a string with format "YYYY-MM-DD"
FormatsVersion = Version_c
End Function

'===============================================================================
Public Function cram( _
  ByVal x As Double, _
  Optional ByVal size As Long = 6&, _
  Optional ByVal useN As Boolean = True, _
  Optional ByVal floatZero As Boolean = True) _
As String
Attribute cram.VB_Description = "Turn the input number into a string of the specified length, keeping as much numeric precision as is possible"
' Turn the input number x into a string of the specified size, keeping as much
' numeric precision as is possible. The minimum size you can specify is 4.
' The default value is 6. If the size is less than 6, numbers with large
' exponents (2 or 3 digits) may yield strings whose size is larger by 1 (if
' size = 5) or by 2 (if size = 4) than the requested value (1 more than these
' values if "useN" is False). If there are spaces in the result, they are on
' the left (that is, the result is right-justified).
'
' Negative exponents are returned as "N" rather than "E-", to give space for
' one more significant digit. Thus "cram( 1.45E-19, 6&)" yields "1.5N19".
' You can inhibit this behavior by setting the optional argument "useN" to
' False, but you will lose a digit when a negative exponent is needed, and
' you must also use a size of 7 or more to accomodate all possible numbers if
' you don't want some returned strings to be larger than requested.
'
' If you need to convert a string "s" that may have an "N" exponent character in
' it back to a number, you can use the form Val(Replace(s, "N", "E-")).
'
' Input values of exactly zero are returned as "0.0" unless the optional
' argument "floatZero" is False, in which case just "0" is returned.
'
' Note: to get full 15-digit values for any possible exponent, set size >= 21.
'       for 6-digit accuracy, set size = 12
'       for K-digit accuracy, set size = 6 + K
'       add 1 to these values if "useN" is False
'
' Because most of the work is done in VB library formatting routines, this
' function is fast - it takes about 20 microseconds on a 3 GHz Pentium 4.

' enforce some sanity - must support at least -1E9 and -1N9, so size >= 4
If 4& > size Then size = 4&
' allow space for minus sign to be added later - sizM is >= 3 after this
Dim sizM As Long
If 0# > x Then sizM = size - 1& Else sizM = size
Dim ax As Double
ax = Abs(x)  ' work with positive number; add sign on later

' first, see if the default conversion will work (for few-digit numbers)
Dim s As String
s = CStr(ax)  ' default VB conversion; will have up to 15 digits of precision
' now tweak the default behaviour to improve it (if possible)
If 0# = x Then
  ' handle zero: if floating, change integer-appearing "0" to "0.0"
  If floatZero Then s = "0.0"
  s = Space$(size - Len(s)) & s
ElseIf (sizM >= Len(s)) And (ax >= 0.000000000000001) Then
  ' for values below 1E-15, CStr() forces E-xx format; result may be too short
  ' may contain exponent sign as E+ or E-; shorten if it does
  s = Replace(s, "E+", "E")  ' so positive exponent values are always signless
  ' we use N for "negative" exponents; also no sign (caller can inhibit this)
  If useN Then s = Replace(s, "E-", "N")
  ' pad on left with spaces so right-justified
  s = Right$(Space$(sizM - 1&) & s, sizM)
Else  ' default was too long, try tweaking
  ' try to fit it in by trimming digits - get min-length scientific format
  Dim digits As Long
  If 15& > sizM Then digits = sizM Else digits = 15&  ' avoid trailing zeros
  s = Format$(ax, String$(digits, "0") & "E-0")  ' digits, then E part
  ' find location, and value, of power of 10
  Dim eLoc As Long
  eLoc = InStr(s, "E")  ' must exist
  Dim p10 As Long
  p10 = val(Mid$(s, eLoc + 1&))  ' extract & convert power of 10
  Dim fmt As String
  If 0& < p10 Then  ' we must add decimal point & positive exponent
    ' start at largest possible accuracy and shrink to fit
    If 4& <= sizM Then fmt = "0." & String$(sizM - 4&, "0") Else fmt = "0"
    Do
      s = Format$(ax, fmt & "E-0")
      fmt = Left$(fmt, Len(fmt) - 1&)  ' try fewer digits
    Loop Until (Len(s) <= sizM) Or (Len(fmt) = 0&)
  ElseIf 0& = p10 Then  ' will fit as an integer - no decimal point needed
    s = Left$(s, sizM)
  ElseIf -sizM <= p10 Then  ' needs decimal point but no exponent
    s = Format$(ax, String$(sizM + p10, "0") & "." & String$(-p10 - 1&, "0"))
    If Len(s) > sizM Then
      ' maybe a number like 9.999 got rounded up to greater length
      Dim last As String
      last = Right$(s, 1&)
      If ("0" = last) Or ("." = last) Then s = Left$(s, sizM)  ' if so, fix
    End If
  Else  ' must add decimal point & negative exponent
    Dim sizP As Long
    If useN Then sizP = 1& Else sizP = 0&
    ' start at largest possible accuracy and shrink to fit
    If 4& <= sizM Then fmt = "0." & String$(sizM - 4&, "0") Else fmt = "0"
    Do
      s = Format$(ax, fmt & "E-0")
      fmt = Left$(fmt, Len(fmt) - 1&)  ' try fewer digits
    Loop Until (Len(s) <= sizM + sizP) Or (Len(fmt) = 0&)
    If useN Then
      ' make negative exponent take only one space (caller can inhibit this)
      s = Replace(s, "E-", "N")
      ' fix cases where, for example, 9.95E-10 in width 6 becomes 1.00N9
      If Len(s) < sizM Then s = Replace(s, "N", "0N")
    Else
      ' fix cases where, for example, 9.95E-10 in width 7 becomes 1.00E-9
      If Len(s) < sizM Then s = Replace(s, "E", "0E")
    End If
  End If
End If
If 0# > x Then  ' we need to restore the minus sign
  s = LTrim$(s)  ' might be leading spaces for large size; remove them first
  s = "-" & s
  If Len(s) < size Then s = Space$(size - Len(s)) & s  ' restore spaces
End If
cram = s
End Function

'==============================================================================
Public Function craz( _
  ByVal x As Double, _
  Optional ByVal size As Long = 6&, _
  Optional ByVal useN As Boolean = True, _
  Optional ByVal floatZero As Boolean = True) _
As String
' Produce a number-crammed string that is formatted for insertion into text.
' Sends the input to "cram" with the supplied arguments, and then add a leading
' zero if needed because result starts with ".", and remove any leading blanks
' and trailing zeros. Note that the result may be shorter than cram's result,
' or possibly one longer if a lead zero was added. In any case, it has the
' minimum length needed to give the requested size. Example usage:
'  Print "The length is " & craz(xx, 12&) & " kilometers"
' This will cram "xx" into 12 or fewer characters
Dim result As String
result = Trim$(cram(x, size, useN, floatZero))
If "." = Left$(result, 1&) Then
  result = "0" & result
ElseIf "-." = Left$(result, 2&) Then
  result = "-0" & Mid$(result, 2&)
End If
Dim expPos As Long, expPart As String
expPos = InStr(result, "E")  ' look for Enn
If 0& = expPos Then expPos = InStr(result, "N")  ' try for Nnn
If 0& < expPos Then  ' there's an exponent, so split the result
  expPart = Mid$(result, expPos)
  result = Left$(result, expPos - 1&)
Else
  expPart = vbNullString
End If
Do While "0" = Right$(result, 1&)  ' strip trailing zeros
  result = Left$(result, Len(result) - 1&)
Loop
craz = result & expPart
End Function

'===============================================================================
Public Function DigitRound( _
  ByVal valToRound As Double, _
  Optional ByVal numDigits As Integer = 6) _
As Double
Attribute DigitRound.VB_Description = "Round the input value to the specified number of significant digits. By default, change 'E-' in exponents to 'N'."
' Round the input value to the specified number of significant digits. Used
' (among other things) to keep lengths short when printed. Returns valToRound
' unchanged if it is zero. If numDigits < 1 it is set to 1; if numDigits > 15
' it is set to 15. This routine takes about 2 microseconds on a 3 GHz Pentium
' 4. It is accurate to one or two low-order bits.
Const ID_C As String = File_c & "digitRound Function"
On Error GoTo ErrorHandler
If valToRound = 0# Then
  DigitRound = 0#  ' special case; simple to do, and causes Log(0) problem
Else
  If numDigits < 1 Then numDigits = 1  ' silently fix invalid argument values
  If numDigits > 15 Then numDigits = 15
  ' this constant is accurate to the last bit when written as a sum of 2 parts
  Const Log10_e As Double = 0.43429448 + 1.903251828E-09  ' makes Log -> Log10
  Dim pow10 As Double
  pow10 = numDigits - 1# - Int(Log(Abs(valToRound)) * Log10_e)  ' digit shift
  Dim scaling As Double
  scaling = 10# ^ (0.5 * pow10)  ' avoid overflow for very small input values
  Dim scaled As Double  ' we will put digits we want into integer part of this
  scaled = (valToRound * scaling) * scaling
  DigitRound = (Int(scaled + 0.5) / scaling) / scaling
End If
Exit Function  '----------------------------------------------------------------
ErrorHandler:  '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
Err.Raise Err.Number, ID_C, Err.Description & vbLf & _
  "valToRound = " & valToRound & "  numDigits = " & numDigits & vbLf & _
  "Problem in " & ID_C
Resume  ' to debug at error statement, "set next statement" here & single-step
End Function

'===============================================================================
Public Function eng( _
  ByVal x As Double, _
  Optional ByVal size As Long = 6&) _
As String
Attribute eng.VB_Description = "NOT CODED Change number to value 1 <= x < 1000 with engineering suffix, if possible"
' TODO - use standard engineering suffix characters to get more precision
'   positive exponents:  k  M  G  T  P  E  Z  Y
'   negative exponents:  m  µ  n  p  f  a  z  y  (note: µ is Chr$(181&))
'   e.g. 2.615E8 -> 261.47M
'   also, supply a read function that reverses the process
Stop  ' work in progress
End Function

'===============================================================================
Public Function engToDbl( _
  ByVal str As String) _
As Double
Attribute engToDbl.VB_Description = "NOT CODED Convert number with engineering suffix to a Double"
' TODO - use standard engineering suffix characters to get more precision
'   positive exponents:  k  M  G  T  P  E  Z  Y
'   negative exponents:  m  µ  n  p  f  a  z  y  (note: µ is Chr$(181&))
'   e.g. 2.615E8 -> 261.47M
'   also, supply a read function that reverses the process
Stop  ' work in progress
End Function

'===============================================================================
Public Function fmt( _
  ByRef stringWithSlots, _
  ParamArray arguments() As Variant) _
As String
Attribute fmt.VB_Description = "Given a string with parameter slots %0, %1, ... this returns a string with the optional arguments (indexed as 0, 1, 2, ...) replacing the corresponding slots"
' Given a string with parameter slots %0, %1, ... this returns a string with
' the optional arguments (indexed as 0, 1, 2, ...) replacing the corresponding
' slots. To print %n where 'n' is a number in the argument range, supply %%n.
' If there is no argument 'n' the slot will be left as '%n'
' Example: fmt("%1 %0 %2 %%1 %3", "A", "B", "C") -> "B A C %1 %3"
Dim ret As String
Static hucs As String  ' highly unlikely character sequence (init on 1st use)
If 0& = Len(hucs) Then hucs = Chr(1&) & Chr(3&) & Chr(2&) & Chr(255&) & Chr(1&)
ret = Replace(stringWithSlots, "%%", hucs)  ' hide "%%"
If 0& <= UBound(arguments) Then  ' one or more arguments exist; else get -1
  Dim j As Long
  Dim tempStr As String
  ' note that a ParamArray is always 0-based, no matter what Option Base says
  For j = UBound(arguments) To 0& Step -1&  ' do %121 before %12 before %1 etc.
    If 0& <> InStr(ret, "%" & j) Then  ' slot j exists
      tempStr = toStr(arguments(j))
      ret = Replace(ret, "%" & j, tempStr)
    Else  ' no such slot
      ret = ret & "<NO SLOT %" & j & ">"
    End If
  Next j
End If
ret = Replace(ret, hucs, "%")  ' unhide "%%", change to "%"
fmt = ret
End Function

'===============================================================================
Public Function fSng(ByVal x As Double) As String
Attribute fSng.VB_Description = "Format with Single precision, but with excess ""0"" in exponent removed, and no leading blanks"
' Format with Single precision, but with excess "0" in exponent removed, and no
' leading blanks.
Dim str As String
str = CStr(CSng(x))
str = Replace$(str, "E+0", "E+")
str = Replace$(str, "E-0", "E-")
fSng = Trim$(str)
End Function

'===============================================================================
Public Function ss( _
  ByVal Count As Long, _
  Optional ByVal pad As Boolean = False) _
As String
Attribute ss.VB_Description = "Concatenate the result of this Function onto the name of a counted quantity in order to make it correctly singular or plural"
' Concatenate the result of this Function onto the name of a counted quantity
' in order to make it correctly singular or plural. That is...
'   CStr(j) & " coin" & ss(j) becomes "0 coins", "1 coin", "2 coins", etc.
' If the Optional argument is True, the return value will be "s" or " ",
' instead of "s" or "". This keeps the length of the returned value constant.
' Note: for irregular plurals (loaf goes to loaves) this won't work properly.
' In such cases, use (e.g.) CStr(j) & IIf(j = 1&, " loaf", " loaves")
If 1& = Count Then
  If pad Then ss = " " Else ss = vbNullString
Else
  ss = "s"
End If
End Function

'===============================================================================
Public Function strConst(ByVal theStr As String) As String
' Ruturn the input string formatted the way a string constant is formatted in
' Visual Basic, with quotes at each end and internal quotes replaced by double
' quotes. Thus Big "Bob" Jones becomes "Big ""Bob"" Jones".
strConst = """" & Replace(theStr, """", """""") & """"
End Function

'===============================================================================
Public Function strFitL(ByVal theStr As String, ByVal wide As Long) As String
Attribute strFitL.VB_Description = "Return a string of the specified width with the input string left-justified in it, blank(s) at end"
' Return a string of the specified width with the input string left-justified
' in it. The returned string will always end with a space. The input string
' will be padded with blanks to get the requested width, or truncated to fit.
If wide < 2& Then wide = 2&  ' silently enforce sanity
Dim res As String
res = Left$(theStr, wide - 1&) & " "
If Len(res) < wide Then res = res & String$(wide - Len(res), " ")
strFitL = res
End Function

'===============================================================================
Public Function strFitR(ByVal theStr As String, ByVal wide As Long) As String
Attribute strFitR.VB_Description = "Return a string of the specified width with the input string right-justified in it, blank(s) at front"
' Return a string of the specified width with the input string right-justified
' in it. The returned string will always start with a space. The input string
' will be padded with blanks to get the requested width, or truncated to fit.
If wide < 2& Then wide = 2&  ' silently enforce sanity
Dim res As String
res = " " & Left$(theStr, wide - 1&)
If Len(res) < wide Then res = String$(wide - Len(res), " ") & res
strFitR = res
End Function

'===============================================================================
Public Function strSplit(ByVal theStr As String, _
  Optional ByVal splitters As String = " ,;:", _
  Optional ByVal maxLen As Long = 80&, _
  Optional ByVal slack As Long = 5&, _
  Optional ByVal indent As String = "& ") _
As String
Attribute strSplit.VB_Description = "Split string into parts separated by vbNewLine's, so they will fit in desired print width"
' Split the input string up into logical lines separated by vbNewLine's if it
' is longer than 'maxLen.'  Split after any of the characters in 'splitters.'
' If there is no splitter that would result in a logical line between 'maxLen'
' and (maxLen-slack) in length, force a split that results in 'maxLen.' Indent
' any logical line after the first by "indent.' This function can be used to
' force a long string to print into a specified width.
Const ID_C As String = File_c & "strSplit Function"
On Error GoTo ErrorHandler
If 0& >= maxLen Then Err.Raise BadArg_c, ID_C, Error(BadArg_c) & vbLf & _
  "Wanted maxLen > 0 but got " & maxLen
If 0& > slack Then Err.Raise BadArg_c, ID_C, Error(BadArg_c) & vbLf & _
  "Wanted slack >= 0 but got " & slack
Dim minlen As Long
minlen = maxLen - slack - 1&
If minlen < 1& Then minlen = 1&
Dim res As String, temp As String
res = vbNullString
temp = theStr  ' remaining portion of input string; start at entire string
Do Until maxLen >= Len(temp)  ' Do if remaining portion is longer than desired
  Dim loc As Long, locMax As Long
  locMax = 0&
  Dim j As Long
  Dim splitChar As String
  For j = 1& To Len(splitters)  ' loop over potential split characters
    splitChar = Mid$(splitters, j, 1&)
    ' find the last location of the splitter, matching case (if that matters)
    loc = InStrRev(temp, splitChar, maxLen, vbBinaryCompare)
    ' update the longest distance we can split at, for any splitter
    If 0& < loc Then If minlen < loc Then locMax = loc
  Next j
  If 0& = locMax Then locMax = maxLen  ' no splitter worked; punt
  res = res & Left$(temp, locMax) & vbNewLine  ' chop off piece & add to result
  temp = indent & Mid$(temp, locMax + 1&)  ' remove piece from residual; indent
Loop
strSplit = res & temp
Exit Function  '----------------------------------------------------------------
ErrorHandler:  '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
Err.Raise Err.Number, ID_C, Err.Description & vbLf & _
  "Problem in " & ID_C
Resume  ' to debug at error statement, "set next statement" here & single-step
End Function

'===============================================================================
Public Function toStr( _
  ByRef thing As Variant) _
As String
' Convert the input Variant argument to a string if that is at all possible.
' A String will be returned in quotes, with internal quote characters changed
' to double-quotes (i.e., Hello, "Joe's" world -> "Hello, ""Joe's"" world").
' Arrays (up to 3D) are returned in parentheses as (val,val,...); the items
' in Collections are returned as [val,val,...] and the items in Dictionaries
' are returned as {val,val,...}. Most Objects are returned as their TypeName,
' followed by "@" and their address in hex, such as "MyClass@0024FA38".
Const ID_C As String = File_c & "toStr Function"

Dim vt As Integer
' note: for an object with a default property, VarType returns the type of the
' default property, not vbObject - use IsObject to be sure
vt = VarType(thing)
Dim tn As String
tn = TypeName(thing)

Dim tempStr As String
If Not IsArray(thing) Then  ' input argument is a scalar
  If IsObject(thing) Then  ' this works even if object has default property
    If 0& <> ObjPtr(thing) Then  ' it's an object
      Dim v As Variant
      ' check for objects we know how to handle
      If "Collection" = tn Then
        tempStr = "["
        For Each v In thing  ' from a Collection, you get the Items, in order
          tempStr = tempStr & toStr(v) & ","
        Next
        If 0& = thing.Count Then tempStr = "[Empty Collection,"
        tempStr = Left$(tempStr, Len(tempStr) - 1&) & "]"  ' kill comma, add ]
      ElseIf "Dictionary" = tn Then
        tempStr = "{"
        For Each v In thing  ' from a Dictionary, you get the keys
          tempStr = tempStr & toStr(thing.Item(v)) & ","  ' use key to get item
        Next
        If 0& = thing.Count Then tempStr = "{Empty Dictionary,"
        tempStr = Left$(tempStr, Len(tempStr) - 1&) & "}"  ' kill comma, add }
      Else  ' not a special case; report its class name and its address
        tempStr = _
          TypeName(thing) & "@" & Right("0000000" & Hex(ObjPtr(thing)), 8&)
      End If
    Else  ' it has an Object pointer, but the pointer is Null
      tempStr = "Nothing"  ' after all, "Nothing" is an Object
    End If
  ' not an Object; handle some special cases that don't work well with CStr
  ElseIf vbEmpty = vt Then  ' Empty = VarType 0
    tempStr = "Empty"
  ElseIf vbNull = vt Then  ' Null = VarType 1
    tempStr = "Null"
  ElseIf vbDate = vt Then  ' Date = VarType 7
    tempStr = "#" & CStr(thing) & "#"
  ElseIf vbString = vt Then  ' String = VarType 8
    tempStr = thing
  Else  ' try to use built-in conversion
    tempStr = CStr(thing)
  End If
  ' change nulls to printables (unassigned fixed-length string, or API result)
  tempStr = Replace(tempStr, vbNullChar, Chr$(127&))  ' OK in most fonts ""
  ' now clean up some CStr funny business
  If "Error 448" = tempStr Then
    tempStr = "Missing"  ' it's a missing argument
  ElseIf "Error" = Left$(tempStr, 5&) Then ' some Error; get descriptive text
    Dim desc As String
    desc = Error(val(Mid$(tempStr, 7&)))  ' peel off error number & get text
    If "Application-defined" = Left$(desc, 19&) Then desc = "User-defined"
    tempStr = tempStr & ": " & desc  ' add on descriptive text
    ' tempStr = """" & Replace(tempStr, """", """""") & """"  ' quote error string
  End If
Else  ' input argument is an array
  Dim ja As Long, jaLo As Long, jaHi As Long
  Dim jb As Long, jbLo As Long, jbHi As Long
  Dim jc As Long, jcLo As Long, jcHi As Long
  Dim jdLo As Long
  On Error Resume Next
  jaLo = LBound(thing)
  If 0& <> Err.Number Then  ' this is an uninitialized dynamic array
    On Error GoTo 0
    tempStr = "(uninitialized dynamic array)"
    GoTo SingleExitPoint
  End If
  tempStr = "("
  jaHi = UBound(thing)
  jbLo = LBound(thing, 2&)  ' try to get lower index of second dimension
  If 0& <> Err.Number Then  ' this is a 1D array
    On Error GoTo 0
    For ja = jaLo To jaHi
      tempStr = tempStr & toStr(thing(ja)) & ","  ' recursive call
    Next ja
    tempStr = Left$(tempStr, Len(tempStr) - 1&) & ")"  ' kill comma, add )
    GoTo SingleExitPoint
  End If
  jbHi = UBound(thing, 2&)
  jcLo = LBound(thing, 3&)  ' try to get lower index of third dimension
  If 0& <> Err.Number Then  ' this is a 2D array
    On Error GoTo 0
    For ja = jaLo To jaHi
      tempStr = tempStr & "("
      For jb = jbLo To jbHi
        tempStr = tempStr & toStr(thing(ja, jb)) & ","  ' recursive call
      Next jb
      tempStr = Left$(tempStr, Len(tempStr) - 1&) & ")" & ","
    Next ja
    tempStr = Left$(tempStr, Len(tempStr) - 1&) & ")"  ' kill comma, add )
    GoTo SingleExitPoint
  End If
  jcHi = UBound(thing, 3&)
  jdLo = UBound(thing, 4&)  ' try to get lower index of fourth dimension
  If 0& <> Err.Number Then  ' this is a 3D array
    On Error GoTo 0
    For ja = jaLo To jaHi
      tempStr = tempStr & "("
      For jb = jbLo To jbHi
        tempStr = tempStr & "("
        For jc = jcLo To jcHi
          tempStr = tempStr & toStr(thing(ja, jb, jc)) & ","  ' recursive call
        Next jc
        tempStr = Left$(tempStr, Len(tempStr) - 1&) & ")" & ","
      Next jb
      tempStr = Left$(tempStr, Len(tempStr) - 1&) & ")" & ","
    Next ja
    tempStr = Left$(tempStr, Len(tempStr) - 1&) & ")"  ' kill comma, add )
  Else  ' this is a 4D or more array; we can't handle it
    ' TODO add more code here to handle more complicated arrays (as needed)
    On Error GoTo 0
    Err.Raise BadArg_c, ID_C, _
      Error(BadArg_c) & vbLf & _
      "Cannot handle array with 4 or more dimensions" & vbLf & _
      "Problem in " & ID_C
  End If
End If
SingleExitPoint:
toStr = tempStr
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


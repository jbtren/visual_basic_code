Attribute VB_Name = "FormatMod"
Attribute VB_Description = "Exports routines useful in formatting output strings"
'
'###############################################################################
'#
'# VBA Module file "FormatMod.bas"
'#
'# VB routines useful in formatting output strings.
'#
'# Started 2010-01-27 by John Trenholme (using older routines)
'#
'# Exports the routines:
'#
'#   Function cram
'#   Function digitRound
'#   Function fmt
'#   Function s
'#   Function toStr
'#
'###############################################################################

Option Base 0  ' array base value when not specified - the default
Option Compare Binary  ' string comparison based on Asc(char) - the default
Option Explicit  ' force explicit declaration of variables - changes default

' Module-global Const values (convention: start with upper-case; suffix "_c")

Private Const Version_c As String = "2011-02-22"
Private Const File_c As String = "FormatMod[" & Version_c & "]."

' ==============================================================================
Public Function cram( _
  ByVal Number As Double, _
  Optional ByVal size As Long = 6&, _
  Optional ByVal useN As Boolean = True) _
As String
Attribute cram.VB_Description = "Turn the input number into a string of the specified size, keeping as much numeric precision as is possible"
' Turn the input number into a string of the specified size, keeping as much
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
' Note: to get full 15-digit values for any possible exponent, set size >= 21.
'       for 6-digit accuracy, set size = 12
'       for K-digit accuracy, set size = 6 + K
'       add 1 to these values if "useN" is False
'
' Because most of the work is done in VB library formatting routines, this
' function is fast - it takes about 20 microseconds on a 3 GHz Pentium 4.

' TODO - use standard engineering suffix characters to get more precision
'   positive exponents:  K  M  G  T  P  E
'   negative exponents:  m  µ  n  p  f  a  (note: µ is Chr$(181&))
'   e.g. 2.615E8 -> 261.47M
'   also, supply a read function that reverses the process

' enforce some sanity - must support at least -1E9 and -1N9, so size >= 4
If 4& > size Then size = 4&
' allow space for minus sign to be added later - size is >= 3 after this
Dim sizM As Long
If 0# > Number Then sizM = size - 1& Else sizM = size
Dim x As Double
x = Abs(Number)  ' work with positive number; add sign on later

' first, see if the default conversion will work (for few-digit numbers)
Dim s As String
s = CStr(x)  ' will have up to 15 digits of precision
If sizM >= Len(s) Then
  ' may contain exponent sign as E+ or E-,; shorten if it does
  s = Replace(s, "E+", "E")  ' so positive exponent values are always signless
  ' we use N for "negative" exponents; also no sign (caller can inhibit this)
  If useN Then s = Replace(s, "E-", "N")
  ' pad on left with spaces so right-justified
  s = Right$(Space$(sizM - 1&) & s, sizM)
Else  ' default was too long, try tweaking
  ' try to fit it in by trimming digits - get min-length scientific format
  Dim digits As Long
  If 15& > sizM Then digits = sizM Else digits = 15&  ' avoid trailing zeros
  s = Format$(x, String$(digits, "0") & "E-0")  ' digits, then E part
  ' find location, and value, of power of 10
  Dim eLoc As Long
  eLoc = InStr(s, "E")  ' must exist
  Dim p10 As Long
  p10 = Val(Mid$(s, eLoc + 1&))  ' extract & convert power of 10
  Dim fmt As String
  If 0& < p10 Then  ' we must add decimal point & positive exponent
    ' start at largest possible accuracy and shrink to fit
    If 4& <= sizM Then fmt = "0." & String$(sizM - 4&, "0") Else fmt = "0"
    Do
      s = Format$(x, fmt & "E-0")
      fmt = Left$(fmt, Len(fmt) - 1&)  ' try fewer digits
    Loop Until (Len(s) <= sizM) Or (Len(fmt) = 0&)
  ElseIf 0& = p10 Then  ' will fit as an integer - no decimal point needed
    s = Left$(s, sizM)
  ElseIf -sizM <= p10 Then  ' needs decimal point but no exponent
    s = Format$(x, String$(sizM + p10, "0") & "." & String$(-p10 - 1&, "0"))
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
      s = Format$(x, fmt & "E-0")
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
If 0# > Number Then  ' we need to restore the minus sign
  s = LTrim$(s)  ' might be leading spaces for large size; remove them first
  s = "-" & s
  If Len(s) < size Then s = Space$(size - Len(s)) & s  ' restore spaces
End If
cram = s
End Function

'===============================================================================
Public Function digitRound( _
  ByVal valToRound As Double, _
  Optional ByVal numDigits As Integer = 6) _
As Double
Attribute digitRound.VB_Description = "Round the input value to the specified number of significant digits"
' Round the input value to the specified number of significant digits. Used
' (among other things) to keep lengths short when printed. Returns valToRound
' unchanged if it is zero. If numDigits < 1 it is set to 1; if numDigits > 15
' it is set to 15. This routine takes about 2 microseconds on a 3 GHz Pentium
' 4. It is accurate to one or two low-order bits.
Const ID_c As String = File_c & "digitRound Function"
On Error GoTo ErrorHandler
If valToRound = 0# Then
  digitRound = 0#  ' special case; simple to do, and causes Log(0) problem
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
  digitRound = (Int(scaled + 0.5) / scaling) / scaling
End If
Exit Function  '----------------------------------------------------------------
ErrorHandler:  '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
Err.Raise Err.Number, ID_c, Err.Description & vbLf & _
  "valToRound = " & valToRound & "  numDigits = " & numDigits & vbLf & _
  "Problem in " & ID_c
Resume  ' to debug at error statement, "set next statement" here & single-step
End Function

'===============================================================================
Public Function FormatModVersion() As String
Attribute FormatModVersion.VB_Description = "The date of the latest revision to this module, formatted as ""yyyy-mm-dd"""
' The date of the latest revision to this module, formatted as "yyyy-mm-dd".
FormatModVersion = Version_c
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
  ' note that a ParamArray is always 0-based
  For j = UBound(arguments) To 0& Step -1&  ' do %12 before %1 etc.
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
Public Function s( _
  ByVal count As Long, _
  Optional ByVal pad As Boolean = False) _
As String
Attribute s.VB_Description = "Concatenate the result of this Function onto the name of a counted quantity in order to make it correctly singular or plural"
' Concatenate the result of this Function onto the name of a counted quantity
' in order to make it correctly singular or plural. That is...
'   j & " coin" & s(j) becomes "0 coins", "1 coin", "2 coins", etc.
' Note: for irregular plurals (loaf goes to loaves) this won't work properly.
' If the Optional argument is True, the return value will be "s" or " ",
' instead of "s" or "". This keeps the length of the returned value constant.
If 1& = count Then
  If pad Then s = " " Else s = vbNullString
Else
  s = "s"
End If
End Function

'===============================================================================
Public Function strFit(ByVal theStr As String, ByVal wide As Long) As String
Attribute strFit.VB_Description = "Return a string of the specified width with the supplied string left-justified in it. The returned string will always end with a space"
' Return a string of the specified width with the supplied string left-justified
' in it. The returned string will always end with a space. The supplied string
' will be padded with blanks to get the requested width, or truncated to fit.
If wide < 2& Then wide = 2&  ' silently enforce sanity
Dim res As String
res = Left$(theStr, wide - 1&) & " "
If Len(res) < wide Then res = res + String$(wide - Len(res), " ")
strFit = res
End Function

'===============================================================================
Public Function toStr( _
  ByRef thing As Variant) _
As String
Attribute toStr.VB_Description = "Convert the input Variant argument to a string if that is at all possible. Arrays (up to 3D) are returned in curly brackets as {val,val,...}"
' Convert the input Variant argument to a string if that is at all possible.
' Arrays (up to 3D) are returned in curly brackets as {val,val,...}.
Const ID_c As String = File_c & "toStr Function"

Dim vt As Integer
' note: for an object with a default property, VarType returns the type of the
' default property, not vbObject - use IsObject to be sure
vt = VarType(thing)

Dim tempStr As String
If Not IsArray(thing) Then  ' input argument is a scalar
  ' handle some special cases that don't work with CStr
  If vbEmpty = vt Then  ' VarType 0
    tempStr = "Empty"
  ElseIf vbNull = vt Then  ' VarType 1
    tempStr = "Null"
  ElseIf IsObject(thing) Then  ' works even if object has default property
    If 0& <> ObjPtr(thing) Then
      ' it's an object; report its class name and its address
      tempStr = _
        TypeName(thing) & "@" & Right("0000000" & Hex(ObjPtr(thing)), 8&)
    Else
      tempStr = "Nothing"  ' null pointer
    End If
  Else  ' try to use built-in conversion
    tempStr = CStr(thing)
    ' change nulls to spaces in case it is an unassigned fixed-length string
    tempStr = Replace(tempStr, vbNullChar, " ")
  End If
  ' now clean up some CStr funny business
  If "Error 448" = tempStr Then
    tempStr = "Missing"  ' it's a missing argument
  ElseIf "Error" = Left$(tempStr, 5&) Then ' some Error; get descriptive text
    Dim desc As String
    desc = Error(Val(Mid$(tempStr, 7&)))  ' peel off error number & get text
    If "Application-defined" <> Left$(desc, 19&) Then  ' VB has useful text
      tempStr = desc  ' use descriptive text instead of "Error NNN"
    End If
  ElseIf 0& = Len(tempStr) Then  ' the result was a null string
    tempStr = "NullString"
  End If
Else  ' input argument is an array
  tempStr = "{"
  Dim ja As Long, jaLo As Long, jaHi As Long
  Dim jb As Long, jbLo As Long, jbHi As Long
  Dim jc As Long, jcLo As Long, jcHi As Long
  Dim jdLo As Long
  jaLo = LBound(thing)
  jaHi = UBound(thing)
  On Error Resume Next
  jbLo = LBound(thing, 2&)  ' try to get lower index of second dimension
  If 0& <> Err.Number Then  ' this is a 1D array
    On Error GoTo 0
    For ja = jaLo To jaHi
      tempStr = tempStr & toStr(thing(ja))  ' recursive call
      If ja < jaHi Then
        tempStr = tempStr & ","
      Else
        tempStr = tempStr & "}"
      End If
    Next ja
  Else  ' this is a 2D or more array
    jbHi = UBound(thing, 2&)
    jcLo = LBound(thing, 3&)  ' try to get lower index of third dimension
    If 0& <> Err.Number Then  ' this is a 2D array
      On Error GoTo 0
      For ja = jaLo To jaHi
        tempStr = tempStr & "{"
        For jb = jbLo To jbHi
          tempStr = tempStr & toStr(thing(ja, jb))  ' recursive call
          If jb < jbHi Then
            tempStr = tempStr & ","
          Else
            tempStr = tempStr & "}"
          End If
        Next jb
        If ja < jaHi Then
          tempStr = tempStr & ","
        Else
          tempStr = tempStr & "}"
        End If
      Next ja
    Else  ' this is a 3D or more array
      jcHi = UBound(thing, 3&)
      jdLo = UBound(thing, 4&)  ' try to get lower index of fourth dimension
      If 0& <> Err.Number Then  ' this is a 3D array
        On Error GoTo 0
        For ja = jaLo To jaHi
          tempStr = tempStr & "{"
          For jb = jbLo To jbHi
            tempStr = tempStr & "{"
            For jc = jcLo To jcHi
              tempStr = tempStr & toStr(thing(ja, jb, jc))  ' recursive call
              If jc < jcHi Then
                tempStr = tempStr & ","
              Else
                tempStr = tempStr & "}"
              End If
            Next jc
            If jb < jbHi Then
              tempStr = tempStr & ","
            Else
              tempStr = tempStr & "}"
            End If
          Next jb
          If ja < jaHi Then
            tempStr = tempStr & ","
          Else
            tempStr = tempStr & "}"
          End If
        Next ja
      Else  ' this is a 4D or more array
        On Error GoTo 0
        Const BadArg_c As Long = 5&  ' Invalid procedure call or argument
        Err.Raise BadArg_c, ID_c, _
          Error(BadArg_c) & vbLf & _
          "Cannot handle array with 4 or more dimensions" & vbLf & _
          "Problem in " & ID_c
      End If
    End If
  End If
End If
toStr = tempStr
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


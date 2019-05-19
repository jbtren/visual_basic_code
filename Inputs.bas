Attribute VB_Name = "Inputs"
Attribute VB_Description = "Exports routines useful in parsing input strings. Devised and coded by John Trenholme."
'
'###############################################################################
'#
'# VBA Module file "Inputs.bas"
'#
'# VB routines useful in parsing input strings.
'#
'# Started 2013-04-22 by John Trenholme
'#
'# Exports the routines:
'#   Function InputsVersion
'#   Function unQuoteText
'#
'###############################################################################

Option Base 0          ' array base value when not specified     - default
Option Compare Binary  ' string comparison based on Asc(char)    - default
Option Explicit        ' force explicit declaration of variables - not default
'Option Private Module  ' no effect in VB6; globals project-only in VBA

' Module-global Const values (convention: start with upper-case; suffix "_c")

Private Const Version_c As String = "2013-04-22"
Private Const File_c As String = "Inputs[" & Version_c & "]."

Private Const BadArg_c As Long = 5&  ' = "Invalid procedure call or argument"

'===============================================================================
Public Function InputsVersion(Optional ByVal trigger As Variant) As String
Attribute InputsVersion.VB_Description = "Date of the latest revision to this code, as a string with format 'YYYY-MM-DD'"
' Date of the latest revision to this code, as a string with format "YYYY-MM-DD"
InputsVersion = Version_c
End Function

'===============================================================================
Public Function unQuoteText( _
  ByRef arg As String, _
  Optional ByVal cut As Boolean = False) _
As String
' Ignore leading and trailing blanks in "arg", and return the first quote-
' delimited item in "arg", or "arg" up to first blank if it contains no quotes.
' Double quotes in arg are replaced by single quotes, and don't count as
' delimiters. If optional "cut" is true, remove item from "arg".
' It is an error if there is non-blank text in "arg" before the first quote.
' Example: [space][space]"This is ""strange"" text" he said.[space]
' returns the String <This is "strange" text> (without the <>), and if
' "cut" is true trims the input down to <he said.[space]>.

' set up error handling
Const ID_c As String = File_c & "unQuoteText"  ' name of this file + routine
Static calls_s As Double  ' number of times this routine has been called
calls_s = calls_s + 1#    ' stops adding at 9,007,199,254,740,992 calls
If RelVer_C Then On Error GoTo ErrorHandler

Dim text As String
text = Trim$(arg)  ' no blanks on either end
Dim qLoc As Long
qLoc = InStr(text, """")
If qLoc = 0& Then  ' no quote anywhere
  Dim bLoc As Long
  bLoc = InStr(text, " ")
  If bLoc > 0& Then
    unQuoteText = Left$(text, bLoc - 1&)  ' trim at first blank
  Else
    unQuoteText = text  ' no quote & no blank; return entire arg
  End If
ElseIf qLoc = 1& Then ' there is a leading quote
  text = Mid$(text, 2&)  ' strip the quote
  qLoc = 0&
  Do
    qLoc = InStr(qLoc + 1&, text, """") ' find first closing quote
    If qLoc = 0& Then  ' no closing quote; punt
      unQuoteText = text
      Exit Do
    ElseIf qLoc = Len(text) Then  ' closing quote is at end
      unQuoteText = Left$(text, qLoc - 1&)
      Exit Do
    ElseIf Mid$(text, qLoc + 1&, 1&) <> """" Then ' no double quote in text
      unQuoteText = Left$(text, qLoc - 1&)
      Exit Do
    Else  ' double quote seen; replace by single quote & continue scan
      text = Left$(text, qLoc - 1&) & Mid$(text, qLoc + 1&)
    End If
  Loop
Else
  Err.Raise 5&, ID_c, "Text before first quote in:" & vbLf & arg
End If
Exit Function

ErrorHandler:  '----------------------------------------------------------------
' save properties of Err object & Erl; "On Error GoTo 0" will erase them
Dim errNum As Long, errDsc As String
errNum = Err.Number
' handle error if one exists; if no error, fall through to Exit
If errNum <> 0& Then  ' there was an error
  errDsc = Err.Description & vbLf
  On Error GoTo 0  ' avoid recursion
  If InStr(errDsc, "Problem in") > 0& Then  ' error below here; add traceback
    errDsc = errDsc & "Called from " & ID_c & " call " & calls_s
  Else  ' error was in this routine (could add more info here)
    errDsc = errDsc & "Problem in " & ID_c & " call " & calls_s
  End If
  ' if line number is available, add it on
  Dim errLin As Long: errLin = Erl  ' unsupported Erl holds line number of error
  If errLin <> 0& Then errDsc = errDsc & ", line " & errLin
  Err.Raise errNum, ID_c, errDsc ' send on up the call chain
  Resume  ' set Next Statement here & hit F8 to go to error line
End If
End Function  '-----------------------------------------------------------------

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


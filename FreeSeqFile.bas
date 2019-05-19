Attribute VB_Name = "FreeSeqFile"
'
'###############################################################################
'# Visual Basic 6 source file "FreeSeqFile.bas"
'#
'# Devised and coded by John Trenholme - initial version 30 Nov 2003
'###############################################################################

Option Explicit

Private Const c_Version As String = "2007-06-17"

'===============================================================================
Public Function incrementLastNumber(ByRef arg As String) As String
' Given an input string, this routine searches backward through it looking for
' a numeric digit. If one is found, the sequence of contiguous digits ending
' with that digit is incremented by one, giving a new string with the last
' number in it bumped by one. The new string is returned as the function value;
' the input string is unchanged.
'
' When the number reaches 99...99, it is silently rolled over to 00...00.
'
' If no number is found, a copy of the input string is silently returned.
'
' This routine is useful, for example, to
' increment to the next file in a numbered sequence of files:
'   Chapter_27_003.bmp -> Chapter_27_004.bmp
'   Chapter_27_999.bmp -> Chapter_27_000.bmp  (may not be what you want!)
'
' To avoid premature rollover, use more digits in the sequence number.
' Most likely, you will start at xxx000xx.xx and pre- or post-increment to
' xxx001xx.xx, and so forth. If for some reason you want or need to pre-
' increment and want to start at xxx000xx.xx, set the initial string to
' xxx999xx.xx and bump that to xxx000xx.xx.
'
'     devised and coded by John Trenholme - version of 4 Sep 2002

Dim ch As String
Dim foundDigit As Boolean
Dim j As Long
Dim str As String

str = arg                                  ' working copy of input string
foundDigit = False
For j = Len(str) To 1 Step -1              ' step backwards through string
  ch = Mid$(str, j, 1)                     ' get new character
  If InStr("0123456789", ch) > 0& Then     ' character is a numeric digit
    foundDigit = True                      ' remember that we found a number
    If ch = "9" Then
      Mid$(str, j, 1) = "0"                ' roll digit over from 9 to 0; carry
    Else
      Mid$(str, j, 1) = Chr$(Asc(ch) + 1&) ' increment this digit
      Exit For                             ' quit, since add did not carry
    End If
  Else                                     ' character was not a numeric digit
    If foundDigit Then Exit For            ' after numeric -> end of number
  End If
Next j                                     ' no digit yet or carry; step left
incrementLastNumber = str
End Function

'===============================================================================
Public Function freeSequencedFile(ByVal baseName As String) As String
' Given a base name of the form xNN..Nx, where NN..N is the last sequence
' of contiguous digits in the base name and "x" represents other characters,
' this routine finds and returns the name of the first unused file of the
' supplied form. For example, if you supply "MyFile000.txt" and files from
' "MyFile000.txt" to "MyFile123.txt" exist on disk, the routine would return
' "MyFile124.txt" and you could then safely open and write to that file.
'
' If there are no digits in the name you supply, it will be returned unchanged
' and if you use the name it will overwrite any existing file with that name.
'
' If all file names of the supplied form exist on disk, the routine will raise
' error 67 ("Too many files") with an explanatory message.
'
' The file names are checked in the current drive and directory unless the base
' name is prefixed with drive, directory or both.

Dim freeName As String
freeName = baseName
' see if base file exists on disk and there are sequence digits
If (Dir$(freeName) <> "") And (incrementLastNumber(freeName) <> freeName) Then
  Do
    freeName = incrementLastNumber(freeName)
    If Len(Dir$(freeName)) = 0& Then Exit Do  ' Dir$ = "" if file does not exist
  Loop Until freeName = baseName  ' if we exit here, all names are in use
  If freeName = baseName Then  ' all sequenced file names are in use
    Err.Raise 67, "Function freeSequencedFile", _
      "ERROR in freeSequencedFile: All sequenced file names based on" & vbLf & _
      """" & baseName & """" & vbLf & _
      "are in use. Can't proceed without overwriting a file." & vbLf & vbLf & _
      "FIX: please delete or move one or more files and try again."
  End If
End If
freeSequencedFile = freeName
End Function

'===============================================================================
Public Function freeSeqVersion() As String
freeSeqVersion = c_Version
End Function

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


Attribute VB_Name = "RotateText_"
Attribute VB_Description = "Module holding the RotateText() subroutine. Coded by John Trenholme."
'
'###############################################################################
'#        _____          _           _          _______            _.
'#       |  __ \        | |         | |        |__   __|          | |
'#       | |__) |  ___  | |_   __ _ | |_   ___    | |   ___ __  __| |_.
'#       |  _  /  / _ \ | __| / _` || __| / _ \   | |  / _ \\ \/ /| __|
'#       | | \ \ | (_) || |_ | (_| || |_ |  __/   | | |  __/ >  < | |_.
'#       |_|  \_\ \___/  \__| \__,_| \__| \___|   |_|  \___|/_/\_\ \__|
'#
'# Visual Basic file "RotateText.bas"
'#
'# Print rotated text on anything with an hDC - John Trenholme - 20 Dec 2005
'#
'###############################################################################

Option Base 0
Option Compare Binary
Option Explicit

Public Const RotateTextVersion As String = "2006-01-06"

'*******************************************************************************
' Win32 API things needed for font control

Private Const ANSI_CHARSET As Long = 0&
Private Const CLIP_DEFAULT_PRECIS As Long = 0&
Private Const DEFAULT_PITCH As Long = 0&
Private Const FF_DONTCARE As Long = 0&
Private Const LF_FACESIZE As Long = 32&
Private Const OEM_CHARSET As Long = 255&
Private Const OUT_DEFAULT_PRECIS As Long = 0&
Private Const OUT_TT_ONLY_PRECIS As Long = 7&
Private Const PROOF_QUALITY As Long = 2&

Private Type LogFont
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFaceName As String * LF_FACESIZE
End Type

Private Declare Function CreateFontIndirect Lib "gdi32" _
  Alias "CreateFontIndirectA" (lpLogFont As LogFont) As Long
Private Declare Function DeleteObject Lib "gdi32" _
  (ByVal hObject As Long) As Long
Declare Function SelectObject Lib "gdi32" _
  (ByVal hDC As Long, ByVal hObject As Long) As Long

'===============================================================================
Public Sub RotateText(ob As Object, _
                      text As String, _
                      ByRef angle As Single, _
                      Optional fontRef As Variant)
Attribute RotateText.VB_Description = "Prints text on anything with an hDC, at a specified angle."
' Prints text on anything with an hDC, at a specified angle.

Dim lf As LogFont

' try to get a handle to the device context of the supplied Object
Dim obDC As Long
On Error Resume Next
obDC = ob.hDC
If obDC = 0& Or Err.Number <> 0& Then
  On Error GoTo 0
  Err.Raise 438&, "RotateText", _
    "in Sub RotateText (file RotateText.bas)" & vbLf & _
    "Cannot get hDC for supplied Object" & vbLf & _
    "Object is of type """ & TypeName(ob) & """"
End If
On Error GoTo 0

' get font properties for use in LogFont
Dim sf As StdFont
If IsMissing(fontRef) Then
  ' no reference StdFont supplied; try to use ob.Font
  On Error Resume Next
  Set sf = ob.Font
  If Err.Number <> 0& Then
    On Error GoTo 0
    Err.Raise 438&, "RotateText", _
      "in Sub RotateText (file RotateText.bas)" & vbLf & _
      "No StdFont was supplied as a reference font, and" & vbLf & _
      "cannot get Font from supplied Object" & vbLf & _
      "Object is of type """ & TypeName(ob) & """"
  End If
  On Error GoTo 0
Else
  If TypeOf fontRef Is StdFont Then
    ' reference StdFont supplied
    Set sf = fontRef
  Else
    Err.Raise 438&, "RotateText", _
      "in Sub RotateText (file RotateText.bas)" & vbLf & _
      "Supplied argument ""fontRef"" is not a StdFont" & vbLf & _
      "It is of type """ & TypeName(fontRef) & """"
  End If
End If

' normalize the angle
angle = angle Mod 360!
If angle < 0! Then angle = angle + 360!  ' because Mod of negative is negative

' set up the LogFont
With lf
  .lfHeight = -sf.Size * 20! / Screen.TwipsPerPixelY  ' 20 = twips per point
  .lfWidth = 0&
  .lfEscapement = angle * 10!
  .lfOrientation = .lfEscapement
  .lfWeight = sf.Weight
  .lfItalic = sf.Italic
  .lfUnderline = sf.Underline
  .lfStrikeOut = sf.Strikethrough
  .lfClipPrecision = CLIP_DEFAULT_PRECIS
  .lfQuality = PROOF_QUALITY
  .lfPitchAndFamily = DEFAULT_PITCH Or FF_DONTCARE
  .lfFaceName = sf.Name & vbNullChar
  ' OEM fonts can't rotate; must force to ANSI if angle is not zero
  .lfCharSet = sf.Charset
  If .lfCharSet = OEM_CHARSET Then
    If .lfEscapement <> 0& Then
      .lfCharSet = ANSI_CHARSET
    End If
  End If
  ' only TrueType fonts can rotate; must specify TT-only if angle is not zero
  If .lfEscapement <> 0& Then
    .lfOutPrecision = OUT_TT_ONLY_PRECIS
  Else
    .lfOutPrecision = OUT_DEFAULT_PRECIS
  End If
  ' save ratio of sizes for reset of CurrentX and CurrentY (if needed)
  If lf.lfEscapement <> 0& Then
    Dim scal As Single
    scal = -.lfHeight / sf.Size
  End If
End With
Set sf = Nothing

' create the rotated font
Dim newFont As Long
newFont = CreateFontIndirect(lf)

' out with the old font; in with the new font
Dim oldFont As Long
oldFont = SelectObject(obDC, newFont)

' save CurrentX and CurrentY, they will be updated for angle not zero
Dim gotCurrentX As Boolean
If lf.lfEscapement <> 0& Then
  Dim oldX As Single
  Dim oldY As Single
  On Error Resume Next  ' will get an error if no such Property
  oldX = ob.CurrentX
  gotCurrentX = (Err.Number = 0&)
  On Error GoTo 0
  If gotCurrentX Then oldY = ob.CurrentY
End If

' print the text
ob.Print text;

' fix up CurrentX and CurrentY if they exist and angle was not zero
' note: because of font-size jitter, this is not exact - sorry!
If gotCurrentX And (lf.lfEscapement <> 0&) Then
  Dim dx As Single
  Dim dy As Single
  dx = scal * (ob.CurrentX - oldX)
  dy = scal * (ob.CurrentY - oldY)
  Dim ca As Single
  Dim sa As Single
  Const degToRad As Single = 3.141593! / 180!
  ca = Cos(angle * degToRad)
  sa = Sin(angle * degToRad)
  ob.CurrentX = oldX + dx * ca - dy * sa
  ob.CurrentY = oldY + dx * sa + dy * ca
End If

' put back old font, and delete new one (resource leak if not done!)
DeleteObject SelectObject(obDC, oldFont)

End Sub

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


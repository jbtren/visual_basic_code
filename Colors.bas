Attribute VB_Name = "Colors"
'#
'###############################################################################
'#
'#       ######    #######   ##         #######   ########    ######
'#      ##    ##  ##     ##  ##        ##     ##  ##     ##  ##    ##
'#      ##        ##     ##  ##        ##     ##  ##     ##  ##
'#      ##        ##     ##  ##        ##     ##  ########    ######
'#      ##        ##     ##  ##        ##     ##  ##   ##          ##
'#      ##    ##  ##     ##  ##        ##     ##  ##    ##   ##    ##
'#       ######    #######   ########   #######   ##     ##   ######
'#
'# Visual Basic for Applications (VBA) & VB6 Module file "Colors.bas"
'#
'# by John Trenholme - Started 2006-09-19
'#
'# Exports:
'#   Function colorDark
'#   Function colorPale
'#   Function ColorRGB
'#   Function ColorSequenceCount
'#   Function ColorVersion
'#   Function Jumpy
'#
'#   Enum ClrPalette
'#
'###############################################################################

Option Base 0          ' array base value, when not specified - the default
Option Compare Binary  ' string comparison based on Asc(char) - the default
Option Explicit        ' forces explicit variable declaration - changes default

Private Const Version_c As String = "2013-10-29"

' Names for the palettes; Enums are Longs
Public Enum ClrPalette
  [Cool to Warm] = 0&
  [Gray Scale]
  [Red Scale]
  [Green Scale]
  [Blue Scale]
  [Cyan Scale]
  [Magenta Scale]
  [Yellow Scale]
  [Blue -> Red]
  [Green -> Red]
  [Blue -> Green]
  [Blue -> Green -> Red]
  [Teal -> Yellow]
  [Dark Wheel]
  [Bright Wheel]
  [Cold to Hot]
  [7 Bands]
  [20 Bands]
  [10 Ramps]
  [Mystery Palette]
  [Another Wheel]
  [NIST Phase]
End Enum

' update this value manually to correspond to the actual count (including 0)
Private Const Sequences_c As Integer = [NIST Phase] + 1&

'===============================================================================
Public Function colorDark( _
  ByVal baseColor As Long, _
  ByVal darkness As Double) _
As Long
' Accepts a standard VB color in &H00BBGGRR format, and moves it toward totally
' black by the fraction 'darkness'. Normally 0 <= darkness <= 1.
Dim r As Double, g As Double, b As Double
r = baseColor And &HFF&
g = (baseColor And &HFF00&) / &H100&
b = (baseColor And &HFF0000) / &H10000
r = (1# - darkness) * r
g = (1# - darkness) * g
b = (1# - darkness) * b
colorDark = RGB(r, g, b)
End Function

'===============================================================================
Public Function colorPale( _
  ByVal baseColor As Long, _
  ByVal paleness As Double) _
As Long
' Accepts a standard VB color in &H00BBGGRR format, and moves it toward totally
' white by the fraction 'paleness'. Normally 0 <= paleness <= 1.
Dim r As Double, g As Double, b As Double
r = baseColor And &HFF&
g = (baseColor And &HFF00&) / &H100&
b = (baseColor And &HFF0000) / &H10000
r = (1# - paleness) * r + paleness * 255#
g = (1# - paleness) * g + paleness * 255#
b = (1# - paleness) * b + paleness * 255#
colorPale = RGB(r, g, b)
End Function

'===============================================================================
Public Function colorRGB( _
  ByVal turn As Double, _
  Optional ByVal whichPalette As ClrPalette = [Cool to Warm], _
  Optional ByVal stripeCount As Double = 0#, _
  Optional ByVal stripeWidth As Double = 0.25) _
As Long
' Traverses a set of colors once for each unit change in "turn" and returns a
' color value as a Long (packed in the standard VB format &H00BBGGRR).  The
' palette of colors to be used is selected by 'whichPalette'. Adds dark stripes
' if 'stripeCount' > 0, with a relative width set by 'stripeWidth'.
' If you want the cycle to span the x range from A to B, set the first argument
' to turn = (x - A) / (B - A). Note that stepping 'turn' by the golden section
' related value 2 - Phi = 0.38196601125 will give a nice sequence of distinct
' colors; this can be further extended by adding darkness or paleness.
Const Pi_c As Double = 3.1415926 + 5.35897932E-08
Const TwoPi_c As Double = 2# * Pi_c

' reduce input value 'turn' Mod 1, so that 0 < color <= 1
' this formula "fails" if -1.1E-16 < turn < 0, giving 1, but that's OK here
Dim c As Double
c = turn - Int(turn)
If turn <> 0# Then
  If c = 0# Then c = 1#  ' non-0 integer values map to 1, not 0
End If

Dim red As Double, green As Double, blue As Double

' select the color sequence to be used
Select Case whichPalette
  Case 1&  ' gray scale
    red = li(c, 40#, 240#)
    green = red
    blue = red
  Case 2&  ' red scale
    red = 256#
    green = li(c, 0#, 240#)
    blue = green
  Case 3&  ' green scale
    red = li(c, 0#, 240#)
    green = li(c, 100#, 256#)
    blue = red
  Case 4&  ' blue scale
    red = li(c, 0#, 240#)
    green = red
    blue = li(c, 165#, 256#)
  Case 5&  ' cyan scale
    green = 256# * Sqr(c)
    blue = green
    red = 0.3 * green
  Case 6&  ' magenta scale
    red = 256# * Sqr(c)
    green = 0.3 * red
    blue = red
  Case 7&  ' yellow scale
    red = 256# * Sqr(c)
    green = 0.9 * red
    blue = 0.2 * red
  Case 8&  ' blue -> red
    red = 256# * Sqr(c)
    green = 0#
    blue = 256# * Sqr(1# - c)
  Case 9&  ' green -> red
    red = 256# * Sqr(c)
    green = 256# * Sqr(1# - c)
    blue = 0#
  Case 10&  ' blue -> green
    red = 0#
    green = 256# * Sqr(c)
    blue = 256# * Sqr(1# - c)
  Case 11&  ' blue -> green -> red
    c = 2# * c
    If c < 1# Then
      red = 0#
      green = 256# * Sqr(c)
      blue = 256# * Sqr(1# - c)
    Else
      c = c - 1#
      red = 256# * Sqr(c)
      green = 256# * Sqr(1# - c)
      blue = 0#
    End If
  Case 12&  ' teal to yellow
    red = li(c, 0#, 256#)
    green = li(c, 70#, 220#)
    blue = li(c, 100#, 60#)
  Case 13&, 14&
    ' dark wheel; bright wheel
    c = 6# * c  ' put in range; scale for 6 sections
    If c < 1# Then
      red = 1#: green = Sqr(c): blue = 0# ' from red to yellow
    ElseIf c < 2# Then
      red = Sqr(2# - c): green = 1#: blue = 0#  ' from yellow to green
    ElseIf c < 3# Then
      red = 0#: green = 1#: blue = Sqr(c - 2#)  ' from green to cyan
    ElseIf c < 4# Then
      red = 0#: green = Sqr(4# - c): blue = 1#  ' from cyan to blue
    ElseIf c < 5# Then
      red = Sqr(c - 4#): green = 0#: blue = 1#  ' from blue to magenta
    ElseIf c >= 5# Then
      red = 1#: green = 0#: blue = Sqr(6# - c)  ' from magenta to red
    End If
    If whichPalette = 13& Then
      red = 160# * red
      green = 160# * green
      blue = 160# * blue
    Else
      red = 256# * red
      green = 256# * green
      blue = 256# * blue
    End If
  Case 15&  ' cold to hot
    c = 6# * c
    If c < 1# Then
      red = li(c, 72#, 0#)
      green = li(c, 72#, 0#)
      blue = li(c, 72#, 170#)
    ElseIf c < 2# Then
      red = li(c - 1#, 0#, 120#)
      green = 0#
      blue = li(c - 1#, 170#, 135#)
    ElseIf c < 3# Then
      red = li(c - 2#, 120#, 256#)
      green = 0#
      blue = li(c - 2#, 135#, 0#)
    ElseIf c < 4# Then
      red = li(c - 3#, 256#, 244#)
      green = li(c - 3#, 0#, 170#)
      blue = 0#
    ElseIf c < 5# Then
      red = li(c - 4#, 244#, 256#)
      green = li(c - 4#, 170#, 230#)
      blue = 0#
    Else
      red = li(c - 5#, 256#, 256#)
      green = li(c - 5#, 230#, 256#)
      blue = li(c - 5#, 0#, 256#)
    End If
  Case 16&  ' 7 bands
    red = 128# * (1# - Sin(c * TwoPi_c))
    green = 128# * (1# - Cos(2# * c * TwoPi_c))
    blue = 128# * (1# - Cos(3# * c * TwoPi_c))
  Case 17&  ' 20 bands
    red = 128# * (1# - Sin(c * TwoPi_c))
    green = 128# * (1# - Cos(3# * c * TwoPi_c))
    blue = 128# * (1# - Cos(10# * c * TwoPi_c))
  Case 18&  ' 10 ramps
    red = 256# * Sqr(c)
    green = 256# * Sqr(5# * c - Int(5# * c))
    blue = 256# * Sqr(10# * c - Int(10# * c))
  Case 19&  ' mystery palette
    c = c * 256#
    If c < 51# Then
      c = c / 51#
      red = li(c, 0#, 48#)
      green = li(c, 0#, 137#)
      blue = li(c, 164#, 163#)
    ElseIf c < 85# Then
      c = (c - 51#) / (85# - 51#)
      red = li(c, 48#, 117#)
      green = li(c, 137#, 158#)
      blue = li(c, 163#, 54#)
    ElseIf c < 153# Then
      c = (c - 85#) / (153# - 85#)
      red = li(c, 117#, 213#)
      green = li(c, 158#, 207#)
      blue = li(c, 54#, 0#)
    ElseIf c < 204# Then
      c = (c - 153#) / (204# - 153#)
      red = li(c, 213#, 216#)
      green = li(c, 207#, 0#)
      blue = li(c, 0#, 221#)
    Else
      c = (c - 204#) / (255# - 204#)
      red = li(c, 216#, 179#)
      green = li(c, 0#, 228#)
      blue = li(c, 221#, 256#)
    End If
  Case 20&  ' a simple color wheel
    Const Tweak_c As Double = 0.35
    red = 200# * (1# - Tweak_c * (1# - Cos(turn * TwoPi_c)))
    green = 180# * (1# - Tweak_c * (1# - Cos((turn + 0.3333) * TwoPi_c)))
    blue = 220# * (1# - Tweak_c * (1# - Cos((turn + 0.6667) * TwoPi_c)))
  Case 21&  ' NIST phase colors, use turn = angleRads / TwoPi
    c = 4# * c
    If c < 1# Then  ' 0° = red
      red = 256#
      green = 235# * c
      blue = 0#
    ElseIf c < 2# Then  ' 90° = yellow
      c = c - 1#
      red = 256# * (1# - c)
      green = 235#
      blue = 256# * c
    ElseIf c < 3# Then  ' 180° = cyan
      c = c - 2#
      red = 0#
      green = 235# * (1# - c)
      blue = 256#
    Else  ' 270° = blue
      c = c - 3#
      red = 256# * Sqr(c)
      green = 0#
      blue = 256# * Sqr(1# - c)
    End If
  Case Else  ' use the default palette: gray-blue-green-red-magenta-cyan-yellow
    c = 6# * c
    If c < 1# Then
      red = li(c, 70#, 0#)
      green = li(c, 70#, 0#)
      blue = li(c, 70#, 256#)
    ElseIf c < 2# Then
      c = c - 1#
      red = 0#
      green = li(c, 0#, 200#)
      blue = li(c, 256#, 0#)
    ElseIf c < 3# Then
      c = c - 2#
      red = li(c, 0#, 256#)
      green = li(c, 200#, 0#)
      blue = 0#
    ElseIf c < 4# Then
      c = c - 3#
      red = 256#
      green = li(c, 0#, 30#)
      blue = li(c, 0#, 245#)
    ElseIf c < 5# Then
      c = c - 4#
      red = li(c, 256#, 60#)
      green = li(c, 30#, 240#)
      blue = li(c, 245#, 256#)
    Else
      c = c - 5#
      red = li(c, 60#, 256#)
      green = li(c, 240#, 256#)
      blue = li(c, 256#, 110#)
    End If
End Select

' put on stripes, if any
If stripeCount > 0# Then
  Dim stripe As Double
  ' avoid silly values
  If stripeWidth < 0.001 Then
    stripeWidth = 0.001
  Else
    If stripeWidth > 1# Then stripeWidth = 1#
  End If
  ' make up the stripe multiplier
  stripe = Abs(Cos(turn * Pi_c * stripeCount) / stripeWidth)
  If stripe > 1# Then stripe = 1#
  ' stripe 'em
  red = stripe * red
  green = stripe * green
  blue = stripe * blue
End If

' put values in range (arguments to RGB get CInt handling, then must be >= 0)
If red < 0# Then red = 0# Else If red > 256# Then red = 256#
If green < 0# Then green = 0# Else If green > 256# Then green = 256#
If blue < 0# Then blue = 0# Else If blue > 256# Then blue = 256#

' offset values by -0.5 because CInt is used on inputs to RGB
' note that CInt(256-0.5) is 256, but RGB args above 255 are set to 255, so OK
colorRGB = RGB(red - 0.5, green - 0.5, blue - 0.5)
End Function

'===============================================================================
Public Function colorSequenceCount() As Long
' This is the count of basic palettes available, not counting striping or
' darkness or paleness. Their index values are 0 to colorSequenceCount() - 1.
colorSequenceCount = Sequences_c
End Function

'===============================================================================
Public Function colorVersion(Optional ByVal trigger As Variant) As String
' The date of the latest revision to this module as a string in the format
' 'YYYY-MM-DD' such as 2024-06-18. It's a function so Excel etc. can use it.
colorVersion = Version_c
End Function

'===============================================================================
Public Function Jumpy(ByVal j As Long, ByVal k As Long) As Double
' Given a smoothly increasing sequence of values from 0 to k-1, or from
' 1 to k, output a sequence of Double values that jumps back and forth by about
' 1/2 but still includes all possible values of j/k.
' Example:  j = 1 2 3 4 5 6 7  k = 7  yields 4/7 1/7 5/7 2/7 6/7 3/7 7/7
' Useful for making successive colors jump back and forth in a palette by using
' (e.g.) ColorRGB(Jumpy(j, 12&)) for 12 colors, with j = 0, 1, ..., 11.
' As an alternative, consider the use of multiples of  2 - Phi = 0.38196601125.
If j <= 0 Then
  Jumpy = 0#
ElseIf j >= k Then
  Jumpy = 1#
Else
  Dim n As Integer
  If (k And 1&) = 0& Then  ' k is even
    If (j And 1&) = 0& Then  ' j is even
      n = (j + k) \ 2&
    Else
      n = (j + 1&) \ 2&
    End If
  Else  ' k is odd
    If (j And 1&) = 0& Then  ' j is even
      n = j \ 2&
    Else
      n = (j + k) \ 2&
    End If
  End If
  Jumpy = n / k
End If
End Function

'-------------------------------------------------------------------------------
Private Function li( _
  ByVal fraction As Double, _
  ByVal first As Double, _
  ByVal last As Double) _
As Double
' Linear interpolation from one value to another. Usually 0 <= fraction <= 1.
li = (1# - fraction) * first + fraction * last
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


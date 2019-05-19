Attribute VB_Name = "ColorsMod"
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
'# Visual Basic for Applications (VBA) Module file "ColorsMod.bas"
'#
'# Routines to handle Excel colors.
'#
'# by John Trenholme - Started 2006-09-19
'#
'# Exports the routines:
'#   Function Colors
'#   Function ColorsSequenceCount
'#   Function ColorsVersion
'#
'###############################################################################

Option Base 0
Option Compare Binary
Option Explicit

Const Version_c As String = "2006-09-25"

' update this value manually to correspond to the actual count (including 0)
Const Sequences_c As Integer = 20

'===============================================================================
Public Function Colors( _
  ByVal turn As Double, _
  Optional ByVal whichPalette As Integer = 0, _
  Optional ByVal stripeCount As Double = 0#, _
  Optional ByVal stripeWidth As Double = 0.25) _
As Long
' Traverses a set of colors once for each unit change in "turn" and returns a
' color value as a Long (packed &H00BBGGRR).  The set of colors to be used is
' selected by 'whichPalette'. Adds dark stripes if 'stripeCount' > 0, with a
' relative width set by 'stripeWidth'.
Const Pi_c As Double = 3.14159265358979
Const TwoPi_c As Double = 2# * Pi_c
Dim c As Double
Dim k As Long
c = turn - Int(turn)  ' fractional part; jumps at integers; OK for negative
Dim red As Double, green As Double, blue As Double
' select the color sequence to be used
Select Case whichPalette
  Case 1  ' gray scale
    red = li(c, 40, 240)
    green = red
    blue = red
  Case 2  ' red scale
    red = 255
    green = li(c, 0, 240)
    blue = green
  Case 3  ' green scale
    red = li(c, 0, 240)
    green = li(c, 100, 256)
    blue = red
  Case 4  ' blue scale
    red = li(c, 0, 240)
    green = red
    blue = li(c, 165, 256)
  Case 5  ' cyan scale
    green = 256# * Sqr(c)
    blue = green
    red = 0.3 * green
  Case 6  ' magenta scale
    red = 256# * Sqr(c)
    green = 0.3 * red
    blue = red
  Case 7  ' yellow scale
    red = 256# * Sqr(c)
    green = 0.9 * red
    blue = 0.2 * red
  Case 8  ' blue -> red
    red = 256# * Sqr(c)
    green = 0#
    blue = 256# * Sqr(1# - c)
  Case 9  ' green -> red
    red = 256# * Sqr(c)
    green = 256# * Sqr(1# - c)
    blue = 0#
  Case 10  ' blue -> green
    red = 0#
    green = 256# * Sqr(c)
    blue = 256# * Sqr(1# - c)
  Case 11  ' blue -> green -> red
    If c < 0.5 Then
      red = 0#
      green = 256# * Sqr(2# * c)
      blue = 256# * Sqr(1# - 2# * c)
    Else
      red = 256# * Sqr(2# * c - 1#)
      green = 256# * Sqr(2# * (1# - c))
      blue = 0#
    End If
  Case 12  ' teal to yellow
    red = li(c, 0, 256)
    green = li(c, 70, 220)
    blue = li(c, 100, 60)
  Case 13, 14
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
    If whichPalette = 13 Then
      red = 160# * red - 0.5
      green = 160# * green - 0.5
      blue = 160# * blue - 0.5
    Else
      red = 256# * red - 0.5
      green = 256# * green - 0.5
      blue = 256# * blue - 0.5
    End If
  Case 15  ' cold to hot
    c = 6# * c
    If c < 1# Then
      red = li(c, 72, 0)
      green = li(c, 72, 0)
      blue = li(c, 72, 170)
    ElseIf c < 2# Then
      red = li(c - 1#, 0, 120)
      green = li(c - 1#, 0, 0)
      blue = li(c - 1#, 170, 135)
    ElseIf c < 3# Then
      red = li(c - 2#, 120, 256)
      green = li(c - 2#, 0, 0)
      blue = li(c - 2#, 135, 0)
    ElseIf c < 4# Then
      red = li(c - 3#, 256, 244)
      green = li(c - 3#, 0, 170)
      blue = li(c - 3#, 0, 0)
    ElseIf c < 5# Then
      red = li(c - 4#, 244, 256)
      green = li(c - 4#, 170, 230)
      blue = li(c - 4#, 0, 0)
    Else
      red = li(c - 5#, 256, 256)
      green = li(c - 5#, 230, 256)
      blue = li(c - 5#, 0, 256)
    End If
  Case 16  ' 7 bands
    red = 128# * (1# - Sin(c * TwoPi_c)) - 0.5
    green = 128# * (1# - Cos(2# * c * TwoPi_c)) - 0.5
    blue = 128# * (1# - Cos(3# * c * TwoPi_c)) - 0.5
  Case 17  ' 20 bands
    red = 128# * (1# - Sin(c * TwoPi_c)) - 0.5
    green = 128# * (1# - Cos(3# * c * TwoPi_c)) - 0.5
    blue = 128# * (1# - Cos(10# * c * TwoPi_c)) - 0.5
  Case 18  ' 10 ramps
    red = 256# * Sqr(c) - 0.5
    green = 256# * Sqr(5# * c - Int(5# * c)) - 0.5
    blue = 256# * Sqr(10# * c - Int(10# * c)) - 0.5
  Case 19
    c = c * 256#
    If c < 51# Then
      c = c / 51#
      red = li(c, 0, 48)
      green = li(c, 0, 137)
      blue = li(c, 164, 163)
    ElseIf c < 85# Then
      c = (c - 51#) / (85# - 51#)
      red = li(c, 48, 117)
      green = li(c, 137, 158)
      blue = li(c, 163, 54)
    ElseIf c < 153# Then
      c = (c - 85#) / (153# - 85#)
      red = li(c, 117, 213)
      green = li(c, 158, 207)
      blue = li(c, 54, 0)
    ElseIf c < 204# Then
      c = (c - 153#) / (204# - 153#)
      red = li(c, 213, 216)
      green = li(c, 207, 0)
      blue = li(c, 0, 221)
    Else
      c = (c - 204#) / (255# - 204#)
      red = li(c, 216, 179)
      green = li(c, 0, 228)
      blue = li(c, 221, 256)
    End If
  Case Else  ' the default is a simple color wheel
    Const Tweak_c As Double = 0.35
    red = 200# * (1# - Tweak_c * (1# - Cos(turn * TwoPi_c))) - 0.5
    green = 180# * (1# - Tweak_c * (1# - Cos((turn + 0.3333) * TwoPi_c))) - 0.5
    blue = 220# * (1# - Tweak_c * (1# - Cos((turn + 0.6667) * TwoPi_c))) - 0.5
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
Const RGBmin_c As Double = -0.5
Const RGBmax_c As Double = 32767.4999999999  ' if over 255, forced to 255
If red < RGBmin_c Then
  red = RGBmin_c
Else
  If red > RGBmax_c Then
    red = RGBmax_c
  End If
End If
If green < RGBmin_c Then
  green = RGBmin_c
Else
  If green > RGBmax_c Then
    green = RGBmax_c
  End If
End If
If blue < RGBmin_c Then
  blue = RGBmin_c
Else
  If blue > RGBmax_c Then
    blue = RGBmax_c
  End If
End If
Colors = RGB(red, green, blue)
End Function

'===============================================================================
Public Function ColorsSequenceCount() As Long
' Tis is the count of basic palettes available, not counting striping.
ColorsSequenceCount = Sequences_c
End Function

'===============================================================================
Public Function ColorsVersion() As String
' The date of the latest revision to this module as a string in the format
' 'YYYY-MM-DD' such as 2004-06-18. It's a function so Excel etc. can use it.
ColorsVersion = Version_c
End Function

'-------------------------------------------------------------------------------
Private Function li( _
  ByVal fraction As Double, _
  ByVal first As Integer, _
  ByVal last As Integer) _
As Double
' Linear interpolation from one integer to another (- 0.5 is for CInt later)
li = (1# - fraction) * first + fraction * last - 0.5
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


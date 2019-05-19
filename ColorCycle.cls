VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColorCycle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'###############################################################################
'# Visual Basic 6 Class file "ColorCycle.cls"
'#
'# Supplies a number of color sequences, mapped to x Mod 1, where x is a Double.
'# Once a particular sequence is selected, input Double values will be mapped to
'# colors in the sequence. The sequence will be traversed once as the input
'# value changes from 0.0 to 1.0, and then will repeat as the input changes from
'# 1.0 to 2.0, and so forth for each unit change of the input.
'#
'# Initial version 3 May 2004 by John Trenholme.
'###############################################################################

Option Base 0
Option Compare Binary
Option Explicit

Private Const c_version As String = "2004-06-11"  ' update when Class changed
Private Const c_SequenceCount As Integer = 11
Private Const c_Pi As Double = 3.14159265358979
Private Const c_TwoPi As Double = 2# * c_Pi

' Local values of Properties.
Private m_Sequence As Integer
Private m_SequenceNames() As String

'++++++++ Event ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub Class_Initialize()
m_Sequence = 0
ReDim m_SequenceNames(0 To c_SequenceCount - 1)
m_SequenceNames(0) = "Gray scale"
m_SequenceNames(1) = "Dark colors"
m_SequenceNames(2) = "Dark striped"
m_SequenceNames(3) = "Bright colors"
m_SequenceNames(4) = "Bright striped"
m_SequenceNames(5) = "Cold to hot"
m_SequenceNames(6) = "Cold-hot striped"
m_SequenceNames(7) = "Cycles"
m_SequenceNames(8) = "Cycles striped"
m_SequenceNames(9) = "Many cycles"
m_SequenceNames(10) = "Many ramps"
End Sub

'======== Method ===============================================================
Public Function Color(ByVal value As Double) As Long
Dim c As Double
Dim k As Long
c = value - Int(value)
Dim red As Double, green As Double, blue As Double
Select Case m_Sequence
  Case 0
    c = 256# * c - 0.5  ' conversion in RGB turns this into 0 -> 255
    Color = RGB(c, c, c)
  Case 1, 2, 3, 4
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
    If (m_Sequence = 2) Or (m_Sequence = 4) Then
      c = Abs(4# * Cos(c * c_TwoPi * 5# / 6#))
      If c > 1# Then c = 1#
      red = c * red
      green = c * green
      blue = c * blue
    End If
    If (m_Sequence = 1) Or (m_Sequence = 2) Then
      Color = RGB(180# * red - 0.5, 180# * green - 0.5, 180# * blue - 0.5)
    Else
      Color = RGB(256# * red - 0.5, 256# * green - 0.5, 256# * blue - 0.5)
    End If
  Case 5, 6
    c = 6# * c
    If c < 1# Then
      red = d(c, 72, 0)
      green = d(c, 72, 0)
      blue = d(c, 72, 170)
    ElseIf c < 2# Then
      red = d(c - 1#, 0, 120)
      green = d(c - 1#, 0, 0)
      blue = d(c - 1#, 170, 135)
    ElseIf c < 3# Then
      red = d(c - 2#, 120, 256)
      green = d(c - 2#, 0, 0)
      blue = d(c - 2#, 135, 0)
    ElseIf c < 4# Then
      red = d(c - 3#, 256, 244)
      green = d(c - 3#, 0, 170)
      blue = d(c - 3#, 0, 0)
    ElseIf c < 5# Then
      red = d(c - 4#, 244, 256)
      green = d(c - 4#, 170, 230)
      blue = d(c - 4#, 0, 0)
    ElseIf c >= 5# Then
      red = d(c - 5#, 256, 256)
      green = d(c - 5#, 230, 256)
      blue = d(c - 5#, 0, 256)
    End If
    If m_Sequence = 6 Then
      c = Abs(4# * Cos(c * c_TwoPi * 5# / 6#))
      If c > 1# Then c = 1#
      red = c * red
      green = c * green
      blue = c * blue
    End If
    Color = RGB(red - 0.5, green - 0.5, blue - 0.5)
  Case 7, 8
    red = 128# * (1# - Sin(c * c_TwoPi))
    green = 128# * (1# - Cos(2# * c * c_TwoPi))
    blue = 128# * (1# - Cos(3# * c * c_TwoPi))
    If m_Sequence = 8 Then
      c = Abs(4# * Cos(c * c_TwoPi * 5#))
      If c > 1# Then c = 1#
      red = c * red
      green = c * green
      blue = c * blue
    End If
    Color = RGB(red - 0.5, green - 0.5, blue - 0.5)
  Case 9
    red = 128# * (1# - Sin(c * c_TwoPi))
    green = 128# * (1# - Cos(3# * c * c_TwoPi))
    blue = 128# * (1# - Cos(10# * c * c_TwoPi))
    Color = RGB(red - 0.5, green - 0.5, blue - 0.5)
  Case 10
    red = 256# * Sqr(c)
    green = 256# * Sqr(5# * c - Int(5# * c))
    blue = 256# * Sqr(10# * c - Int(10# * c))
    Color = RGB(red - 0.5, green - 0.5, blue - 0.5)
End Select
End Function

'-------- Private Method -------------------------------------------------------
' Linear interpolation from one integer to another
Private Function d(ByVal fraction As Double, ByVal first As Integer, _
  ByVal last As Integer) As Double
d = (1# - fraction) * first + fraction * last
End Function

'******** Property *************************************************************
Property Get Sequence() As Integer
Sequence = m_Sequence
End Property

Property Let Sequence(ByVal newValue As Integer)
If newValue < 0 Then newValue = 0
If newValue >= c_SequenceCount Then newValue = c_SequenceCount
m_Sequence = newValue
End Property

'******** Property *************************************************************
Property Get SequenceCount() As Integer
SequenceCount = c_SequenceCount
End Property

'======== Method ===============================================================
Public Function SequenceName(ByVal index As Integer) As String
SequenceName = m_SequenceNames(index)
End Function

'******** Property *************************************************************
Property Get Version() As String
Version = c_version
End Property

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


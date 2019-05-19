Attribute VB_Name = "ParamScanMod"
'
'###############################################################################
'#
'# Visual Basic 6 or VBA file "ParamScanMod.bas"
'#
'# Steps through all combinations of a set of parameters, each of which takes on
'# a set of discrete values in arithmetric progression.
'#
'# Devised & coded by John Trenholme - based on Fortran code of 1996-02-08
'#
'###############################################################################

Private Const Version_c As String = "2009-11-03"

' Usage:
'
'  Note: you can use any lower & upper array bounds, as long as all are the same
'
'  Dim paramStart(1& To 4&) As Double  ' control arrays
'  Dim paramStep(1& To 4&) As Double
'  Dim paramCount(1& To 4&) As Long  ' note: count <= 1 means no variation
'
'  Dim j As Long
'  For j = 1& To 4&  'initialize control arrays
'    paramStart(j) = Array(1.1, 5.9, 4.4, 7.5)(j - 1&)
'    paramStep(j) = Array(0.1, -0.2, 4#, 0.5)(j - 1&)
'    paramCount(j) = Array(2&, 3&, 1&, 4&)(j - 1&)
'  Next j
'
' Either
'  Dim params() As Double  ' no explicit dimensions; will be set same as others
' or
'  ReDim params(1& To 4&) As Double  ' same as others
'
'  Do Until ParamScan(paramStart, paramStep, paramCount, params)
'    ' visit present values in params(); dimensions will be set same as others
'    Debug.Print params(1&), params(2&), params(3&), params(4&)
'  Loop

Option Explicit

Private working_m As Boolean  ' module-global to allow reset

'===============================================================================
Public Function ParamScan( _
  ByRef paramStart() As Double, _
  ByRef paramStep() As Double, _
  ByRef paramCount() As Long, _
  ByRef params() As Double) _
As Boolean
' Step through parameter values. Return in params(). See usage notes above.
Static parCount() As Long  ' this holds the state of the stepping machine
If Not working_m Then
  working_m = True
  parCount = paramCount
  params = paramStart  ' return initial values on first call
Else
  Dim j1 As Long, j2 As Long
  j1 = LBound(paramStart)
  j2 = UBound(paramStart)
  Dim j As Long
  For j = j1 To j2  ' try to step each parameter, in sequence
    If paramCount(j) > 1& Then  ' only change parameters with count > 1
      parCount(j) = parCount(j) - 1&
      If parCount(j) = 0& Then  ' this param has reached its last value
        parCount(j) = paramCount(j)
        params(j) = paramStart(j)
      Else  ' step the parameter and stop here
        params(j) = params(j) + paramStep(j)
        Exit For
      End If
    End If
    If j = j2 Then working_m = False  ' no more parameters to change
  Next j
End If
ParamScan = Not working_m
End Function

'===============================================================================
Public Sub ParamScanReset()
' Force initialization of the ParamScan routine on the next call to it.
working_m = False
End Sub

'===============================================================================
Public Function ParamScanVersion() As String
' Date of the latest revision to this file, in the format "yyyy-mm-dd".
ParamScanVersion = Version_c
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


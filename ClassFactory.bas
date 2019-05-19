Attribute VB_Name = "ClassFactory"
'
'###############################################################################
'#
'# VBA Module file "ClassFactory.bas"
'#
'# Return an instance of a Class, given its name in a text String. Also supplies
'# related routines that use a text name.
'#
'# You must code a line in "initClassFactory" for each Class that you want to
'# work with this Module. The only alternative would be evil self-modifying
'# code in an XLA, so suck it up. At least, everything is localized here.
'#
'# Started 2013-01-10 by John Trenholme
'#
'# Exports the routines:
'#   Function classByName
'#   Function ClassFactoryVersion
'#   Function classVersionByName
'#   Sub initClassFactory
'#
'# Requires a Project Reference to "Microsoft Scripting Runtime" for Dictionary
'#
'###############################################################################

Option Base 0          ' array base value when not specified     - default
Option Compare Binary  ' string comparison based on Asc(char)    - default
Option Explicit        ' force explicit declaration of variables - not default
'Option Private Module  ' no effect in VB6; globals project-only in VBA

' Module-global Const values (convention: start with upper-case; suffix "_c")
Private Const Version_c As String = "2013-01-24"
Private Const File_c As String = "ClassFactory[" & Version_c & "]."

' Module-global variables (convention: suffix "_m")
' Retained between calls; initialize as 0, "" or False (etc.)
' the following definition requires a Reference to "Microsoft Scripting Runtime"
Private dict_m As Scripting.Dictionary  ' holds text-name, Class-Object pairs

'===============================================================================
Public Function classByName(ByVal className As String) As Object
' given a case-insensitive Class name as text, return an Object of that Class
' if there is no such Class in the Dictionary, return Nothing (a null object)
If dict_m Is Nothing Then initClassFactory  ' be sure Dictionary is initialized
Set classByName = Nothing  ' default return; no error, just a non-Object
Dim lcName As String: lcName = LCase$(className)
If dict_m.Exists(lcName) Then Set classByName = dict_m.item(lcName)
End Function

'===============================================================================
Public Function ClassFactoryVersion(Optional ByVal trigger As Variant) As String
' Date of latest revision to this file, as a string in the format "yyyy-mm-dd"
ClassFactoryVersion = Version_c
End Function

'===============================================================================
Public Function classVersionByName( _
  ByVal className As String, _
  Optional ByVal trigger As Variant) _
As String
' given a case-insensitive Class name as text, return the no-argument Version
' Function or Property of an Object of that Class (if it exists)
Dim obj As Object: Set obj = classByName(className)
classVersionByName = "<class '" & className & "' not in Dictionary>"
If Not obj Is Nothing Then
  On Error Resume Next
  classVersionByName = obj.Version()
  ' failed if no 0-argument Version there, or result can't become a String
  If 0& <> Err.Number Then
    classVersionByName = "<no " & className & ".Version() found>"
  End If
  On Error GoTo 0
  Set obj = Nothing  ' dispose of temporary local Object
End If
End Function

'===============================================================================
Public Sub initClassFactory()
' put the name-object pairs into the Dictionary, creating it if needed
' the user usually does not need to call this routine (see "classByName")
If dict_m Is Nothing Then Set dict_m = New Scripting.Dictionary  ' if none, make
dict_m.RemoveAll  ' be sure it's empty
' -------- add a line here for each Class that you want to work with this Module
' -------- the only Class code that runs is Class_Initialize(), so make it fast
addClass New StringBuilder
' -------- end of Class specification lines
End Sub

'-------------------------------------------------------------------------------
Private Sub addClass(ByRef obj As Object)
' add one text-name, Class-Object pair to the Dictionary
Dim textKey As String: textKey = LCase$(TypeName(obj))  ' names are caseless
If Not dict_m.Exists(textKey) Then dict_m.add textKey, obj  ' only add once
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

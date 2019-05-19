Attribute VB_Name = "CollAndDict"
'
'###############################################################################
'#
'# VBA Module file "CollAndDict.bas"
'#
'# Support for handling Collection and Scripting.Dictionary objects.
'#
'# Note: use of a Dictionary requires a Project reference to "Microsoft
'# Scripting Runtime" under Tools | References...
'#
'# Started 2012-03-05 by John Trenholme
'#
'# This module exports the routines:
'#   Sub addOrAlter
'#   Sub arrayToDictionary
'#   Function CollAndDictVersion
'#   Sub dictionaryToArray
'#   Sub dictionaryToRangeDown
'#   Function IsInCollection
'#   Sub sortDictionaryByKeys
'#
'###############################################################################

Option Base 0          ' array base value when not specified - the default
Option Compare Binary  ' string comparison based on Asc(char) - the default
Option Explicit        ' forces explicit variable declaration - changes default

' Module-global Const values (convention: start with upper-case; suffix "_c")
Private Const Version_c As String = "2012-03-08"
Private Const File_c As String = "CollAndDict[" & Version_c & "]."

Private Const Key_c As Long = 1&, Val_c = 2&  ' index identifiers for arrays

'===============================================================================
Public Function CollAndDictVersion(Optional ByRef trigger As Variant) As String
' Date of the latest revision to this code, as a string in format "yyyy-mm-dd"
CollAndDictVersion = Version_c
End Function

'===============================================================================
Public Sub addOrAlter( _
  ByRef cl As Collection, _
  ByRef var As Variant, _
  ByVal key As String)
' Add 'var' to Collection 'cl' at 'key' with replacement if already present
' Usage: when adding to Collection "col"
'        change  col.Add itemToAdd, "keyString"  ' error if it already exists
'        to      AddOrAlter col, itemToAdd, "keyString"  ' no error if exists
' You can do the change by using Replace: "col.Add" -> "AddOrAlter col,"
' note: to do this with a Scripting.Dictionary do dic.Item("keyText") = var
' we claim that no error is possible here
If IsInCollection(cl, key) Then cl.Remove key
cl.Add var, key  ' nothing at that key now, so we can Add without error
End Sub

'===============================================================================
Public Sub arrayToDictionary( _
  ByRef theArray As Variant, _
  ByRef theDict As Scripting.Dictionary)
' Copy the array items into the Dictionary, replacing all existing items.
' Because a Dictionary contains Variants, the array must be a Variant. It should
' be dimensioned (1 To 2, lo To hi).
theDict.RemoveAll
Dim j As Long
For j = LBound(theArray, 2&) To UBound(theArray, 2&)
  theDict(theArray(Key_c, j)) = theArray(Val_c, j)
Next j
End Sub

'===============================================================================
Public Sub dictionaryToArray( _
  ByRef theDict As Scripting.Dictionary, _
  ByRef theArray As Variant)
' Copy all Dictionary items into the array(1 To 2, lo To hi) (empty it first).
' The array should be specified as "Dim myArray() as Variant" so that it can be
' ReDim'd here.
ReDim theArray(Key_c To Val_c, 1& To theDict.Count)
Dim j As Long, key As Variant
j = 0&
For Each key In theDict.Keys
  j = j + 1&
  theArray(Key_c, j) = key
  theArray(Val_c, j) = theDict(key)
Next key
End Sub

'===============================================================================
Public Sub dictionaryToRangeDown( _
  ByRef theDict As Scripting.Dictionary, _
  ByRef theRange As Range)
' Copy the items in a Scripting.Dictionary to an Excel worksheet, starting at
' the top left of the supplied Range and moving down. The Range can be any size.
' The result will be two colums wide, showing key and associated item.
Dim j As Long, key As Variant, ary() As Variant
ReDim ary(1& To theDict.Count, Key_c To Val_c)
j = 0&
For Each key In theDict.Keys
  j = j + 1&
  ary(j, Key_c) = key
  ary(j, Val_c) = theDict(key)
Next key
theRange(1&).Resize(theDict.Count, 2&).value = ary
Erase ary
End Sub

'===============================================================================
Public Function IsInCollection(ByRef cl As Collection, key As String) As Boolean
' Return True if an item with the supplied key is in the Collection, else False.
Dim vv As Variant  ' a Collection can hold anything
On Error Resume Next
vv = cl(key)  ' attempt to access whatever is under the key
IsInCollection = (0& = Err.Number)
vv = Nothing  ' just being careful, in case it was huge & auto-destruct failed
Err.clear  ' don't pass Err properties back up to caller (yes!)
End Function

'===============================================================================
Public Sub sortDictionaryByKeys( _
  ByRef theDict As Scripting.Dictionary, _
  Optional ByVal comp As CompareMethod = vbTextCompare)
' Sort the items in a Scripting.Dictionary into increasing order by their key
' values. The keys are treated as strings.
Dim ary() As Variant
' transfer to an array for fast random access
dictionaryToArray theDict, ary

' Quicksort the array on the keys
' TODO move sort to separate routine, for general use and Collections
' Minimum partition size - optimal by experiment, but optimum is broad
Const MaxPart_c As Long = 14&
Const StackIncrement_c As Long = 25&  ' prevents overflow up to 2^31 items
Dim stack() As Long   ' partition-holding stack
Dim stackTop As Long  ' index of highest available entry in stack (capacity)
stackTop = StackIncrement_c + StackIncrement_c  ' so it's 2, 4, 6, ...
ReDim stack(1& To stackTop)   ' initial size of partition-holding stack

' Initialize data-index values, and the lowest-partition-boundary value
Dim bot As Long, top As Long, lpb As Long, nItems As Long
bot = 1&
nItems = theDict.Count
top = nItems
lpb = top

' Initialize stack pointer
Dim stackPointer As Long  ' index of lowest empty entry in stack
stackPointer = 1&

' First part: repeatedly partition parts until array is mostly sorted

Dim leftEl As Long, midEl As Long, riteEl As Long, tm1 As Long
Dim tempKey As Variant, tempVal As Variant
Dim testKey As Variant, testVal As Variant
Dim cr As Integer

Do
  ' Leave short sections alone (they will be insertion sorted later)
  If top - bot >= MaxPart_c Then  ' this section is large enough to partition
    ' Get location of middle element in this partition
    ' Using this element speeds up sort if already nearly sorted
    midEl = bot + (top - bot) \ 2&  ' written this way to avoid overflow
    ' Sort bottom, middle and top values; put median-of-3 in middle
    ' Use a method that does 2 2/3 comparisons on average
    ' Result of StrComp is: -1 -> a < b, 0 -> a = b, 1 -> a > b
    ' note that StrComp follows the language rules in the active Windows locale
    cr = StrComp(ary(Key_c, bot), ary(Key_c, midEl), comp)
    If cr < 0 Then ' order is one of 123 132 231
      cr = StrComp(ary(Key_c, midEl), ary(Key_c, top), comp)
      If cr > 0 Then  ' order is one of 132 231
        tempKey = ary(Key_c, midEl): tempVal = ary(Val_c, midEl)
        cr = StrComp(ary(Key_c, bot), ary(Key_c, top), comp)
        If cr <= 0 Then  ' order is 132 so swap 2 & 3
          ary(Key_c, midEl) = ary(Key_c, top): ary(Val_c, midEl) = ary(Val_c, top)
        Else  ' order is 231 so rotate right
          ary(Key_c, midEl) = ary(Key_c, bot): ary(Val_c, midEl) = ary(Val_c, bot)
          ary(Key_c, bot) = ary(Key_c, top): ary(Val_c, bot) = ary(Val_c, top)
        End If
        ary(Key_c, top) = tempKey: ary(Val_c, top) = tempVal
      End If  ' no Else clause; order already 123 when data(midEl) < data(top)
    Else  ' order is one of 213 312 321
      cr = StrComp(ary(Key_c, midEl), ary(Key_c, top), comp)
      If cr < 0 Then  ' order is one of  213 312
        tempKey = ary(Key_c, midEl): tempVal = ary(Val_c, midEl)
        cr = StrComp(ary(Key_c, bot), ary(Key_c, top), comp)
        If cr <= 0 Then  ' order is 213 so swap 1 & 2
          ary(Key_c, midEl) = ary(Key_c, bot): ary(Val_c, midEl) = ary(Val_c, bot)
        Else  ' order is 312 so rotate left
          ary(Key_c, midEl) = ary(Key_c, top): ary(Val_c, midEl) = ary(Val_c, top)
          ary(Key_c, top) = ary(Key_c, bot): ary(Val_c, top) = ary(Val_c, bot)
        End If
      Else  ' order is 321 so swap 1 & 3
        tempKey = ary(Key_c, top): tempVal = ary(Val_c, top)
        ary(Key_c, top) = ary(Key_c, bot): ary(Val_c, top) = ary(Val_c, bot)
      End If
      ary(Key_c, bot) = tempKey: ary(Val_c, bot) = tempVal  ' min is in temp
    End If
    ' Initialize sweep pointers; don't move partitioning element (at top)
    leftEl = bot
    riteEl = top - 1&
    ' Swap partitioning element iwith next-to-top element
    testKey = ary(Key_c, midEl): testVal = ary(Val_c, midEl)
    ary(Key_c, midEl) = ary(Key_c, riteEl): ary(Val_c, midEl) = ary(Val_c, riteEl)
    ary(Key_c, riteEl) = testKey: ary(Val_c, riteEl) = testVal
    ' Inner loop; swap elements until lesser are to left, greater to right
    ' Note: (Key_c, leftEl) and (Key_c, riteEl) will act as sentinels during sweep
    Do
      Do                                ' sweep left pointer to right
        leftEl = leftEl + 1&
      Loop While StrComp(ary(Key_c, leftEl), testKey, comp) < 0
      Do                                ' sweep right pointer to left
        riteEl = riteEl - 1&
      Loop While StrComp(ary(Key_c, riteEl), testKey, comp) > 0
      If leftEl >= riteEl Then Exit Do  ' quit when pointers have crossed
      tempKey = ary(Key_c, riteEl): tempVal = ary(Val_c, riteEl)
      ary(Key_c, riteEl) = ary(Key_c, leftEl)
        ary(Val_c, riteEl) = ary(Val_c, leftEl)
      ary(Key_c, leftEl) = tempKey: ary(Val_c, leftEl) = tempVal
    Loop
    ' put median back, in order
    tempKey = ary(Key_c, leftEl): tempVal = ary(Val_c, leftEl)
    tm1 = top - 1&
    ary(Key_c, leftEl) = ary(Key_c, tm1): ary(Val_c, leftEl) = ary(Val_c, tm1)
    ary(Key_c, tm1) = tempKey: ary(Val_c, tm1) = tempVal
    ' Adjust lowest-partition-boundary value for use in insertion-sort phase
    If lpb > leftEl - 1& Then lpb = leftEl - 1&
    ' Expand stack if it is about to overflow from inserting 2 items
    If stackPointer + 2& > stackTop Then
      stackTop = stackTop + StackIncrement_c + StackIncrement_c
      ReDim Preserve stack(1&, stackTop)
    End If
    ' Stack larger partition; set limit pointers for smaller (= tail recursion)
    If (leftEl - bot) > (top - riteEl) Then
      stack(stackPointer) = bot          ' stack left partition
      stackPointer = stackPointer + 1&
      stack(stackPointer) = leftEl - 1&
      bot = leftEl + 1&                  ' set to do right partition next
    Else
      stack(stackPointer) = leftEl + 1&  ' stack right partition
      stackPointer = stackPointer + 1&
      stack(stackPointer) = top
      top = leftEl - 1&                  ' set to do left partition next
    End If
    ' Bump stack pointer to next empty value
    stackPointer = stackPointer + 1&
  Else
    ' Partition is "too short to sort"; try to get next partition from stack
    If stackPointer <= 1& Then Exit Do  ' if stack is empty, exit first part
    stackPointer = stackPointer - 2&    ' drop stack pointer
    bot = stack(stackPointer)           ' unstack top partition
    top = stack(stackPointer + 1&)
  End If
Loop

' Second part: use insertion sort to fix up remaining disorder

' Find smallest element in lowest partition
testKey = ary(Key_c, 1&)  ' start with first element
bot = 1&
For top = 2& To lpb
  If StrComp(testKey, ary(Key_c, top), comp) > 0 Then
    testKey = ary(Key_c, top)
    bot = top
  End If
Next top
' Put smallest element in lowest partition at bottom to use as sentinel
tempKey = ary(Key_c, 1&): tempVal = ary(Val_c, 1&)
ary(Key_c, 1&) = ary(Key_c, bot): ary(Val_c, 1&) = ary(Val_c, bot)
ary(Key_c, bot) = tempKey: ary(Val_c, bot) = tempVal
' Do insertion sort on the entire array (single-test inner loop is faster)
For top = 3& To nItems  ' elements below 'top' are sorted
  ' remember top value
  tempKey = ary(Key_c, top): tempVal = ary(Val_c, top)
  midEl = top - 1&  ' look below top for larger elements
  ' if element is larger...
  Do While StrComp(ary(Key_c, midEl), tempKey, comp) > 0
    ' ... move it up
    ary(Key_c, midEl + 1&) = ary(Key_c, midEl)
      ary(Val_c, midEl + 1&) = ary(Val_c, midEl)
    ' ... and look farther down
    midEl = midEl - 1&
  Loop
  ' put top value into empty hole
  ary(Key_c, midEl + 1&) = tempKey: ary(Val_c, midEl + 1&) = tempVal
Next top

' Move array back to Dictionary
arrayToDictionary ary, theDict

' Sometimes VB gets confused about auto-erase of dynamic arrays
Erase ary, stack
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~


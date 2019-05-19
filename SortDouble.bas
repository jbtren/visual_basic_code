Attribute VB_Name = "SortDouble"
Attribute VB_Description = "Quicksort of Double values in an array, in-place."
'
'       _____               _    _____                 _      _.
'      / ____|             | |  |  __ \               | |    | |
'     | (___    ___   _ __ | |_ | |  | |  ___   _   _ | |__  | |  ___
'      \___ \  / _ \ | '__|| __|| |  | | / _ \ | | | || '_ \ | | / _ \
'      ____) || (_) || |   | |_ | |__| || (_) || |_| || |_) || ||  __/
'     |_____/  \___/ |_|    \__||_____/  \___/  \__,_||_.__/ |_| \___|
'
'###############################################################################
'#
'# Excel Visual Basic for Applications (VBA) Module file "SortDouble.bas"
'#
'# Sort routine that rearranges elements of a Double array in place.
'#
'# Devised and coded by John Trenholme
'#
'# Exports the routine:
'#   Sub sortDbl
'#   Sub sortDblTest
'#   Function sortDblVersion
'#
'###############################################################################

Option Explicit

Private Const Version_c As String = "2010-09-25"

' Magic value used to indicate that user has not supplied jHi & (possibly) jLo
Const MagicLong_c As Long = -2147483648#  ' negative 2^31
' Minimum partition size - optimal by experiment, but optimum is broad
Const MaxPart_c As Long = 14&

'===============================================================================
Public Sub sortDbl( _
  ByRef data() As Double, _
  Optional ByVal jLo As Long = MagicLong_c, _
  Optional ByVal jHi As Long = MagicLong_c)
Attribute sortDbl.VB_Description = "Uses non-recursive, median-of-3 quicksort (plus a final insertion-sort pass) to put the elements from the array data(jLo) to data(jHi) into increasing numerical order, in-place. If jHi is not supplied, it is set to the highest index in the array. If jLo is not supplied, it is set to the lowest index in the array."
' Uses non-recursive, median-of-3 quicksort (plus a final insertion-sort pass)
' to put the elements from the array data(jLo) to data(jHi) into increasing
' numerical order, in-place. If jHi is not supplied, it is set to the highest
' index in the array. If jLo is not supplied, it is set to the lowest index in
' the array. Quicksort is an N log N method on the average, but under extremely
' unlikely input orders it will drop to N^2 behavior. See Robert Sedgewick's
' book "Algorithms [in C[++]]" for a *very* detailed discussion of Quicksort.
' Note that this version is not optimal when many repeated values exist.

' Handle index limits & optional arguments (indicated by "magic" index value)
If jHi = MagicLong_c Then
  On Error Resume Next
  jHi = UBound(data)
  ' raise an error if the input array is not dimensioned (UBound failed)
  If Err.Number <> 0& Then
    Dim ID As String
    ID = "SortDouble[" & Version_c & "].sortDbl"
    Err.Raise 4337&, ID, _
      "Input array not dimensioned" & vbLf & _
      "Problem in " & ID
  End If
  On Error GoTo 0
  If jLo = MagicLong_c Then jLo = LBound(data)  ' if UBound OK, then LBound OK
End If

If jLo >= jHi Then Exit Sub  ' do not sort 1 item

Const StackIncrement_c As Long = 25&  ' prevents overflow up to 2^31 items
Dim stack() As Long  ' partition-holding stack
Dim stackTop As Long
stackTop = StackIncrement_c + StackIncrement_c  ' so it's 2, 4, 6, ...
ReDim stack(1& To stackTop)   ' initial size of partition-holding stack

' Initialize data-index values, and the lowest-partition-boundary value
Dim bot As Long, top As Long, lpb As Long
bot = jLo
top = jHi
lpb = top

' Initialize stack pointer
Dim stackPointer As Long
stackPointer = 1&

' First: repeatedly partition parts until array is mostly sorted

Dim leftEl As Long, midEl As Long, riteEl As Long
Dim temp As Double, test As Double

Do
  ' Leave short sections alone (they will be insertion sorted later)
  If top - bot >= MaxPart_c Then
    ' Get location of middle element in this partition
    ' Using this element speeds up sort if already nearly sorted
    midEl = bot + (top - bot) \ 2&  ' written this way to avoid overflow
    ' Sort bottom, middle and top values; put median-of-3 in middle
    ' Use a method that does 2 2/3 comparisons on average
    If data(bot) < data(midEl) Then  ' order is one of 123 132 231
      If data(midEl) > data(top) Then  ' order is one of 132 231
        temp = data(midEl)
        If data(bot) <= data(top) Then  ' order is 132 so swap 2 & 3
          data(midEl) = data(top)
        Else  ' order is 231 so rotate right
          data(midEl) = data(bot)
          data(bot) = data(top)
        End If
        data(top) = temp
      End If  ' no Else clause; order already 123 when data(midEl) < data(top)
    Else  ' order is one of 213 312 321
      If data(midEl) < data(top) Then  ' order is one of  213 312
        temp = data(midEl)
        If data(bot) <= data(top) Then  ' order is 213 so swap 1 & 2
          data(midEl) = data(bot)
        Else  ' order is 312 so rotate left
          data(midEl) = data(top)
          data(top) = data(bot)
        End If
      Else  ' order is 321 so swap 1 & 3
        temp = data(top)
        data(top) = data(bot)
      End If
      data(bot) = temp
    End If
    ' Initialize sweep pointers; don't move partitioning element
    leftEl = bot
    riteEl = top - 1&
    ' Put partitioning element in next-to-top position
    ' Note: data(leftEl) and data(riteEl) will act as sentinels during sweep
    test = data(midEl)
    data(midEl) = data(riteEl)
    data(riteEl) = test
    ' Swap elements until lesser are to left, greater to right
    Do
      Do                                ' sweep left pointer to right
        leftEl = leftEl + 1&
      Loop While data(leftEl) < test
      Do                                ' sweep right pointer to left
        riteEl = riteEl - 1&
      Loop While data(riteEl) > test
      If leftEl >= riteEl Then Exit Do  ' quit if pointers have crossed
      temp = data(riteEl)               ' put elements in order
      data(riteEl) = data(leftEl)
      data(leftEl) = temp
    Loop
    temp = data(leftEl)                ' put median back, in order
    data(leftEl) = data(top - 1&)
    data(top - 1&) = temp
    ' Adjust lowest-partition-boundary value for use in insertion-sort phase
    If lpb > leftEl - 1& Then lpb = leftEl - 1&
    ' Expand stack if it is about to overflow from inserting 2 items
    If stackPointer + 2& > stackTop Then
      stackTop = stackTop + StackIncrement_c + StackIncrement_c
      ReDim Preserve stack(1&, stackTop)
    End If
    ' Stack larger partition; set limit pointers for smaller
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

' Second: use insertion sort to fix up remaining disorder

' Find smallest element in lowest partition
test = data(jLo)
bot = jLo
For top = jLo + 1& To lpb
  If test > data(top) Then
    test = data(top)
    bot = top
  End If
Next top
' Put smallest element in lowest partition at bottom to use as sentinel
temp = data(jLo)
data(jLo) = data(bot)
data(bot) = temp
' Do insertion sort on the entire array (1-test inner loop is faster)
For top = jLo + 2& To jHi           ' elements below 'top' are sorted
  temp = data(top)                  ' remember top value
  midEl = top - 1&                  ' look below for larger elements
  Do While data(midEl) > temp       ' if element is larger
    data(midEl + 1&) = data(midEl)  ' ... move it up
    midEl = midEl - 1&              ' ... and look farther down
  Loop
  data(midEl + 1&) = temp           ' put top value into empty hole
Next top

' Sometimes VB gets confused about auto-erase of dynamic arrays
Erase stack
End Sub

'===============================================================================
Public Sub sortDblTest()
Attribute sortDblTest.VB_Description = "This is used to set up an array and sort it, for debug or unit test."
' This is used to set up an array and sort it, for debug or unit test.
Const NumEls_c As Long = 15&
Dim vals(1& To NumEls_c) As Double
Dim j As Long
If Rnd(-1) >= 0! Then Randomize 1  ' select pseudo-random sequence here
For j = LBound(vals) To UBound(vals)
  vals(j) = Rnd()
Next j
Debug.Print "=== Sorting"; NumEls_c; "random values"
sortDbl vals
For j = LBound(vals) + 1& To UBound(vals)
  If vals(j - 1&) >= vals(j) Then Stop
Next j
Debug.Print "    sort worked correctly"
End Sub

'===============================================================================
Public Function sortDblVersion() As String
Attribute sortDblVersion.VB_Description = "Returns date of last revision as a string in the format ""YYYY-MM-DD""."
' Returns date of last revision as a string in the format "YYYY-MM-DD".
sortDblVersion = Version_c
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


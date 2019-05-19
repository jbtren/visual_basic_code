Attribute VB_Name = "Sorters"
Attribute VB_Description = "Module supporting direct and indirect (indexed) sorting of Singles and Doubles."
'
'###############################################################################
'#                     ____            __
'#                    / __/___   ____ / /_ ___  ____ ___
'#                   _\ \ / _ \ / __// __// -_)/ __/(_-<
'#                  /___/ \___//_/   \__/ \__//_/  /___/
'#
'# Visual Basic 6 and VBA Module file "Sorters.bas"
'#
'# Direct and indirect (indexed) quicksort of numeric arrays.
'#
'# Devised & coded by John Trenholme - started 2004-06-01
'#
'# Exports the routines:
'#   Function decreaseCount
'#   Function sortersVersion
'#   Sub sortHsDirectDbl
'#   Sub sortQsDirectDbl
'#   Sub sortQsDirectSng
'#   Sub sortQsIndirectDbl
'#   Sub sortQsIndirectSng
'#   Function sortQsStackDepth
'#   Sub sortSsDirectDbl
'#
'###############################################################################

Option Base 0          ' array base value, when not specified - the default
Option Compare Binary  ' string comparison based on Asc(char) - the default
Option Explicit        ' forces explicit variable declaration - changes default

Private Const c_Version As String = "2013-05-10"

'*******************************************************************************
'*
'* Declarations and module-global quantities
'*
'*******************************************************************************

' Stack space for partitions - max. amount used is approx. 5 * log10( N)
' Note that stack holds two entries per stacked partition
' If the stack ever fills, sorting will continue but will be slower
Private Const c_MaxStack As Long = 50&      ' prevent overflow up to 2^31 items
' Max partition size that will be left unsorted
' This value is optimal by test, but values from 5 to 20 work quite well
' Smaller values lead to more stack usage, so we err on the high side
Private Const c_MaxPart As Long = 14&

Private Const c_MaxLong As Long = 2147483647

Private m_stack(1& To c_MaxStack) As Long  ' partition stack
Private m_maxStack As Long                 ' stack high water mark

'*******************************************************************************
'*
'* Routines
'*
'*******************************************************************************

'===============================================================================
Public Function decreaseCount(ByRef data() As Double) As Long
' Return the number of decreases in the values in "data". If the number is zero,
' the array is sorted in increasing order; if it is equal to N-1 where N is the
' number of items in "data", the array is sorted in strictly decreasing order.
Dim j As Long, decrements As Long
decrements = 0&
For j = LBound(data) + 1& To UBound(data)
  If data(j) < data(j - 1&) Then decrements = decrements + 1&
Next j
decreaseCount = decrements
End Function

'===============================================================================
' Return the code's version as a date in a String ("YYYY-MM-DD").
' This is a Function so Excel etc. can use it.
Public Function sortersVersion() As String
sortersVersion = c_Version
End Function

'===============================================================================
Sub sortHsDirectDbl( _
  ByRef data() As Double, _
  Optional ByVal jLo As Long = c_MaxLong, _
  Optional ByVal jHi As Long = -c_MaxLong)
' Uses "heapsort" (invented by J. W. J. Williams, CACM 7 (1964), 347-348)
' to put the elements from data(jLo) to data(jHi) into increasing numerical
' order, in-place.
' This is the implementation from D. Knuth's book "Sorting and Searching"
' with the speedup due to R. W. Floyd (problem 5.2.3-18), recoded to remove
' the need for a sentinel element at the bottom, and to speed things up a bit.

' Handle optional arguments (indicated by "impossible" index values)
If jHi = -c_MaxLong Then
  jHi = UBound(data)
  If jLo = c_MaxLong Then jLo = LBound(data)
End If

If jLo >= jHi Then Exit Sub  ' do not sort 1 or fewer items; silent error

Dim m As Long  ' data-array difference from base-1 indices
m = jLo - 1&

Dim n As Long  ' number of items to sort
n = jHi - jLo + 1&

'-------- first, arrange the data elements into a heap; O(N) work -------
' last half of the data consists of 1-element "heaps" = leaves; don't do them
Dim i As Long, j As Long, k As Long
Dim temp As Double
For i = n \ 2& To 1& Step -1&
  j = i                                ' downheap(i)
  GoTo EntryPointA_L
  Do
    temp = data(j + m)                 ' swap data(j), data(k)
    data(j + m) = data(k + m)
    data(k + m) = temp
    j = k                              ' move pointers down to next level
EntryPointA_L:
    k = j + j                          ' index of left son
    If k > n Then
      Exit Do                          ' if past heap end
    ElseIf k < n Then                  ' test right son - it exists
      If data(k + m) < data(k + jLo) Then k = k + 1&
    End If                             ' k now points to larger son
  Loop While data(j + m) < data(k + m) ' keep going until past heap end
Next i
'Debug.Print "Heapified:"
'For j = jLo To jHi
'  Debug.Print Round(data(j), 3&); " ";
'Next j
'Debug.Print
'Debug.Print "Partial Sorts:"

'-------- second, remove the heap elements in order; O(N log N) work -----
' note it's faster to go all the way down the heap without checking, and
' then back up a little, since usually data(i+1) belongs near the bottom
For i = n - 1& To 2& Step -1&          ' i marks end of present heap
  temp = data(i + jLo)                 ' largest -> output; leaf -> top
  data(i + jLo) = data(jLo)
  data(jLo) = temp
  j = 1&                               ' j is index of father node
  GoTo EntryPointB_L
  ' this loop is where most of the time is spent
  Do
    If k < i Then                      ' test right son if it exists
      If data(k + m) < data(k + jLo) Then k = k + 1&
    End If                             ' k now points to larger son
    data(j + m) = data(k + m)          ' move data up to make space
    j = k                              ' move pointers down to next level
EntryPointB_L:
    k = j + j                          ' index of left son
  Loop While k <= i                    ' keep going until past present heap end
  ' hole is now at bottom; move back up 0 (85%), 1 (13.5%), 2 (1.6%), ...
  GoTo EntryPointC_L                   ' roles of j and k are reversed
  Do                                   ' mostly, we don't enter this loop
    data(j + m) = data(k + m)          ' move data back down to make space
    If k <= jLo Then Exit Do           ' don't go beyond top of heap
    j = k                              ' move pointers up to next level
EntryPointC_L:
    k = j \ 2&                         ' father of node j
  Loop While data(k + m) < temp
  data(j + m) = temp                   ' insert data; heap is rebuilt
'  For j = jLo To jHi
'    Debug.Print Round(data(j), 3&); " ";
'  Next j
'  Debug.Print
Next i

'-------- third, do the last two elements with a simple swap -----
If data(jLo) > data(jLo + 1&) Then
  temp = data(jLo)
  data(jLo) = data(jLo + 1&)
  data(jLo + 1&) = temp
End If
'Debug.Print "Final:"
'For j = jLo To jHi
'  Debug.Print Round(data(j), 3&); " ";
'Next j
'Debug.Print
'If decreaseCount(data) > 0& Then Err.Raise 1&, , "Array not sorted"
End Sub

'===============================================================================
' Uses non-recursive, median-of-3 quicksort (plus a final insertion-sort pass)
' to put the elements from data(jLo) to data(jHi) into increasing numerical
' order, in-place. Quicksort is an N log N method on the average, but under
' extremely unlikely input orders it will drop to N^2 behavior. See Sedgewick's
' book "Algorithms [in C[++]]" for a *very* detailed discussion.
Public Sub sortQsDirectDbl( _
  ByRef data() As Double, _
  Optional ByVal jLo As Long = c_MaxLong, _
  Optional ByVal jHi As Long = -c_MaxLong)
Attribute sortQsDirectDbl.VB_Description = "Puts the Double elements from data(jLo) to data(jHi) into increasing numerical order. Default (jLo & jHi not supplied) is to sort entire array."

' Handle optional arguments (indicated by "impossible" index values)
If jHi = -c_MaxLong Then
  jHi = UBound(data)
  If jLo = c_MaxLong Then jLo = LBound(data)
End If

If jLo >= jHi Then Exit Sub  ' do not sort 1 or fewer items

' Initialize data index values, and the lowest-partition-boundary value
Dim bot As Long, Top As Long, lpb As Long
bot = jLo
Top = jHi
lpb = Top

' Initialize stack pointer and high-water marker
Dim stackPointer As Long
stackPointer = 1&
m_maxStack = stackPointer

' ==== First: repeatedly partition parts until array is mostly sorted ====

Dim Left As Long, mid As Long, rite As Long
Dim temp As Double, test As Double

Do
  ' Leave short sections alone (insertion sort later); avoid stack overflow
  If (Top - bot >= c_MaxPart) And (stackPointer < c_MaxStack) Then
    ' Get location of middle element in this partition
    mid = bot + (Top - bot) \ 2&       ' Written this way to avoid overflow
    ' Sort bottom, middle and top values; put median-of-3 in middle
    If data(bot) > data(Top) Then
      temp = data(bot)
      data(bot) = data(Top)
      data(Top) = temp
    End If
    If data(bot) > data(mid) Then
      temp = data(bot)
      data(bot) = data(mid)
      data(mid) = temp
    End If
    If data(mid) > data(Top) Then
      temp = data(mid)
      data(mid) = data(Top)
      data(Top) = temp
    End If
    ' Swap elements until lesser are to left, greater to right
    Left = bot                         ' Initialize sweep pointers
    rite = Top - 1
    test = data(mid)                   ' Use median value to split
    data(mid) = data(rite)             ' Allow right value to be moved
    data(rite) = test
    Do
      Do                               ' Sweep left pointer to right
        Left = Left + 1&
      Loop Until data(Left) >= test
      Do                               ' Sweep right pointer to left
        rite = rite - 1&
      Loop Until data(rite) <= test
      If Left > rite Then Exit Do      ' Quit if pointers have crossed
      temp = data(rite)                ' Put elements in order
      data(rite) = data(Left)
      data(Left) = temp
    Loop
    temp = data(Left)                  ' Put median back in
    data(Left) = data(Top - 1&)
    data(Top - 1&) = temp
    ' Adjust lowest-partition-boundary value
    If lpb > Left - 1& Then lpb = Left - 1&
    ' Stack larger partition; set limit pointers for smaller
    If (Left - bot) > (Top - rite) Then
      m_stack(stackPointer) = bot        ' Stack left partition
      stackPointer = stackPointer + 1&
      m_stack(stackPointer) = Left - 1&
      bot = Left + 1&                  ' Set to do right partition next
    Else
      m_stack(stackPointer) = Left + 1&  ' Stack right partition
      stackPointer = stackPointer + 1&
      m_stack(stackPointer) = Top
      Top = Left - 1&                  ' Set to do left partition next
    End If
    ' Track max. value of stack pointer actually used
    If m_maxStack < stackPointer Then m_maxStack = stackPointer
    ' Bump stack pointer to next empty value
    stackPointer = stackPointer + 1&
  Else
    ' Short partition or stack overflow; try to get next partition from stack
    If stackPointer <= 1& Then Exit Do ' If stack is empty, exit Part 1
    stackPointer = stackPointer - 2&   ' Drop stack pointer
    If stackPointer >= 1& Then
      bot = m_stack(stackPointer)      ' Unstack top partition
      Top = m_stack(stackPointer + 1&)
    End If
  End If
Loop

' ==== Second: use insertion sort to fix up remaining disorder ====

' Find smallest element in lowest partition
test = data(jLo)
bot = jLo
For Top = jLo + 1& To lpb
  If test > data(Top) Then
    test = data(Top)
    bot = Top
  End If
Next Top
' put smallest element at bottom to use as sentinel
temp = data(jLo)
data(jLo) = data(bot)
data(bot) = temp
' do insertion sort on the entire array (1-test inner loop is faster)
For Top = jLo + 2& To jHi              ' Elements below 'top' are sorted
    temp = data(Top)                   ' Remember top value
    mid = Top - 1&                     ' Look below for larger elements
    Do While data(mid) > temp          ' If element is larger
        data(mid + 1&) = data(mid)     ' ... move it up
        mid = mid - 1&                 ' ... and look farther down
    Loop
    data(mid + 1&) = temp              ' Put top value into empty hole
Next Top
End Sub

'===============================================================================
' Uses non-recursive, median-of-3 quicksort (plus a final insertion-sort pass)
' to put the elements from data(jLo) to data(jHi) into increasing numerical
' order, in-place. Quicksort is an N log N method on the average, but under
' extremely unlikely input orders it will drop to N^2 behavior. See Sedgewick's
' book "Algorithms [in C[++]]" for a *very* detailed discussion.
Public Sub sortQsDirectSng( _
  ByRef data() As Single, _
  Optional ByVal jLo As Long = c_MaxLong, _
  Optional ByVal jHi As Long = -c_MaxLong)
Attribute sortQsDirectSng.VB_Description = "Puts the Single elements from data(jLo) to data(jHi) into increasing numerical order. Default (jLo & jHi not supplied) is to sort entire array."

' Handle optional arguments (indicated by "impossible" index values)
If jHi = -c_MaxLong Then
  jHi = UBound(data)
  If jLo = c_MaxLong Then jLo = LBound(data)
End If

If jLo >= jHi Then Exit Sub  ' do not sort 1 or fewer items

' Initialize data index values, and the lowest-partition-boundary value
Dim bot As Long, Top As Long, lpb As Long
bot = jLo
Top = jHi
lpb = Top

' Initialize stack pointer and high-water marker
Dim stackPointer As Long
stackPointer = 1&
m_maxStack = stackPointer

' First: repeatedly partition parts until array is mostly sorted

Dim Left As Long, mid As Long, rite As Long
Dim temp As Single, test As Single
Do
  ' Leave short sections alone (do later); avoid stack overflow
  If (Top - bot > c_MaxPart) And (stackPointer < c_MaxStack) Then
    ' Get location of middle element in this partition
    mid = bot + (Top - bot) \ 2&       ' Written this way to avoid overflow
    ' Sort bottom, middle and top values; put median-of-3 in middle
    If data(bot) > data(Top) Then
      temp = data(bot)
      data(bot) = data(Top)
      data(Top) = temp
    End If
    If data(bot) > data(mid) Then
      temp = data(bot)
      data(bot) = data(mid)
      data(mid) = temp
    End If
    If data(mid) > data(Top) Then
      temp = data(mid)
      data(mid) = data(Top)
      data(Top) = temp
    End If
    ' Swap elements until lesser are to left, greater to right
    Left = bot                         ' Initialize sweep pointers
    rite = Top - 1
    test = data(mid)                   ' Use median value to split
    data(mid) = data(rite)             ' Allow right value to be moved
    data(rite) = test
    Do
      Do                               ' Sweep left pointer to right
        Left = Left + 1&
      Loop Until data(Left) >= test
      Do                               ' Sweep right pointer to left
        rite = rite - 1&
      Loop Until data(rite) <= test
      If Left > rite Then Exit Do      ' Quit if pointers have crossed
      temp = data(rite)                ' Put elements in order
      data(rite) = data(Left)
      data(Left) = temp
    Loop
    temp = data(Left)                  ' Put median back in
    data(Left) = data(Top - 1&)
    data(Top - 1&) = temp
    ' Adjust lowest-partition-boundary value
    If lpb > Left - 1& Then lpb = Left - 1&
    ' Stack larger partition; set limit pointers for smaller
    If (Left - bot) > (Top - rite) Then
      m_stack(stackPointer) = bot        ' Stack left partition
      stackPointer = stackPointer + 1&
      m_stack(stackPointer) = Left - 1&
      bot = Left + 1&                  ' Set to do right partition next
    Else
      m_stack(stackPointer) = Left + 1&  ' Stack right partition
      stackPointer = stackPointer + 1&
      m_stack(stackPointer) = Top
      Top = Left - 1&                  ' Set to do left partition next
    End If
    ' Track max. value of stack pointer actually used
    If m_maxStack < stackPointer Then m_maxStack = stackPointer
    ' Bump stack pointer to next empty value
    stackPointer = stackPointer + 1&
  Else
    ' Short partition or stack overflow; try to get next partition from stack
    If stackPointer <= 1& Then Exit Do ' If stack is empty, exit Part 1
    stackPointer = stackPointer - 2&   ' Drop stack pointer
    If stackPointer >= 1& Then
      bot = m_stack(stackPointer)      ' Unstack top partition
      Top = m_stack(stackPointer + 1&)
    End If
  End If
Loop

' Second: use insertion sort to fix up remaining disorder

' Find smallest element in lowest partition
test = data(jLo)
bot = jLo
For Top = jLo + 1& To lpb
  If test > data(Top) Then
    test = data(Top)
    bot = Top
  End If
Next Top
' put smallest element at bottom to use as sentinel
temp = data(jLo)
data(jLo) = data(bot)
data(bot) = temp
' do insertion sort on the entire array
For Top = jLo + 2& To jHi              ' Elements below 'top' are sorted
    temp = data(Top)                   ' Remember top value
    mid = Top - 1&                     ' Look below for larger elements
    Do While data(mid) > temp          ' If element is larger
        data(mid + 1&) = data(mid)     ' ... move it up
        mid = mid - 1&                 ' ... and look farther down
    Loop
    data(mid + 1&) = temp              ' Put top value into empty hole
Next Top
End Sub

'===============================================================================
' Uses non-recursive, median-of-3 quicksort (plus a final insertion-sort pass)
' to move the elements of the index array ndx() in-place into an order such that
' the data values from data(ndx(jLo)) to data(ndx(jHi)) are in increasing
' numerical order. The data is not moved; only the index array values are
' changed. This sort is not stable, so index values referring to identical data
' values may change order during the sort. Quicksort is an N log N method on the
' average, but under extremely unlikely input orders it will drop to N^2
' behavior. See Sedgewick's book "Algorithms [in C[++]]" for a *very* detailed
' discussion.
'
' Be sure to initialize ndx to contain the index values of the items you want
' to access in sorted order. The most likely initialization is ndx(j) = j.
Public Sub sortQsIndirectDbl( _
  ByRef ndx() As Long, _
  ByRef data() As Double, _
  Optional ByVal jLo As Long = c_MaxLong, _
  Optional ByVal jHi As Long = -c_MaxLong)
Attribute sortQsIndirectDbl.VB_Description = "Adjusts ""ndx"" so that the Double elements from data(ndx(jLo)) to data(ndx(jHi)) are in increasing numerical order. Default (jLo & jHi not supplied) is to adjust all of ""ndx"". Be sure to initialize ""ndx"" to point to the elements you want to be sorted."
  
' Handle optional arguments (indicated by "impossible" index values)
If jHi = -c_MaxLong Then
  jHi = UBound(ndx)
  If jLo = c_MaxLong Then jLo = LBound(ndx)
End If

If jLo >= jHi Then Exit Sub  ' do not sort 1 or fewer items

' Initialize data index values, and the lowest-partition-boundary value
Dim bot As Long, Top As Long, lpb As Long
bot = jLo
Top = jHi
lpb = Top

' Initialize stack pointer and high-water marker
Dim stackPointer As Long
stackPointer = 1&
m_maxStack = stackPointer

' First: repeatedly partition parts until array is mostly sorted

Dim Left As Long, mid As Long, rite As Long
Dim temp As Long, test As Double
Do
  ' Leave short sections alone (insertion sort later); avoid stack overflow
  If (Top - bot > c_MaxPart) And (stackPointer < c_MaxStack) Then
    ' Get location of middle element in this partition
    mid = bot + (Top - bot) \ 2&       ' Written this way to avoid overflow
    ' Sort bottom, middle and top values; put median-of-3 in middle
    If data(ndx(bot)) > data(ndx(Top)) Then
      temp = ndx(bot)
      ndx(bot) = ndx(Top)
      ndx(Top) = temp
    End If
    If data(ndx(bot)) > data(ndx(mid)) Then
      temp = ndx(bot)
      ndx(bot) = ndx(mid)
      ndx(mid) = temp
    End If
    If data(ndx(mid)) > data(ndx(Top)) Then
      temp = ndx(mid)
      ndx(mid) = ndx(Top)
      ndx(Top) = temp
    End If
    ' Swap elements until lesser are to left, greater to right
    Left = bot                         ' Initialize sweep pointers
    rite = Top - 1&
    temp = ndx(mid)                    ' Allow right value to be moved
    ndx(mid) = ndx(rite)
    ndx(rite) = temp
    test = data(temp)                  ' Use median value to split
    Do
      Do                               ' Sweep left pointer to right
        Left = Left + 1&
      Loop Until data(ndx(Left)) >= test
      Do                               ' Sweep right pointer to left
        rite = rite - 1&
      Loop Until data(ndx(rite)) <= test
      If Left > rite Then Exit Do      ' Quit if pointers have crossed
      temp = ndx(rite)                 ' Put elements in order
      ndx(rite) = ndx(Left)
      ndx(Left) = temp
    Loop
    temp = ndx(Left)                   ' Put median back in
    ndx(Left) = ndx(Top - 1&)
    ndx(Top - 1&) = temp
    ' Adjust lowest-partition-boundary value
    If lpb > Left - 1& Then lpb = Left - 1&
    ' Stack larger partition; set limit pointers for smaller
    If (Left - bot) > (Top - rite) Then
      m_stack(stackPointer) = bot        ' Stack left partition
      stackPointer = stackPointer + 1&
      m_stack(stackPointer) = Left - 1&
      bot = Left + 1&                  ' Set to do right partition next
    Else
      m_stack(stackPointer) = Left + 1&  ' Stack right partition
      stackPointer = stackPointer + 1&
      m_stack(stackPointer) = Top
      Top = Left - 1&                  ' Set to do left partition next
    End If
    ' Track max. value of stack pointer actually used
    If m_maxStack < stackPointer Then m_maxStack = stackPointer
    ' Bump stack pointer to next empty value
    stackPointer = stackPointer + 1&
  Else
    ' Short partition or stack overflow; try to get next partition from stack
    If stackPointer <= 1& Then Exit Do ' If stack is empty, exit Part 1
    stackPointer = stackPointer - 2&   ' Drop stack pointer
    bot = m_stack(stackPointer)        ' Unstack top partition
    Top = m_stack(stackPointer + 1&)
  End If
Loop

' Second: use insertion sort to fix up remaining disorder

' Find smallest element in lowest partition
test = data(ndx(jLo))
bot = jLo
For Top = jLo + 1& To lpb
  If test > data(ndx(Top)) Then
    test = data(ndx(Top))
    bot = Top
  End If
Next Top
' put index of smallest element at bottom to use as sentinel
temp = ndx(jLo)
ndx(jLo) = ndx(bot)
ndx(bot) = temp
For Top = jLo + 2& To jHi              ' Elements below 'top' are sorted
    temp = ndx(Top)                    ' Remember top index
    test = data(temp)                  ' Remember top value
    mid = Top - 1&                     ' Look below for larger elements
    Do While data(ndx(mid)) > test     ' If element is larger
        ndx(mid + 1&) = ndx(mid)       ' ... move it up
        mid = mid - 1&                 ' ... and look farther down
    Loop
    ndx(mid + 1&) = temp               ' Put top value into empty hole
Next Top
End Sub

'===============================================================================
' Uses non-recursive, median-of-3 quicksort (plus a final insertion-sort pass)
' to move the elements of the index array ndx() in-place into an order such that
' the data values from data(ndx(jLo)) to data(ndx(jHi)) are in increasing
' numerical order. The data is not moved; only the index array values are
' changed. This sort is not stable, so index values referring to identical data
' values may change order during the sort. Quicksort is an N log N method on the
' average, but under extremely unlikely input orders it will drop to N^2
' behavior. See Sedgewick's book "Algorithms [in C[++]]" for a *very* detailed
' discussion.
'
' Be sure to initialize to contain the index values of the items you want
' to access in sorted order. The most likely initialization is ndx(j) = j.
Public Sub sortQsIndirectSng( _
  ByRef ndx() As Long, _
  ByRef data() As Single, _
  Optional ByVal jLo As Long = c_MaxLong, _
  Optional ByVal jHi As Long = -c_MaxLong)
Attribute sortQsIndirectSng.VB_Description = "Adjusts ""ndx"" so that the Single elements from data(ndx(jLo)) to data(ndx(jHi)) are in increasing numerical order. Default (jLo & jHi not supplied) is to adjust all of ""ndx"". Be sure to initialize ""ndx"" to point to the elements you want to be sorted."

' Handle optional arguments (indicated by "impossible" index values)
If jHi = -c_MaxLong Then
  jHi = UBound(ndx)
  If jLo = c_MaxLong Then jLo = LBound(ndx)
End If

If jLo >= jHi Then Exit Sub  ' do not sort 1 or fewer items

' Initialize data index values, and the lowest-partition-boundary value
Dim bot As Long, Top As Long, lpb As Long
bot = jLo
Top = jHi
lpb = Top

' Initialize stack pointer and high-water marker
Dim stackPointer As Long
stackPointer = 1&
m_maxStack = stackPointer

' First: repeatedly partition parts until array is mostly sorted

Dim Left As Long, mid As Long, rite As Long
Dim temp As Long, test As Single
Do
  ' Leave short sections alone (insertion sort later); avoid stack overflow
  If (Top - bot > c_MaxPart) And (stackPointer < c_MaxStack) Then
    ' Get location of middle element in this partition
    mid = bot + (Top - bot) \ 2&       ' Written this way to avoid overflow
    ' Sort bottom, middle and top values; put median-of-3 in middle
    If data(ndx(bot)) > data(ndx(Top)) Then
      temp = ndx(bot)
      ndx(bot) = ndx(Top)
      ndx(Top) = temp
    End If
    If data(ndx(bot)) > data(ndx(mid)) Then
      temp = ndx(bot)
      ndx(bot) = ndx(mid)
      ndx(mid) = temp
    End If
    If data(ndx(mid)) > data(ndx(Top)) Then
      temp = ndx(mid)
      ndx(mid) = ndx(Top)
      ndx(Top) = temp
    End If
    ' Swap elements until lesser are to left, greater to right
    Left = bot                         ' Initialize sweep pointers
    rite = Top - 1&
    temp = ndx(mid)                    ' Allow right value to be moved
    ndx(mid) = ndx(rite)
    ndx(rite) = temp
    test = data(temp)                  ' Use median value to split
    Do
      Do                               ' Sweep left pointer to right
        Left = Left + 1&
      Loop Until data(ndx(Left)) >= test
      Do                               ' Sweep right pointer to left
        rite = rite - 1&
      Loop Until data(ndx(rite)) <= test
      If Left > rite Then Exit Do      ' Quit if pointers have crossed
      temp = ndx(rite)                 ' Put elements in order
      ndx(rite) = ndx(Left)
      ndx(Left) = temp
    Loop
    temp = ndx(Left)                   ' Put median back in
    ndx(Left) = ndx(Top - 1&)
    ndx(Top - 1&) = temp
    ' Adjust lowest-partition-boundary value
    If lpb > Left - 1& Then lpb = Left - 1&
    ' Stack larger partition; set limit pointers for smaller
    If (Left - bot) > (Top - rite) Then
      m_stack(stackPointer) = bot        ' Stack left partition
      stackPointer = stackPointer + 1&
      m_stack(stackPointer) = Left - 1&
      bot = Left + 1&                  ' Set to do right partition next
    Else
      m_stack(stackPointer) = Left + 1&  ' Stack right partition
      stackPointer = stackPointer + 1&
      m_stack(stackPointer) = Top
      Top = Left - 1&                  ' Set to do left partition next
    End If
    ' Track max. value of stack pointer actually used
    If m_maxStack < stackPointer Then m_maxStack = stackPointer
    ' Bump stack pointer to next empty value
    stackPointer = stackPointer + 1&
  Else
    ' Short partition or stack overflow; try to get next partition from stack
    If stackPointer <= 1& Then Exit Do ' If stack is empty, exit Part 1
    stackPointer = stackPointer - 2&   ' Drop stack pointer
    bot = m_stack(stackPointer)        ' Unstack top partition
    Top = m_stack(stackPointer + 1&)
  End If
Loop

' Second: use insertion sort to fix up remaining disorder

' Find smallest element in lowest partition
test = data(ndx(jLo))
bot = jLo
For Top = jLo + 1& To lpb
  If test > data(ndx(Top)) Then
    test = data(ndx(Top))
    bot = Top
  End If
Next Top
' put index of smallest element at bottom to use as sentinel
temp = ndx(jLo)
ndx(jLo) = ndx(bot)
ndx(bot) = temp
For Top = jLo + 2& To jHi              ' Elements below 'top' are sorted
    temp = ndx(Top)                    ' Remember top index
    test = data(temp)                  ' Remember top value
    mid = Top - 1&                     ' Look below for larger elements
    Do While data(ndx(mid)) > test     ' If element is larger
        ndx(mid + 1&) = ndx(mid)       ' ... move it up
        mid = mid - 1&                 ' ... and look farther down
    Loop
    ndx(mid + 1&) = temp               ' Put top value into empty hole
Next Top
End Sub

'===============================================================================
' Returns the maximum index used in the stack during the last quicksort.
Public Function sortQsStackDepth() As Long
sortQsStackDepth = m_maxStack
End Function

'===============================================================================
' Uses optimum-increment Shellsort to put the elements from data(jLo) to
' data(jHi) into increasing numerical order, in-place. This method was first
' published by Donald Shell as "A High-Speed Sorting Procedure", CACM 2(7),
' pp. 30-32, 1959. Shell sort is easy to code and quite fast if the proper set
' of increments is used. Here we use combined Ciura + Tokuda-ratio increments.
' This version puts the lowest element into the bottom first, as a sentinel.
Public Sub sortSsDirectDbl( _
  ByRef data() As Double, _
  Optional ByVal jLo As Long = c_MaxLong, _
  Optional ByVal jHi As Long = -c_MaxLong)

' Handle optional arguments (indicated by "impossible" index values)
If jHi = -c_MaxLong Then
  jHi = UBound(data)
  If jLo = c_MaxLong Then jLo = LBound(data)
End If

If jLo >= jHi Then Exit Sub  ' do not sort 1 or fewer items

Dim n As Long  ' number of items to be sorted
n = jHi - jLo + 1&

' Find smallest element
Dim k As Long  ' index of smallest
k = jLo
Dim temp As Double
temp = data(jLo)
Dim j As Long
For j = k + 1& To jHi
  If temp < data(j) Then
    temp = data(j)
    k = j
  End If
Next j
' Put smallest element at bottom as a sentinel (swap with lowest index)
data(k) = data(jLo)
data(jLo) = temp

' the speed of Shellsort depends on the sequence of increments used
' see Ciura (2001 - Best Increments for the Average Case of Shellsort) where he
' carried out an extensive study leading to the sequence 1 4 10 23 ... 701
' this can be extended by using Tokuda's term ratio of 9/4; here the formula
' used is: inc(N) = Int(Exp(0.87671 + 0.81093 * N)) for N >= 7
Dim incs As Variant
incs = VBA.Array( _
  1, 4, 10, 23, 57, 132, 301, 701, 1578, 3551, 7990, 17978, 40451, 91016, _
  204787, 460772, 1036738, 2332660, 5248484, 11809088)  ' incs(0) = 1 etc.
' find the largest increment that's <= 1/3 the array size
Dim maxInc As Long
maxInc = n \ 3& + 1&  ' 2,1 3,2 4,2 5,2 6,3 ...
For k = 1& To UBound(incs)
  If incs(k) >= maxInc Then Exit For
Next k  ' max increment is in incs(k - 1)

' carry out the Shell sort
Dim i As Long, inc As Long
For i = k - 1& To 0& Step -1&  ' use a series of decreasing increments
  inc = incs(i)
  For j = inc + jLo To jHi
    temp = data(j)
    For k = j - inc To jLo Step -inc
      If data(k) <= temp Then Exit For
      data(k + inc) = data(k)
    Next k
    data(k + inc) = temp
  Next j
Next i

For j = jLo + 1& To jHi
  If data(j - 1&) > data(j) Then Stop
Next j
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


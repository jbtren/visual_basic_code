VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Permutations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "This class supplies a series of arrays of N values that are all the permutations of { 0, 1, 2, ... , N - 1 }. They can be used as index values into an array, to get the array items in all permuted orders. Devised and coded by John Trenholme."
'
'###############################################################################
'#
'# Visual Basic 6 (or VBA) class file "Permutations.cls".
'#
'# John Trenholme - started 2 Feb 2005
'#
'###############################################################################

Option Base 0
Option Explicit
Option Compare Binary

Private Const c_version As String = "2005-03-09"  ' set to date of last update

'----- Introduction -----
'
' This class supplies a series of arrays of N index values. Each index array
' supplies a unique permutation of the values { 0, 1, 2, ... , N - 1 }.
' For example, if N = 3, you will get the arrays:
' { 0, 1, 2 }, { 1, 0, 2 }, { 2, 0, 1 }, { 0, 2, 1 }, { 1, 2, 0 }, { 2, 1, 0 }
'
' That is to say, the first index array will have i(0) = 0, i(1) = 1 and
' i(2) = 2. The second will have i(0) = 1, i(1) = 0 and i(1) = 2, and
' so on.
'
' The method used is that of Brian Heap (Comp. J. 6 (1963), 293-294). It causes
' the values with low index values in the returned arrays, such as i(0) and
' i(1), to vary most rapidly, while values at higher index values vary more and
' more slowly.
'
' You use the index values to access the items in an array by, e.g., X(i(j)).
' You will access them in a different permuted order for each new i() array.
'
' The returned index arrays are zero-based. If your array is 1-based, add 1 to
' the returned index values before using them, and so forth. If you want the
' index arrays to have a base other than 0, set the Property IndexArrayBase.
'
' Although the value of N is a Long, it is impractical to use very large values
' for N. The number of permutations of N items is equal to N!. This number
' rapidly becomes very large. For example, we have the following:
'
'   10! = 3,628,800
'   13! = 6,227,020,800
'   16! = 20,922,789,888,000
'   19! = 121,645,100,408,832,000
'
' At 10 million cases per second, it will take 1/3 second to process 10! items,
' 10 minutes to process 13! items, 24 days to process 16! items, and 385 years
' to process 19! items.

'----- Usage -----
'
' Declare a size and an array to hold the items you wish to access in permuted
' order.
'   Dim m as Long
'   m = theSize
'   Dim s(0& To m - 1&) As Something
'
' Fill the permuted-order-item array
'
'   For j = LBound(s) To UBound(ms)  ' or use literals (0& to m - 1& here)
'     s(j) = Whatever
'   Next j
'
' Declare a dynamic array to hold the returned permuted-index arrays
'
'   Dim i() as Long  ' no dimension(s) in dynamic array declaration
'
' Declare an object of this class, and make the object

'   Dim pm As Permutations
'   Set pm = New Permutations
'
' Set the Size property equal to the number of items
'
'   pm.Size = m
'
' Inside a loop, get the sequence of index arrays and use them
'
'   Do While pm.MoreToCome
'     i = pm.IndexArray  ' assignment (re)allocates dynamic array
'     ' array i() has index values for this permutation
'     For j = LBound(i) To UBound(i)  ' or use literals (0& to m - 1& here)
'       < Act on s(i(j)) >
'     Next j
'   Loop
'
' Dispose of the object memory and the dynamic-array memory
'
'   Set pm = Nothing
'   Erase i

'---------- Module-scope variables ---------------------------------------------
Private m_changed As Boolean     ' input-has-changed flag
Private m_b As Long              ' index array base (default 0)
Private m_indices() As Long      ' local index array
Private m_moreToCome As Boolean  ' more-to-come flag
Private m_size As Long           ' number N of items to be permuted
Private m_swaps() As Long        ' factorial-counting which-swap array

'++++++++++ Event ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub Class_Initialize()
Attribute Class_Initialize.VB_Description = "Sets up for the default case Size = 2, IndexArrayBase = 0. You should probably change these to your own values."
' set the default values (setting any of these sets m_changed = True)
Me.IndexArrayBase = 0&
Me.Size = 2&  ' what is the sound of one item permuting?
End Sub

'++++++++++ Event ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub Class_Terminate()
Attribute Class_Terminate.VB_Description = "Releases memory allocated to storage of the object's local arrays. It's a good idea to set an object of this class to Nothing when you're done with it, so that this routine will be called."
Erase m_indices, m_swaps ' release dynamically-allocated array memory
End Sub

'******** Property *************************************************************
Property Get IndexArray() As Long()
Attribute IndexArray.VB_Description = "Returns the present index array, and sets up for the next one. Call repeatedly after MoreToCome is True. If the array that's supplied is the last one, the next will wrap to the first and MoreToCome will be set to False, so you can quit."
Attribute IndexArray.VB_UserMemId = 0
' Returns the present index array, and sets up for the next one.
' If the array that's supplied is the last one, the next will wrap to the first
' but MoreToCome will be set to False, so the caller can quit asking (see
' "Usage" section above for an example).
Dim j As Long

' if something is different, (re)initialize
If m_changed Then
  m_changed = False
  ReDim m_indices(m_b To m_b + m_size - 1&)  ' allocate index-array memory
  ReDim m_swaps(0& To m_size - 2&)           ' allocate which-swap array
  ' (re)fill array with initial indices { 0, 1, 2, ... , N - 1 }
  ' set swap array (factorial counting) to initial state
  For j = 0& To m_size - 2&
    m_indices(m_b + j) = j
    m_swaps(j) = 0&
  Next j
  m_indices(m_b + m_size - 1&) = m_size - 1&
End If

' set the return value to the present index array
IndexArray = m_indices

' set up to produce the next permuted-index array, using factorial counting
m_moreToCome = True  ' claim that we're not done, until proven otherwise

' produce the next permuted-index array

' roll over any values that carry
j = 0&
Do While m_swaps(j) > j
  m_swaps(j) = 0&
  j = j + 1&
  If j = m_size - 1& Then  ' carried off end - that's all, folks!
    m_changed = True  ' force a reset
    m_moreToCome = False
    Exit Sub
  End If
Loop

Dim k As Long  ' position to swap with j + 1
If (j And 1) Then  ' j is odd
  k = m_b
Else
  k = m_b + m_swaps(j)
End If

Dim temp As Long  ' holds value during swap
temp = m_indices(m_b + j + 1&)
 m_indices(m_b + j + 1&) = m_indices(k)
m_indices(k) = temp

' finally, increment the value that won't carry
m_swaps(j) = m_swaps(j) + 1&
End Property

'******** Property *************************************************************
Property Get IndexArrayBase() As Long
Attribute IndexArrayBase.VB_Description = "Unless you set this to some other base value, the returned array is zero-based."
' The base index of the index array (default 0). Index array values are in
' elements i(base), i(base + 1), ... , i(base + K - 1).
IndexArrayBase = m_b
End Property

Property Let IndexArrayBase(ByVal rhs As Long)
m_changed = True
m_moreToCome = True
m_b = rhs
End Property

'******** Property *************************************************************
Property Get MoreToCome() As Boolean
Attribute MoreToCome.VB_Description = "True if the index array you are about to get is not the last one; False if it is. Sample code: Do While obj.MoreToCome {your statements} Loop."
' Returns True if there are one or more more additional index arrays available.
' Returns False if the array that's about to be supplied is "beyond" the last
' and so has wrapped around to the first again.
MoreToCome = m_moreToCome
End Property

'******** Property *************************************************************
Property Get Size() As Long
Attribute Size.VB_Description = "The number of items in the set that the permutations will be drawn from. Silently forced >= 2. Put the set items in an array such as X(0 To Size - 1). Get index arrays i() from IndexArray. Then access X(i(j)) for all j."
' The number of items to be permuted. Put those items in an array such as
' X(0 To Size - 1). Get an index array i(). Then access X(i(j)) for all j.
Size = m_size
End Property

Property Let Size(ByVal rhs As Long)
m_changed = True
m_moreToCome = True
If rhs < 2& Then rhs = 2&  ' silently fix up absurd input
m_size = rhs
End Property

'******** Property *************************************************************
Property Get Version() As String
Attribute Version.VB_Description = "Latest revision date of this class as a string in the format ""YYYY-MM-DD""."
' Latest revision date of this class as a string ("YYYY-MM-DD").
Version = c_version
End Property

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

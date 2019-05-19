VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "K_subsets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "This class supplies a series of arrays of K index values into an array X of M items (1 <= K <= M). Each index array defines a unique subset of the items in X. Devised and coded by John Trenholme."
'
'###############################################################################
'#
'# Visual Basic 6 (or VBA) class file "K_subsets.cls".
'#
'# John Trenholme - started 31 Jan 2005
'#
'###############################################################################

Option Base 0
Option Explicit
Option Compare Binary

Private Const c_version As String = "2005-03-09"  ' set to date of last update

'----- Introduction -----
'
' This class supplies a series of arrays of K index values into an array X of M
' items (K <= M). Each index array defines a unique subset of the items in X.
' The arrays are supplied in lexicographic order. The index values run from 0
' to M - 1. For example, if M = 4 and K = 2, you will get the arrays:
'   { 0, 1}, { 0, 2}, { 0, 3}, { 1, 2}, { 1, 3}, { 2, 3}
'
' That is to say, the first index array will have i(0) = 0 and i(1) = 1.
' The second will have i(0) = 0 and i(1) = 2, and so on.
'
' The method used is mixed-radix counting, with carries processed so that index
' values are always monotonically increasing and therefore unique.
'
' You access the items in the set by, e.g., X(i(j)).
'
' Note that you do not get permutations of the index values (such as { 1, 0})
' so if those are needed you must permute them yourself.
'
' Both the returned index arrays, and the item array X, are assumed to be
' zero-based. If your array of M items is 1-based, add 1 to the returned
' index values before using them, and so forth. If you want the index array to
' have a base other than 0, set the Property IndexArrayBase.
'
' Although the set and subset sizes are Longs, it is impractical to use very
' large values for these quantities. The number of subsets of size K, taken
' from a set of size M, is given by:
'
'   M! / ( K! * (M - K)!)
'
' This number rapidly becomes very large. For example, the following M,K pairs
' give about a trillion (1E+12) subsets:
'
'   1E+12,1  1.41E+6,2  18172,3  2215,4  656,5  302,6  178,7  123,8  93,9
'   76,10  66,11  59,12  54,13  50,14  48,15  43,22
'
' Some other values: 41,10 -> 1E+9  148,10 -> 1E+15  290,10 -> 1E+18
'
' At 30 million cases per second, it takes 9 hours to examine 1E+12 cases, and
' a year to examine 1E+15 cases. Don't even consider doing 1E+18 cases.

'----- Usage -----
'
' Declare a size and an array to hold the set items
'
'   Dim m as Long
'   m = theSize
'   Dim s(0& To m - 1&) As Something
'
' Fill the set-item array
'
'   For j = LBound(s) To UBound(ms)  ' or use literals (0& to m - 1& here)
'     s(j) = Whatever
'   Next j
'
' Declare a dynamic array to hold the subset index arrays
'
'   Dim i() as Long  ' no dimension(s) in dynamic array declaration
'
' Declare an object of this class, and make the object

'   Dim ks As K_subsets
'   Set ks = New K_subsets
'
' Set the sizes of the set and subset
'
'   ks.Size = m
'   ks.SubSize = k
'
' Inside a loop, get the sequence of index arrays and use them
'
'   Do While ks.MoreToCome
'     i = ks.IndexArray  ' assignment (re)allocates dynamic array
'     ' array i() now has index values for this subset, plus bounds values
'     For j = LBound(i) To UBound(i)  ' or use literals (0& to k - 1& here)
'       < Act on s(i(j)) >
'     Next j
'   Loop
'
' Dispose the object memory and the dynamic-array memory
'
'   Set ks = Nothing
'   Erase i

'---------- Module-scope variables ---------------------------------------------
Private m_changed As Boolean     ' input-has-changed flag
Private m_b As Long              ' index array base (default 0)
Private m_indices() As Long      ' local index array
Private m_moreToCome As Boolean  ' more-to-come flag
Private m_size As Long           ' number M of items in set
Private m_subSize As Long        ' number K of items in each subset

'++++++++++ Event ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub Class_Initialize()
Attribute Class_Initialize.VB_Description = "Sets up for the default case Size = 1, SubSize = 1. You should probably change these to your own values."
' set some default values (setting any of these sets m_changed = True)
Me.IndexArrayBase = 0&
Me.Size = 1&  ' smallest sensible value
Me.SubSize = 1&  ' only possible value when Size = 1
End Sub

'++++++++++ Event ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub Class_Terminate()
Attribute Class_Terminate.VB_Description = "Releases memory allocated to storage of the object's local index array, so it's a good idea to set an object of this class to Nothing when you're done with it."
Erase m_indices  ' release dynamically-allocated index-array memory
End Sub

'******** Property *************************************************************
Property Get IndexArray() As Long()
Attribute IndexArray.VB_Description = "Returns the present index array, and sets up for the next one. Call repeatedly while MoreToCome is True. If the array that's supplied is the last one, the next will wrap to the first and MoreToCome will be set to False, so you can quit."
Attribute IndexArray.VB_UserMemId = 0
' Returns the present index array, and sets up for the next one.
' If the array that's supplied is the last one, the next will wrap to the first
' but MoreToCome will be set to False, so the caller can quit asking (see
' "Usage" section above for an example).
Dim j As Long

' if something is different, (re)initialize
If m_changed Then
  m_changed = False
  If m_subSize > m_size Then m_subSize = m_size  ' silently impose sanity
  ReDim m_indices(m_b To m_b + m_subSize - 1&)  ' allocate index-array memory
  ' (re)fill array with initial subset { 0, 1, 2, ... , K - 1 }
  For j = 0& To m_subSize - 1&
    m_indices(m_b + j) = j
  Next j
End If

' set the return value to the present index array
IndexArray = m_indices

' set up to produce the next index array, using mixed-radix counting
m_moreToCome = True  ' claim that we're not done, until proven otherwise
j = 0&  ' point to right of last index value

' produce the next index array
Do
  If j = m_subSize Then  ' we have carried beyond the leftmost index value
    m_moreToCome = False    ' set up to roll over & restart sequence of arrays
    Exit Do
  End If
  j = j + 1&  ' step left to next index value
  ' increment this index value
  m_indices(m_b + m_subSize - j) = m_indices(m_b + m_subSize - j) + 1&
' test the inremented value against the maximum possible value at this location
' carry if it overflowed, by repeating the loop with j moved one position left
Loop While m_indices(m_b + m_subSize - j) > m_size - j
' if this is the last subset, set up to roll over to the first
If Not m_moreToCome Then m_indices(m_b) = 0&
' if there are any less-significant digits, set them to their lowest values
' this is a do-nothing loop if only the last index was incremented (no carries)
For j = m_subSize - j + 1& To m_subSize - 1&
  m_indices(m_b + j) = m_indices(m_b + j - 1&) + 1&
Next j
End Property

Property Let IndexArray(rhs() As Long)
' This allows the user to start at an arbitrary point in the sequence.
' Be sure that you supply a valid subset index array! This is assured if you
' use a saved array that was supplied by the class.
m_indices = rhs
Me.IndexArrayBase = LBound(rhs)
Me.SubSize = UBound(rhs) - LBound(rhs) + 1&
If m_size < m_subSize Then m_size = m_subSize  ' silently impose sanity
m_changed = False  ' in this case, we don't want a restart
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
Attribute Size.VB_Description = "The number of items in the set that the subsets will be drawn from. Silently forced >= 1. Put the set items in an array such as X(0 To Size - 1). Get an index array i(). Then access X(i(j)) for all j."
' The number of items in the set that the subsets will be drawn from. Put those
' items in an array such as X(0 To Size - 1). Get an index array i(). Then
' access X(i(j)) for all j.
Size = m_size
End Property

Property Let Size(ByVal rhs As Long)
m_changed = True
m_moreToCome = True
If rhs < 1& Then rhs = 1&  ' silently fix up absurd input
m_size = rhs
End Property

'******** Property *************************************************************
Property Get SubSize() As Long
Attribute SubSize.VB_Description = "The number of items in the subsets. Silently forced >= 1 and <= Size.   Each returned IndexArray will contain SubSize index values, arranged in strictly increasing order. There will be Size!/SubSize!/(Size-SubSize)! IndexArrays."
' The number of items in the subsets.  Must be <= Size. Each index array will
' contain SubSize index values, arranged in strictly increasing order. There
' will be Size!/SubSize!/(Size-SubSize)! index arrays.
SubSize = m_subSize
End Property

Property Let SubSize(ByVal rhs As Long)
m_changed = True
m_moreToCome = True
If rhs < 1& Then rhs = 1&  ' silently fix up absurd input
m_subSize = rhs
End Property

'******** Property *************************************************************
Property Get Version() As String
Attribute Version.VB_Description = "Latest revision date of this class as a string in the format ""YYYY-MM-DD""."
' Latest revision date of this class as a string ("YYYY-MM-DD").
Version = c_version
End Property

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

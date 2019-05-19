VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NiceNumbers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Supplies a sequence of major and minor ""nice"" numbers (i.e. that are 1, 2 or 5 times a power of 10) based on supplied interval from ""First"" to ""Last"" as extended by ""Beyond."" To use, create an object of this class, set the values of First, Last, and (optionally) AtLeast, SubIntervals and Beyond. Results are in Major and Minor, which return arrays of Doubles. To get ""nice"" multiples of (say) Pi, divide the ""First"" and ""Last"" values by Pi before calling, and multiply the return values by Pi. Devised and coded by John Trenholme."
'
'###############################################################################
'#  _   _                   _   _                    _                         #
'# ( ) ( ) _               ( ) ( )                  ( )                        #
'# | `\| |(_)   ___    __  | `\| | _   _   ___ ___  | |_      __   _ __   ___  #
'# | , ` || | /'___) /'__`\| , ` |( ) ( )/' _ ` _ `\| '_`\  /'__`\( '__)/',__) #
'# | |`\ || |( (___ (  ___/| |`\ || (_) || ( ) ( ) || |_) )(  ___/| |   \__, \ #
'# (_) (_)(_)`\____)`\____)(_) (_)`\___/'(_) (_) (_)(_,__/'`\____)(_)   (____/ #
'#                                                                             #
'###############################################################################
'#
'# Visual Basic (VBA or VB6) class file "NiceNumbers.cls"
'#
'# Devised and coded by John Trenholme - started 6 Aug 2003
'#
'###############################################################################
'#
'# This class calculates and dispenses "nice" numbers in the numeric interval
'# between First and Last, as extended by Beyond. Nice numbers are 1, 2 or 5
'# times an integer power of 10. Nice numbers are useful for graph axis
'# ticks and labels, or for contour values, or anywhere human-friendly numbers
'# are desired.
'#
'# You will get back two Variants holding arrays of nice values. The Major array
'# will contain at least AtLeast values, and up to three times that many,
'# depending on the particular values of First and Last you supply. The Minor
'# array will contain nice numbers between, and perhaps beyond (but not
'# including), the Major values; you can control how many Minor values per Major
'# interval with SubIntervals, but it's best to leave SubIntervals at the
'# default value of 0, which tells the object to select the better of 4 or 5.
'#
'# To get "nice" multiples of (say) Pi, divide the "First" and "Last" values by
'# Pi before calling, and multiply the return values by Pi. And so forth.
'#
'# Normally, the Major and Minor values returned will all be within the
'# supplied interval, or just at its ends (as extended by Beyond), and there
'# may be Minor values beyond the end-most Major values. This is called "tight"
'# positioning. However, if Loose is set to True the positioning is "loose" and
'# the returned values will start and stop on a Major value, with vaues being
'# added at one or both ends, if necessary, to make that happen. This is, for
'# example, for use with programs (such as Excel) that make ugly axis-label
'# values unless the plot starts at a Major value.
'#
'# If x.Last is less than x.First, the results will be supplied in decreasing
'# order, rather than the usual increasing order.
'#
'# Note: the interval should be larger than 1.0E-290 or so, and the interval
'# should be more than 1.0E-13 times the mean value of the ends, or near-
'# underflow will lead to irregular and incorrect output intervals.
'#
'# Note: if it is important that the returned major and minor values are exact
'# multiples of 1, 2 or 5 times a power of 10 (perhaps you want to print them),
'# you may use the DigitRound function as indicated below. This will clean up
'# the last few bits, which may be wrong due to roundoff - but it's slower.

' Usage example:
'
'   Dim niceX As NiceNumbers
'   Set niceX = New NiceNumbers
'   niceX.First = 3.5        ' start of interval
'   niceX.Last = 12.2        ' end of interval
'   niceX.AtLeast = 2#       ' will get at least this many major values
'   niceX.Beyond = 0.001     ' fraction beyond ends; this is the default value
'   niceX.SubIntervals = 0&  ' let object find the best value (the default)
'
'   Dim xMajors As Variant
'   xMajors = niceX.Majors   ' get major values
'   Dim j as Long
'   Dim xMajor as Variant
'   For Each xMajor In xMajors
'     ' optionally, xMajor = niceX.DigitRound(xMajor, 14)
'     '... do something with a major value in "xMajor" ...
'   Next xMajor
'
'   ' If you want to use minor values (perhaps for axis minor ticks) then...
'   Dim xMinors As Variant
'   xMinors = niceX.Minors   ' get minor values
'   Dim xMinor as Variant
'   For Each xMinor In xMinors
'     ' optionally, xMinor = niceX.DigitRound(xMinor, 14)
'     '... do something with a minor value in "xMinor" ...
'   Next xMinor
'
'   Set niceX = Nothing      ' recover memory used by nice-number object

Option Base 1          ' array base value when not specified - the default
Option Compare Binary  ' string comparison is based on Asc(char) - the default
Option Explicit        ' forces explicit variable declaration - changes default

' This is returned by Property .Version - set to date of last update
Private Const Version_c As String = "2013-02-05"
Private Const c_File As String = "NiceNumbers.cls[" & Version_c & "]."

'*******************************************************************************
'*
'* Data members
'*
'*******************************************************************************

' input data saved in this object
Private m_atLeast As Double   ' produce at least this many major values
Private m_beyond As Double    ' fraction of interval to go beyond supplied ends
Private m_first As Double     ' first number in numeric interval
Private m_last As Double      ' last number in numeric interval
Private m_loose As Boolean    ' True -> include nearest major values beyond ends
Private m_subInts As Long     ' user-spec'd minor intervals per major interval

' internal values
Private m_baseIndex As Long   ' lowest index on all returned arrays (default 1)
Private m_bigStep As Double   ' size of the major step
Private m_changed As Boolean  ' flag indicates that some input value has changed
Private m_jMajor As Long      ' number of major values
Private m_jMinor As Long      ' number of minor values
Private m_pwr10 As Long       ' power of 10 that scales 1 <= |nice| < 10
Private m_sign As Double      ' sign of steps (i.e., sign of Last - First)
Private m_stepsUse As Long    ' actual minor sub-intervals per major interval
Private m_subMinor As Long    ' sub-minor intervals per minor interval

' arrays - kept as Variants so they can be returned from a class
Private m_majors As Variant   ' array of major values for return
Private m_minors As Variant   ' array of minor values for return
'

'*******************************************************************************
'*
'* Property Get and Let routines
'*
'*******************************************************************************

'===============================================================================
Public Property Get AtLeast() As Double
Attribute AtLeast.VB_Description = "You get at least this many major values in the interval from First to Last (as extended by Beyond), and up to 3 times as many as this. If you try to set this below 1, it will be set to 1. You probably should use a value of 2 or greater. Defaults to 3."
' You get at least this many major values in the interval from First to Last
' (as extended by Beyond), and up to 3 times as many as this. If you try to set
' this below 1, it will be set to 1. You probably should use a value of 2 or
' greater. Defaults to 3.
AtLeast = m_atLeast
End Property

Public Property Let AtLeast(ByVal MinMajorValueCount As Double)
' Protect against user insanity; insure that we return at least one major value.
' It is silly to set this less than 2, but we let the user do that.
If MinMajorValueCount < 1# Then MinMajorValueCount = 1#
' Because of the cost of recalculation, changes to any input parameter just set
' a flag, and recalculation is triggered only when output is requested and one
' or more input parameters have changed.
If m_atLeast <> MinMajorValueCount Then m_changed = True
m_atLeast = MinMajorValueCount
End Property

'===============================================================================
Public Property Get BaseIndex() As Long
Attribute BaseIndex.VB_Description = "First (i.e., lowest) index value in all the returned arrays. Default = 1"
' First (i.e., lowest) index value in all the returned arrays. Default = 1.
BaseIndex = m_baseIndex
End Property

Public Property Let BaseIndex(ByVal arraysFirstIndex As Long)
' Set the input-changed flag if it has changed, but don't calculate anything yet
If m_baseIndex <> arraysFirstIndex Then m_changed = True
m_baseIndex = arraysFirstIndex
End Property

'===============================================================================
Public Property Get Beyond() As Double
Attribute Beyond.VB_Description = "Your interval will be extended beyond the specified ""First"" and ""Last"" points by this fraction of the length of the interval. Set to 0.001 (the default) or so to pick up values just slightly outside the interval. Limited to be > -0.499."
' Your interval will be extended beyond the specified "First" and "Last" points
'  by this fraction of the length of the interval. Set to 0.001 (the default)
' or so to pick up values just slightly outside the interval. This avoids the
' problem where a value is just missed at one end or the other of a plot.
' Limited to be > -0.499 to avoid switching ends.
Beyond = m_beyond
End Property

Public Property Let Beyond(ByVal FractionPastEnds As Double)
' Protect against a value so large and negative that it reverses First & Last
If FractionPastEnds < -0.499 Then FractionPastEnds = -0.499
' Set the input-changed flag if it has changed, but don't calculate anything yet
If m_beyond <> FractionPastEnds Then m_changed = True
m_beyond = FractionPastEnds
End Property

'===============================================================================
Public Property Get First() As Double
Attribute First.VB_Description = "The nominal start of the interval containing the ""nice"" numbers. See ""Last"" and ""Beyond."" If greater than ""Last,"" values will be returned in reversed order (larger to smaller). Defaults to 0."
' The nominal start of the interval containing the "nice" numbers. See "Last"
' and "Beyond." If greater than "Last," values will be returned in reversed
' order (larger to smaller). Defaults to 0.
First = m_first
End Property

Public Property Let First(ByVal intervalStart As Double)
' Set the input-changed flag if it has changed, but don't calculate anything yet
If m_first <> intervalStart Then m_changed = True
m_first = intervalStart
End Property

'===============================================================================
Public Property Get Last() As Double
Attribute Last.VB_Description = "The nominal end of the interval containing the ""nice"" numbers. See ""First"" and ""Beyond."" If less than ""First,"" values will be returned in reversed order (larger to smaller). Defaults to 1."
' The nominal end of the interval containing the "nice" numbers. See "First"
' and "Beyond." If less than "First," values will be returned in reversed
' order (larger to smaller). Defaults to 1.
Last = m_last
End Property

Public Property Let Last(ByVal intervalEnd As Double)
' Set the input-changed flag if it has changed, but don't calculate anything yet
If m_last <> intervalEnd Then m_changed = True
m_last = intervalEnd
End Property

'===============================================================================
Public Property Get Loose() As Boolean
Attribute Loose.VB_Description = "If False (the default) the values are ""tight"" and do not include major values. If True, the values are ""loose"" and include values out to the first major increment beyond the ""First"" and ""Last"" values, if those values do not lie exactly (within ""Beyond"") on a major value."
' If False (the default) the values are "tight" and do not include major values
' beyond the supplied interval. In plotting, this is used if you want the plot
' to extend right up to the end of the plotting area. If Loose is True, the
' values are "loose" and include values out to the first major increment beyond
' the "First" and "Last" values, presuming those values do not lie exactly
' (within "Beyond") on a major value. In plotting, this is used if you want
' major values at the ends of the plot, or if you are talking to a program
' (such as Excel) that is not happy unless plots start on major values.
Last = m_loose
End Property

Public Property Let Loose(ByVal ExtendToMajor As Boolean)
' Set the input-changed flag if it has changed, but don't calculate anything yet
If m_loose <> ExtendToMajor Then m_changed = True
m_loose = ExtendToMajor
End Property

'===============================================================================
Public Property Get MajorCount() As Double
Attribute MajorCount.VB_Description = "Number of major values within the interval from ""First"" to ""Last,"" as extended by ""Beyond"" and ""Loose."""
' Number of major values within the interval from "First" to "Last," as
' extended by "Beyond" and "Loose."
If m_changed Then calculate  ' calculated result; input change -> recalculation
MajorCount = m_jMajor
End Property

'===============================================================================
Public Property Get Majors() As Variant
Attribute Majors.VB_Description = "Array of major values within the interval, in a Variant that holds an array of Doubles. Do ""Dim x as Variant: x = nn.Majors"" and use ""MajorCount"" values starting at ""BaseIndex"" (default 1)."
' Major "nice" numbers, as an array of Doubles inside a Variant. This is one of
' the two calculated result arrays. Note that this array is 1-based unless
' "BaseIndex" is changed by the user.
If m_changed Then calculate  ' calculated result; input change -> recalculation
Majors = m_majors
End Property

'===============================================================================
Public Property Get MajorStep() As Double
Attribute MajorStep.VB_Description = "Increment between values in the ""Majors"" array. If ""First"" > ""Last,"" it will be negative."
' The increment between values in the "Majors" array. If "First" > "Last," it
' will be negative.
If m_changed Then calculate  ' calculated result; input change -> recalculation
MajorStep = m_bigStep * m_sign
End Property

'===============================================================================
Public Property Get MinorCount() As Double
Attribute MinorCount.VB_Description = "Number of minor values within the interval from ""First"" to ""Last,"" as extended by ""Beyond"" and ""Loose."" Does not count major values."
' Number of minor values within the interval from "First" to "Last," as
'  extended by "Beyond" and "Loose." Does not count major values.
If m_changed Then calculate  ' calculated result; input change -> recalculation
MinorCount = m_jMinor
End Property

'===============================================================================
Public Property Get Minors() As Variant
Attribute Minors.VB_Description = "Array of minor values within the interval, in a Variant that holds an array of Doubles. Do ""Dim y as Variant: y = nn.Minors"" and use ""MinorCount"" values starting at ""BaseIndex"" (default 1). Does not include major values."
' Minor "nice" numbers, as an array of Doubles inside a Variant. This is one of
' the two calculated result arrays. Note that this array is 1-based unless
' "BaseIndex" is changed by the user. This will contain only the minor values,
' without any major values.
If m_changed Then calculate  ' calculated result; input change -> recalculation
Minors = m_minors
End Property

'===============================================================================
Public Property Get MinorStep() As Double
Attribute MinorStep.VB_Description = "Increment between minor values. If ""First"" > ""Last,"" it will be negative."
' The increment between values in the "Minors" array. If "First" > "Last," it
' will be negative.
If m_changed Then calculate  ' calculated result; input change -> recalculation
MinorStep = m_bigStep * m_sign / m_stepsUse
End Property

'===============================================================================
Public Property Get Power10() As Long
' Power of 10 that scales largest major or minor nice number to 1 <= |nn| < 10.
' That is, in all cases 1 <= |nnMax| / 10# ^ Power10 < 10
' To be able to print numbers 1 <= |number| < 10 from returned nice numbers
' divide them all by 10 ^ Power10 and then print "times 10^Power10" or similar
' nearby.
If m_changed Then calculate  ' calculated result; input change -> recalculation
Power10 = m_pwr10
End Property

'===============================================================================
Public Property Get SubIntervals() As Long
Attribute SubIntervals.VB_Description = "Number of minor steps included in one major step. If you set this to 0 (the default value) an appropriate value will be calculated and used, depending on the major step (1,2,5 -> 5,4,5). If you specify a value > 1, that many minor steps will always be used. The return value is equal to the number actually used. Input values < 2 will be silently set to 0."
' Number of minor steps included in one major step. If you set this to 0 (the
' default value) an appropriate value will be calculated and used, depending on
' the major step (1,2,5 -> 5,4,5). If you specify a value > 1, that many minor
' steps will always be used. The return value is equal to the number actually
' used. Input values < 2 will be silently set to 0.
If m_changed Then calculate  ' calculated result; input change -> recalculation
SubIntervals = m_stepsUse
End Property

Public Property Let SubIntervals(ByVal SubIntervalCount As Long)
' Negative values, or unity, cause trouble, so silently set them to 0.
If SubIntervalCount < 2& Then SubIntervalCount = 0&
' Set the input-changed flag if it has changed, but don't calculate anything yet
If m_subInts <> SubIntervalCount Then m_changed = True
m_subInts = SubIntervalCount
End Property

'===============================================================================
Public Property Get SubSubIntervals() As Long
Attribute SubSubIntervals.VB_Description = "Number of sub-minor steps included in one minor step, depending on the major step (1,2,5 -> 5,5,4). If you want sub-sub intervals, calculate them using this and ""MinorStep."" Value based on default ""SubIntervals"" - use only if ""SubIntervals"" = 0."
' Number of sub-minor steps included in one minor step, depending on the major
' step (1,2,5 -> 5,5,4). If you want sub-sub intervals, calculate them using
' this and "MinorStep." Value based on default "SubIntervals" - use only if
' "SubIntervals" = 0.
If m_changed Then calculate  ' calculated result; input change -> recalculation
SubSubIntervals = m_subMinor
End Property

'===============================================================================
Public Property Get Version() As String
Attribute Version.VB_Description = "The date of the latest revision to this class; string in the format ""YYYY-MM-DD."""
' Date of the latest revision to this Class; string in the format "yyyy-mm-dd".
Version = Version_c
End Property

'*******************************************************************************
'*
'* Event handlers
'*
'*******************************************************************************

'-------------------------------------------------------------------------------
Private Sub Class_Initialize()
Attribute Class_Initialize.VB_Description = "Internal routine that sets up some default input parameter values (BaseIndex = 1, First = 0, Last = 1, AtLeast = 4, SubIntervals = 0, Beyond = 0.001, Loose = False), just in case you request output without setting any input. But why would you do that? Better to, at least, set ""First"" and ""Last"" before requesting output."
' We make up nominal values, so requests for nice numbers from a newly-created
' object will supply something reasonable even if the user has not set all the
' input values. The values for AtLeast, SubIntervals and Beyond are sensible,
' so the user can just set First and Last and get acceptable results.
m_baseIndex = 1&
m_first = 0#
m_last = 1#
m_atLeast = 3#
m_subInts = 0&  ' -> let class make a sensible choice
m_beyond = 0.001
m_loose = False
' set flag so next output request will calculate results based on these values
m_changed = True
End Sub

'-------------------------------------------------------------------------------
Private Sub Class_Terminate()
' do nothing
End Sub

'*******************************************************************************
'*
'* Procedures
'*
'*******************************************************************************

'-------------------------------------------------------------------------------
Private Sub calculate()
Attribute calculate.VB_Description = "Internal routine called to calculate new values if any input parameter has changed since the last request for ""Majors"" or ""Minors"" or other calculated quantities."
' Calculate the arrays of major and minor "nice" values, and other quantities.
' This has to be called only if input data has changed.
Const ID_C As String = c_File & "calculate Sub"
If Not Me.InDesign() Then On Error GoTo ErrorHandler  ' compiled code -> Error

' remember initial-order sign (zero interval will be fixed up later)
If m_first <= m_last Then
  m_sign = 1#
Else
  m_sign = -1#
End If

' add on the specified extra space beyond the ends (could be negative)
' this slightly larger interval value catches almost-there values
Dim add As Double
add = m_beyond * (m_last - m_first)

' make local values with endpoints in numeric order xLo <= xHi
Dim xLo As Double
xLo = m_sign * (m_first - add)  ' low end, pushed lower by Beyond
Dim xHi As Double
xHi = m_sign * (m_last + add)  ' high end, pushed higher by Beyond

' protect against minor intervals below roundoff (silently)
Const c_tiny As Double = 2.23E-16  ' 2^(-52)
Dim intervalMin As Double
If m_subInts = 0& Then intervalMin = 4.5 Else intervalMin = m_subInts
intervalMin = c_tiny * (Abs(xLo) + Abs(xHi)) * m_atLeast * intervalMin
' fixup in case xLo = xHi = 0 (or close to it)
If intervalMin < 1E-300 Then intervalMin = 1E-300
' push ends out if they are too close
If xHi - xLo < intervalMin Then
  xLo = 0.5 * (xLo + xHi - intervalMin)
  xHi = xLo + intervalMin
End If

' get the unscaled tentative interval size
Dim trialSize As Double
trialSize = (xHi - xLo) / m_atLeast ' note AtLeast >= 1.0, so this is >= 0
m_bigStep = trialSize

' increase or reduce interval size to the range 1.0 <= m_bigStep < 10.0
' we use a scale factor that is an exact multiple of 10 to avoid use of 0.1
Dim scal
scal = 1#                ' scale factor (a power of 10)
Do While m_bigStep < 1#  ' increase m_bigStep if < 1
  scal = scal * 10#      ' actually inverse of scale factor; see below
  m_bigStep = trialSize * scal
Loop
' if m_bigStep was < 1, it is now >= 1 but < 10, and scal > 1
If 1# = scal Then  ' m_bigStep started >= 1; we may have to reduce it
  Do While m_bigStep >= 10#  ' decrease m_bigStep if too large
    scal = scal * 10#
    m_bigStep = trialSize / scal
  Loop
Else  ' m_bigStep started < 1, so scal > 1 and we need to invert it
  scal = 1# / scal  ' only divide by 10^N once; that's not exact in floats
End If
' we now have trialSize = m_bigStep * scal, with 1 <= m_bigStep < 10
'Debug.Assert Abs(m_bigStep * scal / trialSize - 1#) < 3.331E-16
'Debug.Assert m_bigStep >= 1#
'Debug.Assert m_bigStep < 10#

' get nice interval size (1.0, 2.0, 5.0), nice minor interval count (5, 4, 5)
' and also nice sub-minor interval count (5, 5, 4)
Dim nSub As Long
If m_bigStep >= 5! Then      ' 5.0 <= m_bigStep < 10.0
  m_bigStep = 5! * scal
  nSub = 5&                  ' sub is 1.0
  m_subMinor = 5&            ' subSub is 0.2
ElseIf m_bigStep >= 2! Then  ' 2.0 <= m_bigStep < 5.0
  m_bigStep = 2! * scal
  nSub = 4&                  ' sub is 0.5
  m_subMinor = 5&            ' subSub is 0.1
Else                         ' 1.0 <= m_bigStep < 2.0
  m_bigStep = scal
  nSub = 5&                  ' sub is 0.2
  m_subMinor = 4&            ' subSub is 0.05
End If

' set number of minor intervals per major interval
If m_subInts = 0& Then    ' special value of 0 is flag: use calculated value
  m_stepsUse = nSub       ' caller trusts us to do the right thing
Else
  m_stepsUse = m_subInts  ' caller has their own opinion; respect it (value > 1)
  m_subMinor = 0&         ' we don't know what to do here, so punt
End If

' take off an integer number of major intervals near center of interval
' keep the result in a Double to get 53-bit accuracy (Longs are too small)
' if this is not done, jNice and jMinor (below) can overflow
Dim intAvg As Double
intAvg = Int((0.5 * xLo + 0.5 * xHi) / m_bigStep)
Dim xAvg As Double
xAvg = intAvg * m_bigStep
xLo = xLo - xAvg  ' remember that xLo and xHi have been pushed out by Beyond
xHi = xHi - xAvg

' get an integer number of major intervals, above the lower end
Dim jNice As Long
jNice = Int(xLo / m_bigStep) + 2#

Dim jSubInt As Long                    ' modulo counter of minor intervals
jSubInt = 0&                           ' modulo is zero at major interval
Dim subStep As Double
subStep = m_bigStep / m_stepsUse       ' minor step value
Dim jMinor As Long
If m_loose Then  ' user wants "loose" positioning
  ' move to major value at or just below lower end of interval
  Do While m_bigStep * jNice > xLo
    jNice = jNice - 1#
  Loop
  jMinor = jNice * m_stepsUse          ' count of minor steps
Else  ' user wants "tight" positioning
  ' move to major value just below lower end of interval
  Do While m_bigStep * jNice >= xLo
    jNice = jNice - 1#
  Loop
  ' move up by minor intervals until within interval (including lower end)
  ' keep modulo counter for minor values
  jMinor = jNice * m_stepsUse          ' count of minor steps
  Do While subStep * jMinor < xLo      ' keep stepping while below lower end
    jMinor = jMinor + 1#
    jSubInt = jSubInt + 1&             ' use modulo count for minor intervals
    If jSubInt >= m_stepsUse Then jSubInt = 0&
  Loop
End If

' calculate values and put into arrays
Dim xNice As Double
Dim jSpace As Long
' make sure arrays have enough space to hold results; will surely fit into this
jSpace = m_baseIndex + 3& + Int((xHi - xLo) / subStep)
ReDim m_majors(m_baseIndex To jSpace)  ' yes, you can ReDim a "scalar" Variant
ReDim m_minors(m_baseIndex To jSpace)
m_jMajor = m_baseIndex - 1&  ' major array index, less 1 because we pre-increment
m_jMinor = m_jMajor          ' minor array index, less 1 because we pre-increment
Dim valMax As Double  ' keep track of largest absolute value
valMax = 0#
Do
  xNice = subStep * jMinor      ' get a minor-step nice number
  ' we have gone beyond the end, & have also produced enough major values
  If (xNice > xHi) And (m_jMajor >= m_baseIndex + m_atLeast - 1#) Then
    If m_loose Then  ' user wants "loose" positioning
      If jSubInt = 1& Then Exit Do  ' quit only if we just did a major value
    Else  ' user wants "tight" positioning
      Exit Do  ' quit immediately, at either a minor or major value
    End If
  End If
  Dim xValue As Double
  xValue = xNice + xAvg  ' add the average value back on
  If jSubInt = 0& Then  ' this is a major interval
    m_jMajor = m_jMajor + 1&  ' pre-increment the index
    ' save as a major value, restoring sign info
    m_majors(m_jMajor) = xValue * m_sign
  Else                  ' this is a minor interval
    m_jMinor = m_jMinor + 1&  ' pre-increment the index
    ' save as a minor value, restoring sign info
    m_minors(m_jMinor) = xValue * m_sign
  End If
  ' find largest major or minor value, in absolute value
  If valMax < Abs(xValue) Then valMax = Abs(xValue)
  jMinor = jMinor + 1#  ' move ahead to next value
  jSubInt = jSubInt + 1&       ' use modulo count for sub-intervals
  If jSubInt >= m_stepsUse Then jSubInt = 0&
Loop

' find the power of 10 that scales max returned value to 1 <= |value| < 10
Dim valAdj As Double
valAdj = valMax
m_pwr10 = 0&  ' power of 10
scal = 1#     ' scale factor
Do While valAdj < 1#    ' increase valAdj if < 1
  m_pwr10 = m_pwr10 - 1&
  scal = scal * 10#
  valAdj = valMax * scal
Loop
' if valAdj was < 1, it is now >= 1 but < 10, and m_pwr10 > 1
If 0& = m_pwr10 Then  ' valAdj started >= 1; we may have to reduce it
  Do While valAdj >= 10#  ' decrease valAdj if too large
    m_pwr10 = m_pwr10 + 1&
    scal = scal * 10#
    valAdj = valMax / scal
  Loop
End If
'Debug.Assert Abs(valMax / 10# ^ m_pwr10 - 5.5) <= 4.50000000000001

' resize arrays so they are just large enough - then LBound & UBound work OK
If m_jMajor >= m_baseIndex Then
  ReDim Preserve m_majors(m_baseIndex To m_jMajor)
Else  ' there has been a serious logic error
  If Not Me.InDesign() Then On Error GoTo 0
  Err.Raise 20002, ID_C, _
    erMsg("no major value in interval", xLo, xHi, scal, subStep, xNice, ID_C)
End If
If m_jMinor >= m_baseIndex Then
  ReDim Preserve m_minors(m_baseIndex To m_jMinor)
Else  ' there has been a serious logic error
  If Not Me.InDesign() Then On Error GoTo 0
  Err.Raise 20003, ID_C, _
    erMsg("no minor value in interval", xLo, xHi, scal, subStep, xNice, ID_C)
End If

' calculated results are based on present input values, so clear change flag
m_changed = False
Exit Sub

ErrorHandler:
' we get here if compiled, and there was an error during calculation
Err.Raise 20001, ID_C, _
  erMsg("internal error: " & Error(Err.Number), _
    xLo, xHi, scal, subStep, xNice, ID_C)
End Sub

'===============================================================================
Public Function DigitRound( _
  ByVal valToRound As Double, _
  Optional ByVal numDigits As Integer = 6) _
As Double
' Round the input value to the specified number of digits. Used (among other
' things) to keep lengths short when printed. Returns input unchanged if it is
' zero. If numDigits < 1 it is set to 1; if numDigits > 15 it is set to 15.
' This routine takes about 2 microseconds on a 3 GHz Pentium 4.
' Note: this is the equivalent of a Class Static Function; VB lacks those
Const ID_C As String = c_File & "DigitRound Function"
On Error GoTo ErrHandler
If valToRound = 0# Then
  DigitRound = 0#  ' special case; simple to do, and causes Log(0) problem
Else
  If numDigits < 1 Then numDigits = 1  ' silently fix invalid argument values
  If numDigits > 15 Then numDigits = 15
  ' the following Const value is accurate to the last IEEE 754 bit
  Const Log10_e As Double = 0.43429448 + 1.903251828E-09  ' makes Log -> Log10
  Dim pow10 As Double
  pow10 = numDigits - 1# - Int(Log(Abs(valToRound)) * Log10_e)  ' digit shift
  Dim scaling As Double
  scaling = 10# ^ (0.5 * pow10)  ' avoid overflow for very small input values
  Dim scaled As Double  ' we will put digits we want into integer part of this
  scaled = (valToRound * scaling) * scaling
  DigitRound = (Int(scaled + 0.5) / scaling) / scaling
End If
Exit Function  '----------------------------------------------------------------

ErrHandler:
Dim errNum As Long, errDes As String
errNum = Err.Number: errDes = Err.Description
Err.Raise errNum, "DigitRound", errDes & vbLf & _
  "valToRound = " & valToRound & "    numDigits = " & numDigits & vbLf & _
  "Problem in " & ID_C
Resume  ' to allow debugging; set next statement here & single-step
End Function

'-------------------------------------------------------------------------------
Private Function erMsg( _
  ByRef why As String, _
  ByVal xLo As Double, _
  ByVal xHi As Double, _
  ByVal power As Double, _
  ByVal subStep As Double, _
  ByVal xNice As Double, _
  ByRef who As String) _
As String
erMsg = _
"*** NiceNumbers internal error - " & why & " ***" & _
vbNewLine & _
"Values in object (version " & Version_c & "):" & _
vbNewLine & _
"First: " & m_first & "  Last: " & m_last & "  Beyond: " & m_beyond & _
vbNewLine & _
"AtLeast: " & m_atLeast & "  SubIntervals: " & m_stepsUse & _
"  SubSubIntervals: " & m_subMinor & _
vbNewLine & _
"xLo: " & xLo & "  xHi: " & xHi & "  power: " & power & _
vbNewLine & _
"bigStep: " & m_bigStep & "  subStep: " & subStep & "  xNice: " & xNice & _
vbNewLine & _
"Problem in " & who
End Function

'===============================================================================
Public Function InDesign() As Boolean
' Return True if program is running in IDE (editor) design environment, and
' False if program is running as a standalone EXE that has been compiled by VB6.
' Useful for "hooking" only when standalone, or adjusting for the speed
' difference between compiled and interpreted. So in your program you can say:
' If [Not] inDesign() Then ...
'         John Trenholme - 2010-01-31
InDesign = False
On Error Resume Next  ' set to ignore error in Assert
Debug.Assert 1& \ 0&  ' attempts this illegal feat only in IDE
If 0& <> Err.Number Then
  InDesign = True  ' comment this out to get compiled behavior while in IDE
  Err.Clear  ' "Resume Next" passes Err state back up to caller (yes!) so Clear
End If
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

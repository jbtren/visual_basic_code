'===== gammaLog =============================================================
Public Function gammaLog( _
  ByVal x As Double) _
As Double
Attribute gammaLog.VB_Description = "Natural logarithm of the absolute value of the Gamma function g(x) = (x-1)! at the supplied argument. To get Gamma(x), use gammaSign(x) * Exp(gammaLog(x)), but note that this will overflow if x > 171.6"
' Natural logarithm of the absolute value of the Gamma function g(x) = (x-1)!.
'
' To get Gamma(x), use gammaSign(x) * Exp(gammaLog(x)), but note that this
' will overflow if x > 171.6
'
' Based on a Lanczos-type-series routine by Allen Miller of CSIRO. Has one less
' term in the series inside the log, and has the resulting "5.5" tweaked.
'
' The return value is consistently within 1E-14 of the true function (using
' absolute error when the function is < 1, relative error when the function is
' > 1). Much of the time, this error measure is within 1E-15 of the true value.
'
' Error limits: -2.93E18 < x < -5.57E-309 and 5.57E-309 < x < 2.55E305
'
' Version of 20 July 2007 by John Trenholme
Const Split As Double = 0.5
Dim u As Double
If x >= Split Then u = x Else u = 1# - x
Dim v As Double
v = (u - 0.5) * Log(u + 5.877) - u + Log(0.007026535811543 + _
  2.42425246210031 / u - 3.96218410578371 / (u + 1#) + _
  2.00272142074634 / (u + 2#) - 0.335162034406637 / (u + 3#) + _
  0.012977763638767 / (u + 4#) - 0.000027272752074 / (u + 5#))
If x >= Split Then
  gammaLog = v
Else
  gammaLog = Log(Abs(Pi_c / Sin(Pi_c * x))) - v - 6.5E-15
End If
End Function

'===== gammaSign ============================================================
Public Function gammaSign( _
  ByVal x As Double) _
As Double
Attribute gammaSign.VB_Description = "Sign of the Gamma function at the supplied argument. Will be negative for some negative arguments. Zero at poles 0, -1, -2, -3..."
' Returns the sign of the Gamma function, which is negative only when the
' argument x is negative, and x also lies between an odd integer and the integer
' that is one larger. Note that Gamma has poles at x = 0, -1, -2, and so on.
' This function returns 0 at those points.
'
' To get Gamma(x), use gammaSign(x) * Exp(gammaLog(x)), but note that this
' will overflow if x > 171.6
'
' Version of 20 July 2007 by John Trenholme
If x <= 0# Then
  x = 0.5 * x - 0.25
  x = Abs(x - Int(x) - 0.5) - 0.25
End If
gammaSign = Sgn(x)
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~ end of file ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

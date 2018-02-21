'
'	**************************************************************************
'	**
'	**    Class  SpecialFunction (C#)
'	**
'	**************************************************************************
'	**    Copyright (C) 1984 Stephen L. Moshier (original C version - Cephes Math Library)
'	**    Copyright (C) 1996 Leigh Brookshaw	(Java version)
'	**    Copyright (C) 2005 Miroslav Stampar	(C# version)
'	**                  2011 Xavier Flix    	(VB.NET version [->this<-])
'	**                  2017 Xavier Flix    	(Minor code refactoring)
'	**
'	**    This program is free software; you can redistribute it and/or modify
'	**    it under the terms of the GNU General Public License as published by
'	**    the Free Software Foundation; either version 2 of the License, or
'	**    (at your option) any later version.
'	**
'	**    This program is distributed in the hope that it will be useful,
'	**    but WITHOUT ANY WARRANTY; without even the implied warranty of
'	**    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'	**    GNU General Public License for more details.
'	**
'	**    You should have received a copy of the GNU General Public License
'	**    along with this program; if not, write to the Free Software
'	**    Foundation, Inc., 675 Mass Ave, Cambridge, MA 02139, USA.
'	**************************************************************************
'	**
'	**    This class is an extension of System.Math. It includes a number
'	**    of special functions not found in the Math class.
'	**
'	************************************************************************



'*
'	 * This class contains physical constants and special functions not found
'	 * in the System.Math class.
'	 * Like the System.Math class this class is final and cannot be
'	 * subclassed.
'	 * All physical constants are in cgs units.
'	 * NOTE: These special functions do not necessarily use the fastest
'	 * or most accurate algorithms.
'	 *
'	 * @version $Revision: 1.8 $, $Date: 2005/09/12 09:52:34 $
'	 

Public Class SpecialFunctions
    ' Machine constants

    Private Const MACHEP As Double = 0.000000000000000111022302462516
    Private Const MAXLOG As Double = 709.782712893384
    Private Const MINLOG As Double = -745.133219101941
    Private Const MAXGAM As Double = 171.624376956303
    Private Const SQTPI As Double = 2.506628274631
    Private Const SQRTH As Double = 0.707106781186548
    Private Const LOGPI As Double = 1.1447298858494


    ' Physical Constants in cgs Units

    ''' <summary>
    ''' Boltzman Constant. Units erg/deg(K) 
    ''' </summary>
    Public Const BOLTZMAN As Double = 0.00000000000000013807

    ''' <summary>
    ''' Elementary Charge. Units statcoulomb 
    ''' </summary>
    Public Const ECHARGE As Double = 0.00000000048032

    ''' <summary>
    ''' Electron Mass. Units g 
    ''' </summary>
    Public Const EMASS As Double = 9.1095E-28

    ''' <summary>
    ''' Proton Mass. Units g 
    ''' </summary>
    Public Const PMASS As Double = 1.6726E-24

    ''' <summary>
    ''' Gravitational Constant. Units dyne-cm^2/g^2
    ''' </summary>
    Public Const GRAV As Double = 0.00000006672

    ''' <summary>
    ''' Planck constant. Units erg-sec 
    ''' </summary>
    Public Const PLANCK As Double = 6.6262E-27

    ''' <summary>
    ''' Speed of Light in a Vacuum. Units cm/sec 
    ''' </summary>
    Public Const LIGHTSPEED As Double = 29979000000.0

    ''' <summary>
    ''' Stefan-Boltzman Constant. Units erg/cm^2-sec-deg^4 
    ''' </summary>
    Public Const STEFANBOLTZ As Double = 0.000056703

    ''' <summary>
    ''' Avogadro Number. Units  1/mol 
    ''' </summary>
    Public Const AVOGADRO As Double = 6.022E+23

    ''' <summary>
    ''' Gas Constant. Units erg/deg-mol 
    ''' </summary>
    Public Const GASCONSTANT As Double = 83144000.0

    ''' <summary>
    ''' Gravitational Acceleration at the Earth's surface. Units cm/sec^2 
    ''' </summary>
    Public Const GRAVACC As Double = 980.67

    ''' <summary>
    ''' Solar Mass. Units g 
    ''' </summary>
    Public Const SOLARMASS As Double = 1.989E+33 ' 1.99E+33

    ''' <summary>
    ''' Solar Radius. Units cm
    ''' </summary>
    Public Const SOLARRADIUS As Double = 69570000000.0 ' 69600000000.0

    ''' <summary>
    ''' Solar Luminosity. Units erg/sec
    ''' </summary>
    Public Const SOLARLUM As Double = 3.828E+33 ' 3.9E+33

    ''' <summary>
    ''' Solar Flux. Units erg/cm^2-sec
    ''' </summary>
    Public Const SOLARFLUX As Double = 64100000000.0

    ''' <summary>
    ''' Astronomical Unit (radius of the Earth's orbit). Units cm
    ''' </summary>
    Public Const AU As Double = 14959787069100.0 ' 15000000000000.0

    ''' <summary>
    ''' Don't let anyone instantiate this class. 
    ''' </summary>
    Private Sub New()
    End Sub

    ' Function Methods

    ''' <summary>
    ''' Returns the base 10 logarithm of the specified number.
    ''' </summary>
    ''' <param name="x"></param>
    ''' <returns></returns>
    Public Shared Function Log10(x As Double) As Double
        If x <= 0.0 Then Throw New ArithmeticException("range exception")
        Return Math.Log(x) / 2.30258509299405
    End Function

    Public Shared Function Log(x As Double, b As Double) As Double
        Return Math.Log(x) / Math.Log(b)
    End Function

    ''' <summary>
    ''' Returns the hyperbolic cosine of the specified number.
    ''' </summary>
    ''' <param name="x"></param>
    ''' <returns></returns>
    Public Shared Function Cosh(x As Double) As Double
        If x < 0.0 Then x = Math.Abs(x)
        x = Math.Exp(x)
        Return 0.5 * (x + 1 / x)
    End Function

    ''' <summary>
    ''' Returns the hyperbolic sine of the specified number.
    ''' </summary>
    ''' <param name="x"></param>
    ''' <returns></returns>
    Public Shared Function Sinh(x As Double) As Double
        Dim a As Double
        If x = 0.0 Then Return x
        a = x
        If a < 0.0 Then a = Math.Abs(x)
        a = Math.Exp(a)
        If x < 0.0 Then
            Return -0.5 * (a - 1 / a)
        Else
            Return 0.5 * (a - 1 / a)
        End If
    End Function

    ''' <summary>
    ''' Returns the hyperbolic tangent of the specified number.
    ''' </summary>
    ''' <param name="x"></param>
    ''' <returns></returns>
    Public Shared Function Tanh(x As Double) As Double
        Dim a As Double
        If x = 0.0 Then Return x
        a = x
        If a < 0.0 Then a = Math.Abs(x)
        a = Math.Exp(2.0 * a)
        If x < 0.0 Then
            Return -(1.0 - 2.0 / (a + 1.0))
        Else
            Return (1.0 - 2.0 / (a + 1.0))
        End If
    End Function

    ''' <summary>
    ''' Returns the hyperbolic arc cosine of the specified number.
    ''' </summary>
    ''' <param name="x"></param>
    ''' <returns></returns>
    Public Shared Function Acosh(x As Double) As Double
        If x < 1.0 Then Throw New ArithmeticException("range exception")
        Return Math.Log(x + Math.Sqrt(x * x - 1))
    End Function

    ''' <summary>
    ''' Returns the hyperbolic arc sine of the specified number.
    ''' </summary>
    ''' <param name="xx"></param>
    ''' <returns></returns>
    Public Shared Function Asinh(xx As Double) As Double
        Dim x As Double
        Dim sign As Integer
        If xx = 0.0 Then Return xx
        If xx < 0.0 Then
            sign = -1
            x = -xx
        Else
            sign = 1
            x = xx
        End If
        Return sign * Math.Log(x + Math.Sqrt(x * x + 1))
    End Function

    ''' <summary>
    ''' Returns the hyperbolic arc tangent of the specified number.
    ''' </summary>
    ''' <param name="x"></param>
    ''' <returns></returns>
    Public Shared Function Atanh(x As Double) As Double
        If x > 1.0 OrElse x < -1.0 Then Throw New ArithmeticException("range exception")
        Return 0.5 * Math.Log((1.0 + x) / (1.0 - x))
    End Function

    ''' <summary>
    ''' Returns the Bessel function of order 0 of the specified number.
    ''' </summary>
    ''' <param name="x"></param>
    ''' <returns></returns>
    Public Shared Function BezelJ0(x As Double) As Double
        Dim ax As Double

        If (InlineAssignHelper(ax, Math.Abs(x))) < 8.0 Then
            Dim y As Double = x * x
            Dim ans1 As Double = 57568490574.0 + y * (-13362590354.0 + y * (651619640.7 + y * (-11214424.18 + y * (77392.33017 + y * (-184.9052456)))))
            Dim ans2 As Double = 57568490411.0 + y * (1029532985.0 + y * (9494680.718 + y * (59272.64853 + y * (267.8532712 + y * 1.0))))

            Return ans1 / ans2
        Else
            Dim z As Double = 8.0 / ax
            Dim y As Double = z * z
            Dim xx As Double = ax - 0.785398164
            Dim ans1 As Double = 1.0 + y * (-0.001098628627 + y * (0.00002734510407 + y * (-0.000002073370639 + y * 0.0000002093887211)))
            Dim ans2 As Double = -0.01562499995 + y * (0.0001430488765 + y * (-0.000006911147651 + y * (0.0000007621095161 - y * 0.0000000934935152)))

            Return Math.Sqrt(0.636619772 / ax) * (Math.Cos(xx) * ans1 - z * Math.Sin(xx) * ans2)
        End If
    End Function

    ''' <summary>
    ''' Returns the Bessel function of order 1 of the specified number.
    ''' </summary>
    ''' <param name="x"></param>
    ''' <returns></returns>
    Public Shared Function BezelJ1(x As Double) As Double
        Dim ax As Double
        Dim y As Double
        Dim ans1 As Double, ans2 As Double

        If (InlineAssignHelper(ax, Math.Abs(x))) < 8.0 Then
            y = x * x
            ans1 = x * (72362614232.0 + y * (-7895059235.0 + y * (242396853.1 + y * (-2972611.439 + y * (15704.4826 + y * (-30.16036606))))))
            ans2 = 144725228442.0 + y * (2300535178.0 + y * (18583304.74 + y * (99447.43394 + y * (376.9991397 + y * 1.0))))
            Return ans1 / ans2
        Else
            Dim z As Double = 8.0 / ax
            Dim xx As Double = ax - 2.356194491
            y = z * z

            ans1 = 1.0 + y * (0.00183105 + y * (-0.00003516396496 + y * (0.000002457520174 + y * (-0.000000240337019))))
            ans2 = 0.04687499995 + y * (-0.0002002690873 + y * (0.000008449199096 + y * (-0.00000088228987 + y * 0.000000105787412)))
            Dim ans As Double = Math.Sqrt(0.636619772 / ax) * (Math.Cos(xx) * ans1 - z * Math.Sin(xx) * ans2)
            If x < 0.0 Then ans = -ans
            Return ans
        End If
    End Function

    ''' <summary>
    ''' Returns the Bessel function of order n of the specified number.
    ''' </summary>
    ''' <param name="n"></param>
    ''' <param name="x"></param>
    ''' <returns></returns>
    Public Shared Function BezelJn(n As Integer, x As Double) As Double
        Dim j As Integer, m As Integer
        Dim ax As Double, bj As Double, bjm As Double, bjp As Double, sum As Double, tox As Double,
         ans As Double
        Dim jsum As Boolean

        Dim ACC As Double = 40.0
        Dim BIGNO As Double = 10000000000.0
        Dim BIGNI As Double = 0.0000000001

        If n = 0 Then Return BezelJ0(x)
        If n = 1 Then Return BezelJ1(x)

        ax = Math.Abs(x)
        If ax = 0.0 Then
            Return 0.0
        ElseIf ax > CDbl(n) Then
            tox = 2.0 / ax
            bjm = BezelJ0(ax)
            bj = BezelJ1(ax)
            For j = 1 To n - 1
                bjp = j * tox * bj - bjm
                bjm = bj
                bj = bjp
            Next
            ans = bj
        Else
            tox = 2.0 / ax
            m = 2 * ((n + CInt(Math.Truncate(Math.Sqrt(ACC * n)))) \ 2)
            jsum = False
            bjp = InlineAssignHelper(ans, InlineAssignHelper(sum, 0.0))
            bj = 1.0
            For j = m To 1 Step -1
                bjm = j * tox * bj - bjp
                bjp = bj
                bj = bjm
                If Math.Abs(bj) > BIGNO Then
                    bj *= BIGNI
                    bjp *= BIGNI
                    ans *= BIGNI
                    sum *= BIGNI
                End If
                If jsum Then sum += bj
                jsum = Not jsum
                If j = n Then ans = bjp
            Next
            sum = 2.0 * sum - bj
            ans /= sum
        End If
        Return If(x < 0.0 AndAlso n Mod 2 = 1, -ans, ans)
    End Function

    ''' <summary>
    ''' Returns the Bessel function of the second kind, of order 0 of the specified number.
    ''' </summary>
    ''' <param name="x"></param>
    ''' <returns></returns>
    Public Shared Function BezelY0(x As Double) As Double
        If x < 8.0 Then
            Dim y As Double = x * x

            Dim ans1 As Double = -2957821389.0 + y * (7062834065.0 + y * (-512359803.6 + y * (10879881.29 + y * (-86327.92757 + y * 228.4622733))))
            Dim ans2 As Double = 40076544269.0 + y * (745249964.8 + y * (7189466.438 + y * (47447.2647 + y * (226.1030244 + y * 1.0))))

            Return (ans1 / ans2) + 0.636619772 * BezelJ0(x) * Math.Log(x)
        Else
            Dim z As Double = 8.0 / x
            Dim y As Double = z * z
            Dim xx As Double = x - 0.785398164

            Dim ans1 As Double = 1.0 + y * (-0.001098628627 + y * (0.00002734510407 + y * (-0.000002073370639 + y * 0.0000002093887211)))
            Dim ans2 As Double = -0.01562499995 + y * (0.0001430488765 + y * (-0.000006911147651 + y * (0.0000007621095161 + y * (-0.0000000934945152))))
            Return Math.Sqrt(0.636619772 / x) * (Math.Sin(xx) * ans1 + z * Math.Cos(xx) * ans2)
        End If
    End Function

    ''' <summary>
    ''' Returns the Bessel function of the second kind, of order 1 of the specified number.
    ''' </summary>
    ''' <param name="x"></param>
    ''' <returns></returns>
    Public Shared Function BezelY1(x As Double) As Double
        If x < 8.0 Then
            Dim y As Double = x * x
            Dim ans1 As Double = x * (-4900604943000.0 + y * (1275274390000.0 + y * (-51534381390.0 + y * (734926455.1 + y * (-4237922.726 + y * 8511.937935)))))
            Dim ans2 As Double = 24995805700000.0 + y * (424441966400.0 + y * (3733650367.0 + y * (22459040.02 + y * (102042.605 + y * (354.9632885 + y)))))
            Return (ans1 / ans2) + 0.636619772 * (BezelJ1(x) * Math.Log(x) - 1.0 / x)
        Else
            Dim z As Double = 8.0 / x
            Dim y As Double = z * z
            Dim xx As Double = x - 2.356194491
            Dim ans1 As Double = 1.0 + y * (0.00183105 + y * (-0.00003516396496 + y * (0.000002457520174 + y * (-0.000000240337019))))
            Dim ans2 As Double = 0.04687499995 + y * (-0.0002002690873 + y * (0.000008449199096 + y * (-0.00000088228987 + y * 0.000000105787412)))
            Return Math.Sqrt(0.636619772 / x) * (Math.Sin(xx) * ans1 + z * Math.Cos(xx) * ans2)
        End If
    End Function

    ''' <summary>
    ''' Returns the Bessel function of the second kind, of order n of the specified number.
    ''' </summary>
    ''' <param name="n"></param>
    ''' <param name="x"></param>
    ''' <returns></returns>
    Public Shared Function BezelYn(n As Integer, x As Double) As Double
        Dim by As Double, bym As Double, byp As Double, tox As Double

        If n = 0 Then Return BezelY0(x)
        If n = 1 Then Return BezelY1(x)

        tox = 2.0 / x
        by = BezelY1(x)
        bym = BezelY0(x)
        For j As Integer = 1 To n - 1
            byp = j * tox * by - bym
            bym = by
            by = byp
        Next
        Return by
    End Function

    ''' <summary>
    ''' Returns the factorial of the specified number.
    ''' </summary>
    ''' <param name="x"></param>
    ''' <returns></returns>
    Public Shared Function Fac(x As Double) As Double
        Dim d As Double = Math.Abs(x)
        If Math.Floor(d) = d Then
            Return CDbl(Fac(CInt(Math.Truncate(x))))
        Else
            Return Gamma(x + 1.0)
        End If
    End Function

    ''' <summary>
    ''' Returns the factorial of the specified number.
    ''' </summary>
    ''' <param name="j"></param>
    ''' <returns></returns>
    Public Shared Function Fac(j As Long) As Double
        Dim i As Long = j
        Dim d As Double = 1.0
        If j < 0 Then i = Math.Abs(j)
        While i > 1
            d *= Math.Max(Threading.Interlocked.Decrement(i), i + 1)
        End While
        If j < 0 Then
            Return -d
        Else
            Return d
        End If
    End Function

    ''' <summary>
    ''' Returns the gamma function of the specified number.
    ''' </summary>
    ''' <param name="x"></param>
    ''' <returns></returns>
    Public Shared Function Gamma(x As Double) As Double
        Static p1 As Double() = {0.000160119522476752,
                                0.00119135147006586,
                                0.0104213797561762,
                                0.0476367800457137,
                                0.207448227648436,
                                0.494214826801497,
                                1.0}
        Static q2 As Double() = {-0.000023158187332412,
                                0.000539605580493303,
                                -0.00445641913851797,
                                0.011813978522206,
                                0.0358236398605499,
                                -0.234591795718243,
                                0.0714304917030273,
                                1.0}

        Dim p3 As Double
        Dim z As Double
        Dim q4 As Double = Math.Abs(x)

        If q4 > 33.0 Then
            If x < 0.0 Then
                p3 = Math.Floor(q4)
                If p3 = q4 Then Throw New OverflowException()

                'int i = (int)p;
                z = q4 - p3
                If z > 0.5 Then
                    p3 += 1.0
                    z = q4 - p3
                End If
                z = q4 * Math.Sin(Math.PI * z)
                If z = 0.0 Then Throw New OverflowException()

                z = Math.Abs(z)
                z = Math.PI / (z * Stirf(q4))

                Return -z
            Else
                Return Stirf(x)
            End If
        End If

        z = 1.0
        While x >= 3.0
            x -= 1.0
            z *= x
        End While

        While x < 0.0
            If x = 0.0 Then
                Throw New ArithmeticException($"{NameOf(Gamma)}: Singular")
            ElseIf x > -0.000000001 Then
                Return (z / ((1.0 + 0.577215664901533 * x) * x))
            End If
            z /= x
            x += 1.0
        End While

        While x < 2.0
            If x = 0.0 Then
                Throw New ArithmeticException($"{NameOf(Gamma)}: Singular")
            ElseIf x < 0.000000001 Then
                Return (z / ((1.0 + 0.577215664901533 * x) * x))
            End If
            z /= x
            x += 1.0
        End While

        If (x = 2.0) OrElse (x = 3.0) Then Return z

        x -= 2.0
        p3 = Polevl(x, p1, 6)
        q4 = Polevl(x, q2, 7)
        Return z * p3 / q4
    End Function

    ''' <summary>
    ''' Return the gamma function computed by Stirling's formula.
    ''' </summary>
    ''' <param name="x"></param>
    ''' <returns></returns>
    Private Shared Function Stirf(x As Double) As Double
        Static STIR As Double() = {0.000787311395793094,
                                -0.000229549961613378,
                                -0.00268132617805781,
                                0.00347222221605459,
                                0.0833333333333482}
        Dim MAXSTIR As Double = 143.01608

        Dim w As Double = 1.0 / x
        Dim y As Double = Math.Exp(x)
        Dim v As Double

        w = 1.0 + w * Polevl(w, STIR, 4)

        If x > MAXSTIR Then
            ' Avoid overflow in Math.Pow() 

            v = Math.Pow(x, 0.5 * x - 0.25)
            y = v * (v / y)
        Else
            y = Math.Pow(x, x - 0.5) / y
        End If
        Return SQTPI * y * w
    End Function

    ''' <summary>
    ''' Returns the complemented incomplete gamma function.
    ''' </summary>
    ''' <param name="a"></param>
    ''' <param name="x"></param>
    ''' <returns></returns>
    Public Shared Function Igamc(a As Double, x As Double) As Double
        Dim big As Double = 4.5035996273705E+15
        Dim biginv As Double = 0.000000000000000222044604925031
        Dim ans As Double, ax As Double, c As Double, yc As Double, r As Double, t As Double, y As Double, z As Double
        Dim pk As Double, pkm1 As Double, pkm2 As Double, qk As Double, qkm1 As Double, qkm2 As Double

        If x <= 0 OrElse a <= 0 Then Return 1.0
        If x < 1.0 OrElse x < a Then Return 1.0 - Igam(a, x)

        ax = a * Math.Log(x) - x - Lgamma(a)
        If ax < -MAXLOG Then Return 0.0

        ax = Math.Exp(ax)

        ' continued fraction 

        y = 1.0 - a
        z = x + y + 1.0
        c = 0.0
        pkm2 = 1.0
        qkm2 = x
        pkm1 = x + 1.0
        qkm1 = z * x
        ans = pkm1 / qkm1

        Do
            c += 1.0
            y += 1.0
            z += 2.0
            yc = y * c
            pk = pkm1 * z - pkm2 * yc
            qk = qkm1 * z - qkm2 * yc
            If qk <> 0 Then
                r = pk / qk
                t = Math.Abs((ans - r) / r)
                ans = r
            Else
                t = 1.0
            End If

            pkm2 = pkm1
            pkm1 = pk
            qkm2 = qkm1
            qkm1 = qk
            If Math.Abs(pk) > big Then
                pkm2 *= biginv
                pkm1 *= biginv
                qkm2 *= biginv
                qkm1 *= biginv
            End If
        Loop While t > MACHEP

        Return ans * ax
    End Function

    ''' <summary>
    ''' Returns the incomplete gamma function.
    ''' </summary>
    ''' <param name="a"></param>
    ''' <param name="x"></param>
    ''' <returns></returns>
    Public Shared Function Igam(a As Double, x As Double) As Double
        Dim ans As Double, ax As Double, c As Double, r As Double

        If x <= 0 OrElse a <= 0 Then Return 0.0
        If x > 1.0 AndAlso x > a Then Return 1.0 - Igamc(a, x)

        ' Compute  x**a * exp(-x) / gamma(a)  

        ax = a * Math.Log(x) - x - Lgamma(a)
        If ax < -MAXLOG Then Return (0.0)
        ax = Math.Exp(ax)

        ' power series 

        r = a
        c = 1.0
        ans = 1.0

        Do
            r += 1.0
            c *= x / r
            ans += c
        Loop While c / ans > MACHEP

        Return ans * ax / a
    End Function

    ''' <summary>
    ''' Returns the area under the left hand tail (from 0 to x)
    ''' of the Chi square probability density function with
    ''' v degrees of freedom.
    ''' </summary>
    ''' <param name="df">degrees of freedom</param>
    ''' <param name="x">double value</param>
    ''' <returns></returns>
    Public Shared Function Chisq(df As Double, x As Double) As Double
        If x < 0.0 OrElse df < 1.0 Then Return 0.0
        Return Igam(df / 2.0, x / 2.0)
    End Function

    ''' <summary>
    ''' Returns the area under the right hand tail (from x to
    ''' infinity) of the Chi square probability density function
    ''' with v degrees of freedom.
    ''' </summary>
    ''' <param name="df">degrees of freedom</param>
    ''' <param name="x">double value</param>
    ''' <returns></returns>
    Public Shared Function Chisqc(df As Double, x As Double) As Double
        If x < 0.0 OrElse df < 1.0 Then Return 0.0
        Return Igamc(df / 2.0, x / 2.0)
    End Function


    ''' <summary>
    ''' Returns the sum of the first k terms of the Poisson distribution.
    ''' </summary>
    ''' <param name="k">number of terms</param>
    ''' <param name="x">double value</param>
    ''' <returns></returns>
    Public Shared Function Poisson(k As Integer, x As Double) As Double
        If k < 0 OrElse x < 0 Then Return 0.0
        Return Igamc(CDbl(k + 1), x)
    End Function


    ''' <summary>
    ''' Returns the sum of the terms k+1 to infinity of the Poisson distribution.
    ''' </summary>
    ''' <param name="k">start</param>
    ''' <param name="x">double value</param>
    ''' <returns></returns>
    Public Shared Function Poissonc(k As Integer, x As Double) As Double
        If k < 0 OrElse x < 0 Then Return 0.0
        Return Igam(CDbl(k + 1), x)
    End Function

    ''' <summary>
    ''' Returns the area under the Gaussian probability density function, integrated from minus infinity to a.
    ''' </summary>
    ''' <param name="a"></param>
    ''' <returns></returns>
    Public Shared Function Normal(a As Double) As Double
        Dim x As Double, y As Double, z As Double

        x = a * SQRTH
        z = Math.Abs(x)

        If z < SQRTH Then
            y = 0.5 + 0.5 * Erf(x)
        Else
            y = 0.5 * Erfc(z)
            If x > 0 Then y = 1.0 - y
        End If

        Return y
    End Function

    ''' <summary>
    ''' Returns the complementary error function of the specified number.
    ''' </summary>
    ''' <param name="a"></param>
    ''' <returns></returns>
    Public Shared Function Erfc(a As Double) As Double
        Dim x As Double, y As Double, z As Double, p1 As Double, q2 As Double

        Dim p3 As Double() = {0.000000000246196981473531, 0.564189564831069, 7.4632105644227, 48.6371970985681, 196.520832956077, 526.445194995477,
         934.528527171958, 1027.55188689516, 557.535335369399}
        '1.0
        Dim q4 As Double() = {13.2281951154745, 86.707214088599, 354.93777888782, 975.708501743205, 1823.9091668791, 2246.33760818711,
         1656.66309194161, 557.535340817728}

        Dim r As Double() = {0.564189583547755, 1.27536670759978, 5.01905042251181, 6.16021097993054, 7.40974269950449, 2.978866653721}
        '1.00000000000000000000E0, 
        Dim s As Double() = {2.26052863220117, 9.39603524938001, 12.0489539808097, 17.0814450747566, 9.60896809063286, 3.36907645100082}

        If a < 0.0 Then
            x = -a
        Else
            x = a
        End If

        If x < 1.0 Then Return 1.0 - Erf(a)

        z = -a * a

        If z < -MAXLOG Then
            If a < 0 Then
                Return (2.0)
            Else
                Return (0.0)
            End If
        End If

        z = Math.Exp(z)

        If x < 8.0 Then
            p1 = Polevl(x, p3, 8)
            q2 = P1evl(x, q4, 8)
        Else
            p1 = Polevl(x, r, 5)
            q2 = P1evl(x, s, 6)
        End If

        y = (z * p1) / q2

        If a < 0 Then
            y = 2.0 - y
        End If

        If y = 0.0 Then
            If a < 0 Then
                Return 2.0
            Else
                Return (0.0)
            End If
        End If

        Return y
    End Function

    ''' <summary>
    ''' Returns the error function of the specified number.
    ''' </summary>
    ''' <param name="x"></param>
    ''' <returns></returns>
    Public Shared Function Erf(x As Double) As Double
        Dim y As Double, z As Double
        Dim t As Double() = {9.60497373987052, 90.0260197203843, 2232.00534594684, 7003.32514112805, 55592.3013010395}
        '1.00000000000000000000E0,
        Dim u As Double() = {33.5617141647503, 521.357949780153, 4594.3238297098, 22629.0000613891, 49267.3942608636}

        If Math.Abs(x) > 1.0 Then Return (1.0 - Erfc(x))

        z = x * x
        y = x * Polevl(z, t, 4) / P1evl(z, u, 5)
        Return y
    End Function

    ''' <summary>
    ''' Evaluates polynomial of degree N
    ''' </summary>
    ''' <param name="x"></param>
    ''' <param name="coef"></param>
    ''' <param name="N"></param>
    ''' <returns></returns>
    Private Shared Function Polevl(x As Double, coef As Double(), N As Integer) As Double
        Dim ans As Double = coef(0)

        For i As Integer = 1 To N
            ans = ans * x + coef(i)
        Next

        Return ans
    End Function

    ''' <summary>
    ''' Evaluates polynomial of degree N with assumption that coef[N] = 1.0
    ''' </summary>
    ''' <param name="x"></param>
    ''' <param name="coef"></param>
    ''' <param name="N"></param>
    ''' <returns></returns>		
    Private Shared Function P1evl(x As Double, coef As Double(), N As Integer) As Double
        Dim ans As Double

        ans = x + coef(0)

        For i As Integer = 1 To N - 1
            ans = ans * x + coef(i)
        Next

        Return ans
    End Function

    ''' <summary>
    ''' Returns the natural logarithm of gamma function.
    ''' </summary>
    ''' <param name="x"></param>
    ''' <returns></returns>
    Public Shared Function Lgamma(x As Double) As Double
        Dim p As Double, q As Double, w As Double, z As Double

        Dim a As Double() = {0.000811614167470508, -0.000595061904284301, 0.000793650340457717, -0.002777777777301, 0.0833333333333332}
        Dim b As Double() = {-1378.25152569121, -38801.6315134638, -331612.992738871, -1162370.97492762, -1721737.0082084, -853555.664245765}
        ' 1.00000000000000000000E0, 

        Dim c As Double() = {-351.815701436523, -17064.2106651881, -220528.590553854, -1139334.44367983, -2532523.07177583, -2018891.41433533}

        If x < -34.0 Then
            q = -x
            w = Lgamma(q)
            p = Math.Floor(q)
            If p = q Then Throw New OverflowException()

            z = q - p
            If z > 0.5 Then
                p += 1.0
                z = p - q
            End If
            z = q * Math.Sin(Math.PI * z)
            If z = 0.0 Then Throw New OverflowException()
            z = LOGPI - Math.Log(z) - w
            Return z
        End If

        If x < 13.0 Then
            z = 1.0
            While x >= 3.0
                x -= 1.0
                z *= x
            End While
            While x < 2.0
                If x = 0.0 Then Throw New OverflowException()
                z /= x
                x += 1.0
            End While
            If z < 0.0 Then z = -z
            If x = 2.0 Then Return Math.Log(z)
            x -= 2.0
            p = x * Polevl(x, b, 5) / P1evl(x, c, 6)
            Return (Math.Log(z) + p)
        End If

        If x > 2.556348E+305 Then Throw New OverflowException()

        q = (x - 0.5) * Math.Log(x) - x + 0.918938533204673
        If x > 100000000.0 Then Return (q)

        p = 1.0 / (x * x)
        If x >= 1000.0 Then
            q += ((0.000793650793650794 * p - 0.00277777777777778) * p + 0.0833333333333333) / x
        Else
            q += Polevl(p, a, 4) / x
        End If
        Return q
    End Function


    ''' <summary>
    ''' Returns the incomplete beta function evaluated from zero to xx.
    ''' </summary>
    ''' <param name="aa"></param>
    ''' <param name="bb"></param>
    ''' <param name="xx"></param>
    ''' <returns></returns>
    Public Shared Function Ibeta(aa As Double, bb As Double, xx As Double) As Double
        Dim a As Double, b As Double, t As Double, x As Double, xc As Double, w As Double,
         y As Double
        Dim flag As Boolean

        If aa <= 0.0 OrElse bb <= 0.0 Then Throw New ArithmeticException($"{NameOf(Ibeta)}: Domain error")

        If (xx <= 0.0) OrElse (xx >= 1.0) Then
            If xx = 0.0 Then Return 0.0
            If xx = 1.0 Then Return 1.0
            Throw New ArithmeticException($"{NameOf(Ibeta)}: Domain error")
        End If

        flag = False
        If (bb * xx) <= 1.0 AndAlso xx <= 0.95 Then
            t = Pseries(aa, bb, xx)
            Return t
        End If

        w = 1.0 - xx

        ' Reverse a and b if x is greater than the mean. 

        If xx > (aa / (aa + bb)) Then
            flag = True
            a = bb
            b = aa
            xc = xx
            x = w
        Else
            a = aa
            b = bb
            xc = w
            x = xx
        End If

        If flag AndAlso (b * x) <= 1.0 AndAlso x <= 0.95 Then
            t = Pseries(a, b, x)
            If t <= MACHEP Then
                t = 1.0 - MACHEP
            Else
                t = 1.0 - t
            End If
            Return t
        End If

        ' Choose expansion for better convergence. 

        y = x * (a + b - 2.0) - (a - 1.0)
        If y < 0.0 Then
            w = Incbcf(a, b, x)
        Else
            w = Incbd(a, b, x) / xc
        End If

        ' Multiply w by the factor
        '			   a      b   _             _     _
        '			  x  (1-x)   | (a+b) / ( a | (a) | (b) ) .   

        y = a * Math.Log(x)
        t = b * Math.Log(xc)
        If (a + b) < MAXGAM AndAlso Math.Abs(y) < MAXLOG AndAlso Math.Abs(t) < MAXLOG Then
            t = Math.Pow(xc, b)
            t *= Math.Pow(x, a)
            t /= a
            t *= w
            t *= Gamma(a + b) / (Gamma(a) * Gamma(b))
            If flag Then
                If t <= MACHEP Then
                    t = 1.0 - MACHEP
                Else
                    t = 1.0 - t
                End If
            End If
            Return t
        End If

        ' Resort to logarithms.  

        y += t + Lgamma(a + b) - Lgamma(a) - Lgamma(b)
        y += Math.Log(w / a)
        If y < MINLOG Then
            t = 0.0
        Else
            t = Math.Exp(y)
        End If

        If flag Then
            If t <= MACHEP Then
                t = 1.0 - MACHEP
            Else
                t = 1.0 - t
            End If
        End If
        Return t
    End Function

    ''' <summary>
    ''' Returns the continued fraction expansion #1 for incomplete beta integral.
    ''' </summary>
    ''' <param name="a"></param>
    ''' <param name="b"></param>
    ''' <param name="x"></param>
    ''' <returns></returns>
    Private Shared Function Incbcf(a As Double, b As Double, x As Double) As Double
        Dim xk As Double, pk As Double, pkm1 As Double, pkm2 As Double, qk As Double, qkm1 As Double, qkm2 As Double
        Dim k1 As Double, k2 As Double, k3 As Double, k4 As Double, k5 As Double, k6 As Double, k7 As Double, k8 As Double
        Dim r As Double, t As Double, ans As Double, thresh As Double
        Dim n As Integer
        Dim big As Double = 4.5035996273705E+15
        Dim biginv As Double = 0.000000000000000222044604925031

        k1 = a
        k2 = a + b
        k3 = a
        k4 = a + 1.0
        k5 = 1.0
        k6 = b - 1.0
        k7 = k4
        k8 = a + 2.0

        pkm2 = 0.0
        qkm2 = 1.0
        pkm1 = 1.0
        qkm1 = 1.0
        ans = 1.0
        r = 1.0
        n = 0
        thresh = 3.0 * MACHEP
        Do
            xk = -(x * k1 * k2) / (k3 * k4)
            pk = pkm1 + pkm2 * xk
            qk = qkm1 + qkm2 * xk
            pkm2 = pkm1
            pkm1 = pk
            qkm2 = qkm1
            qkm1 = qk

            xk = (x * k5 * k6) / (k7 * k8)
            pk = pkm1 + pkm2 * xk
            qk = qkm1 + qkm2 * xk
            pkm2 = pkm1
            pkm1 = pk
            qkm2 = qkm1
            qkm1 = qk

            If qk <> 0 Then r = pk / qk
            If r <> 0 Then
                t = Math.Abs((ans - r) / r)
                ans = r
            Else
                t = 1.0
            End If

            If t < thresh Then Return ans

            k1 += 1.0
            k2 += 1.0
            k3 += 2.0
            k4 += 2.0
            k5 += 1.0
            k6 -= 1.0
            k7 += 2.0
            k8 += 2.0

            If (Math.Abs(qk) + Math.Abs(pk)) > big Then
                pkm2 *= biginv
                pkm1 *= biginv
                qkm2 *= biginv
                qkm1 *= biginv
            End If
            If (Math.Abs(qk) < biginv) OrElse (Math.Abs(pk) < biginv) Then
                pkm2 *= big
                pkm1 *= big
                qkm2 *= big
                qkm1 *= big
            End If
        Loop While Threading.Interlocked.Increment(n) < 300

        Return ans
    End Function

    ''' <summary>
    ''' Returns the continued fraction expansion #2 for incomplete beta integral.
    ''' </summary>
    ''' <param name="a"></param>
    ''' <param name="b"></param>
    ''' <param name="x"></param>
    ''' <returns></returns>
    Private Shared Function Incbd(a As Double, b As Double, x As Double) As Double
        Dim xk As Double, pk As Double, pkm1 As Double, pkm2 As Double, qk As Double, qkm1 As Double, qkm2 As Double
        Dim k1 As Double, k2 As Double, k3 As Double, k4 As Double, k5 As Double, k6 As Double, k7 As Double, k8 As Double
        Dim r As Double, t As Double, ans As Double, z As Double, thresh As Double
        Dim n As Integer
        Dim big As Double = 4.5035996273705E+15
        Dim biginv As Double = 0.000000000000000222044604925031

        k1 = a
        k2 = b - 1.0
        k3 = a
        k4 = a + 1.0
        k5 = 1.0
        k6 = a + b
        k7 = a + 1.0

        k8 = a + 2.0

        pkm2 = 0.0
        qkm2 = 1.0
        pkm1 = 1.0
        qkm1 = 1.0
        z = x / (1.0 - x)
        ans = 1.0
        r = 1.0
        n = 0
        thresh = 3.0 * MACHEP
        Do
            xk = -(z * k1 * k2) / (k3 * k4)
            pk = pkm1 + pkm2 * xk
            qk = qkm1 + qkm2 * xk
            pkm2 = pkm1
            pkm1 = pk
            qkm2 = qkm1
            qkm1 = qk

            xk = (z * k5 * k6) / (k7 * k8)
            pk = pkm1 + pkm2 * xk
            qk = qkm1 + qkm2 * xk
            pkm2 = pkm1
            pkm1 = pk
            qkm2 = qkm1
            qkm1 = qk

            If qk <> 0 Then r = pk / qk
            If r <> 0 Then
                t = Math.Abs((ans - r) / r)
                ans = r
            Else
                t = 1.0
            End If

            If t < thresh Then Return ans

            k1 += 1.0
            k2 -= 1.0
            k3 += 2.0
            k4 += 2.0
            k5 += 1.0
            k6 += 1.0
            k7 += 2.0
            k8 += 2.0

            If (Math.Abs(qk) + Math.Abs(pk)) > big Then
                pkm2 *= biginv
                pkm1 *= biginv
                qkm2 *= biginv
                qkm1 *= biginv
            End If
            If (Math.Abs(qk) < biginv) OrElse (Math.Abs(pk) < biginv) Then
                pkm2 *= big
                pkm1 *= big
                qkm2 *= big
                qkm1 *= big
            End If
        Loop While Threading.Interlocked.Increment(n) < 300

        Return ans
    End Function

    ''' <summary>
    ''' Returns the power series for incomplete beta integral. Use when b*x is small and x not too close to 1.
    ''' </summary>
    ''' <param name="a"></param>
    ''' <param name="b"></param>
    ''' <param name="x"></param>
    ''' <returns></returns>
    Private Shared Function Pseries(a As Double, b As Double, x As Double) As Double
        Dim s As Double, t As Double, u As Double, v As Double, n As Double, t1 As Double, z As Double, ai As Double

        ai = 1.0 / a
        u = (1.0 - b) * x
        v = u / (a + 1.0)
        t1 = v
        t = u
        n = 2.0
        s = 0.0
        z = MACHEP * ai
        While Math.Abs(v) > z
            u = (n - b) * x / n
            t *= u
            v = t / (a + n)
            s += v
            n += 1.0
        End While
        s += t1
        s += ai

        u = a * Math.Log(x)
        If (a + b) < MAXGAM AndAlso Math.Abs(u) < MAXLOG Then
            t = Gamma(a + b) / (Gamma(a) * Gamma(b))
            s = s * t * Math.Pow(x, a)
        Else
            t = Lgamma(a + b) - Lgamma(a) - Lgamma(b) + u + Math.Log(s)
            If t < MINLOG Then
                s = 0.0
            Else
                s = Math.Exp(t)
            End If
        End If
        Return s
    End Function

    Private Shared Function InlineAssignHelper(Of T)(ByRef target As T, value As T) As T
        target = value
        Return value
    End Function
End Class

Imports System.Threading

Module ModuleMain
    '11.0007

    Private cancelThreads As Boolean
    Private Delegate Function FactorialAlgorithm(value As Double) As Double
    Private lckObj As New Object()
    Private suspendThread As Boolean

    Sub Main()
        Console.Clear()

        Console.BackgroundColor = ConsoleColor.Blue
        Console.ForegroundColor = ConsoleColor.White
        Console.WriteLine("Factorial Calculator" + Space(Console.WindowWidth - 20))

        Dim crsr As New Thread(AddressOf FactorialCursor)
        crsr.Start()

        Dim value As Double

        Do
            ResetColors()

            Dim inputText = Console.ReadLine()
            If inputText = "" Then Exit Do

            If Double.TryParse(inputText, value) Then
                Console.CursorTop -= 1
                Console.WriteLine(value.ToString() + "! ")
                Console.Write(Space(Console.WindowWidth - 1))
                Console.WriteLine()
                Console.CursorLeft = 0

                Console.CursorVisible = False
                ClearResultsArea()
                Console.CursorVisible = True

                Dim eleapsedTime As New List(Of TimeSpan)
                Dim iterations As Integer = 10

                suspendThread = True

                eleapsedTime.Add(RunAlgorithm("Recursive", value, AddressOf FactRecursive, iterations))
                eleapsedTime.Add(RunAlgorithm("While", value, AddressOf FactWhile, iterations))
                eleapsedTime.Add(RunAlgorithm("For", value, AddressOf FactFor, iterations))
                eleapsedTime.Add(RunAlgorithm("Split", value, AddressOf FactSplit, iterations))
                eleapsedTime.Add(RunAlgorithm("Gamma", value, AddressOf SpecialFunctions.Fac, iterations))

                suspendThread = False

                SyncLock lckObj
                    Console.CursorTop -= (eleapsedTime.Count * 3.5)
                    For Each et As TimeSpan In eleapsedTime
                        Dim ets = et.ToString()
                        Console.CursorLeft = 2

                        Console.BackgroundColor = ConsoleColor.Black
                        Console.Write(ets)
                        Console.BackgroundColor = ConsoleColor.DarkMagenta

                        For x = 2 To et.Ticks / eleapsedTime.Max.Ticks * (Console.WindowWidth - 2)
                            Console.CursorLeft = x
                            Console.Write(If(x - 2 < ets.Length, ets(x - 2), " "))
                        Next
                        Console.CursorTop += 4
                    Next

                    ResetColors()

                    Console.CursorTop = 2
                    Console.CursorLeft = 0
                    Console.Write(Space(Console.WindowWidth - 1))
                    Console.CursorLeft = 0
                End SyncLock
            Else
                Console.ForegroundColor = ConsoleColor.Red
                Console.WriteLine("Invalid Number")

                Console.CursorTop = 2
                Console.CursorLeft = 0
                Console.Write(Space(Console.WindowWidth - 1))
                Console.CursorLeft = 0
            End If
        Loop

        'Console.ReadKey()
        cancelThreads = True
    End Sub

    Private Sub ResetColors()
        Console.BackgroundColor = ConsoleColor.Black
        Console.ForegroundColor = ConsoleColor.White
    End Sub

    Private Sub ClearResultsArea()
        SyncLock lckObj
            Dim ox As Integer = Console.CursorLeft
            Dim oy As Integer = Console.CursorTop

            For x = ox To Console.WindowWidth - 2
                For y = oy To Console.WindowHeight - 2
                    Console.SetCursorPosition(x, y)
                    Console.Write(" ")
                Next
            Next

            Console.CursorLeft = ox
            Console.CursorTop = oy
        End SyncLock
    End Sub

    Private Function RunAlgorithm(title As String, value As Double, algorithm As FactorialAlgorithm, iterations As Integer) As TimeSpan
        Dim sw As Stopwatch = New Stopwatch()
        Dim result As Double
        Dim eleapsedTime As TimeSpan

        For i As Integer = 1 To iterations
            SyncLock lckObj
                sw.Start()
                Try
                    result = algorithm(value)
                Catch ex As StackOverflowException
                    result = Double.NaN
                Catch ex As OverflowException
                    result = Double.NaN
                End Try
                sw.Stop()
            End SyncLock
            eleapsedTime += sw.Elapsed
        Next
        eleapsedTime = TimeSpan.FromTicks(eleapsedTime.Ticks / iterations)

        Console.ForegroundColor = ConsoleColor.Cyan
        Console.WriteLine(title + ":")
        Console.ForegroundColor = ConsoleColor.White
        Try
            Console.WriteLine(String.Format("  {0}! = {1}", value, FormatRAsN(result)))
        Catch ex As Exception
            Console.WriteLine(String.Format("  {0}! = {1}", value, ex.Message))
        End Try
        Console.WriteLine()
        Console.WriteLine()

        Return eleapsedTime
    End Function

    Private Sub FactorialCursor()
        Dim blank As String
        Do
            If Not suspendThread Then
                SyncLock lckObj
                    blank = Space(Console.WindowWidth - 2 - Console.CursorLeft)
                    Console.Write("!" + blank)
                    Console.CursorLeft -= blank.Length + 1
                End SyncLock
            End If

            Thread.Sleep(250)
        Loop Until cancelThreads
    End Sub

    ' http://stackoverflow.com/questions/611552/c-sharp-converting-20-digit-precision-double-to-string-and-back-again
    Private Function FormatRAsN(value As Double) As String
        Dim r = ""
        Dim s = value.ToString("R")

        If Char.IsDigit(s(0)) Then
            If s.Contains("."c) Then
                Dim tokens() = s.Split("."c)
                r = "." + tokens(1)
                s = tokens(0)
            End If

            Dim k As Integer = 0
            For i As Integer = s.Length - 1 To 0 Step -1
                k += 1
                If k = 4 Then
                    k = 1
                    r = "," + r
                End If
                r = s(i) + r
            Next
        Else
            r = s
        End If

        Return r
    End Function

    ' Non-Gamma Algorithms

    Public Function FactRecursive(value As Double) As Double
        If value < 2 Then Return 1
        Return value * FactRecursive(value - 1)
    End Function

    Public Function FactWhile(value As Double) As Double
        Dim result As Double = 1
        While value >= 2
            result *= value
            value -= 1
        End While

        Return result
    End Function

    Public Function FactFor(value As Double) As Double
        Dim result As Double = 1

        For value = value To 2 Step -1
            result *= value
        Next

        Return result
    End Function

    ' http://www.luschny.de/math/factorial/FastFactorialFunctions.htm
    ' http://www.luschny.de/math/factorial/csharp/FactorialSplit.cs.html
    Public Function FactSplit(value As Double) As Double
        value = Math.Truncate(value)

        Dim curVal As Double = 1
        Dim h As Integer = 0
        Dim shift As Integer = 0
        Dim high As Integer = 1
        Dim log2n As Integer = Math.Floor(SpecialFunctions.Log(value, 2))
        Dim len As Integer

        Dim p As Double = 1
        Dim r As Double = 1

        Dim Product As Func(Of Double, Double) = Function(n As Double) As Double
                                                     Dim m As Integer = n \ 2
                                                     If m = 0 Then
                                                         curVal += 2
                                                         Return curVal
                                                     End If
                                                     If n = 2 Then
                                                         ' Wish VB.NET supported this:
                                                         '    return (currentN += 2) * (currentN += 2)
                                                         Dim tmp As Double = (curVal + 2) * (curVal + 4)
                                                         curVal += 4
                                                       Return tmp
                                                   End If
                                                   Return Product(n - m) * Product(m)
                                               End Function

        While h <> value
            shift += h
            h = value >> log2n
            log2n -= 1
            len = high
            high = (h - 1) Or 1
            len = (high - len) \ 2

            If len > 0 Then
                p *= Product(len)
                r *= p
            End If
        End While

        Return r << shift
    End Function
End Module

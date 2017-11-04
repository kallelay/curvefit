Module externals

    Public XYZ_MODE As Boolean = False


    'Script generator
    Function fromListToPython(list As List(Of Double)) As String
        fromListToPython = "["
        For Each item In list
            fromListToPython &= Replace(CStr(item), ",", ".") & ", "
        Next
        fromListToPython = fromListToPython.Substring(0, fromListToPython.Length - 2)
        fromListToPython &= "]"
    End Function

    Function generateScript(xdata$, ydata$, functionstr$, Optional Plot As Boolean = False, Optional weights$ = "")
        generateScript = _
        "import numpy as np" & vbNewLine & _
        "from scipy.optimize import curve_fit" & vbNewLine & _
        "xdata = np.array(" & xdata & ");" & vbNewLine & _
        "ydata = np.array(" & ydata & ");" & vbNewLine & _
        functionstr$ & vbNewLine & _
        If(weights <> "", weights, "") & vbNewLine & _
        "popt, pcov = curve_fit(f = func, xdata = xdata, ydata = ydata)" & vbNewLine & _
        "print(popt)" '& vbNewLine & _
        ' '"print(pcov)" & vbNewLine & _
        '"print(""@"")" & vbNewLine '& _

        ' Clipboard.SetText(generateScript)
        '
        '#print func(xdata, *popt)

        If Plot Then generateScript &= vbNewLine & "import matplotlib.pyplot as plt" & vbNewLine & _
"res = func(xdata,*popt)" & vbNewLine & _
"print(""@""+str(res)+""@"");" & vbNewLine & _
"plt.plot(xdata, ydata, 'b-', label='data')" & vbNewLine & _
"plt.plot(xdata, func(xdata, *popt), 'r--', label='fit')" & vbNewLine & _
"plt.xlabel('x')" & vbNewLine & _
"plt.ylabel('y')" & vbNewLine & _
        "plt.legend()" & vbNewLine & _
        "plt.show()" & vbNewLine


    End Function


    'Python launchers

    Sub startPython(pypath$)
        P = New Process()
        log = Nothing
        Dim pinfo As New ProcessStartInfo
        Try
            P.StartInfo.FileName = "py"
            P.StartInfo.Arguments = "-3 """ & pypath & """"
            P.StartInfo.RedirectStandardError = True
            P.StartInfo.RedirectStandardOutput = True
            P.StartInfo.UseShellExecute = False
            P.StartInfo.CreateNoWindow = True
            P.EnableRaisingEvents = True
            Application.DoEvents()

            AddHandler P.ErrorDataReceived, AddressOf errDataRec
            AddHandler P.OutputDataReceived, AddressOf outpDataRec
            AddHandler P.Exited, AddressOf cleanUp
            P.Start()
            P.PriorityClass = ProcessPriorityClass.High

            P.BeginErrorReadLine()
            P.BeginOutputReadLine()


        Catch ex As Exception

            Try
                P.StartInfo.FileName = "python3"
                P.StartInfo.Arguments = pypath
                P.StartInfo.RedirectStandardError = True
                P.StartInfo.RedirectStandardOutput = True
                P.StartInfo.UseShellExecute = False
                P.StartInfo.CreateNoWindow = True
                P.EnableRaisingEvents = True
                Application.DoEvents()

                AddHandler P.ErrorDataReceived, AddressOf errDataRec
                AddHandler P.OutputDataReceived, AddressOf outpDataRec
                AddHandler P.Exited, AddressOf cleanup
                P.Start()
                P.PriorityClass = ProcessPriorityClass.High

                P.BeginErrorReadLine()
                P.BeginOutputReadLine()
            Catch ex2 As Exception
               
                Try
                    P.StartInfo.FileName = "python"
                    P.StartInfo.Arguments = pypath
                    P.StartInfo.RedirectStandardError = True
                    P.StartInfo.RedirectStandardOutput = True
                    P.StartInfo.UseShellExecute = False
                    P.StartInfo.CreateNoWindow = True
                    P.EnableRaisingEvents = True
                    Application.DoEvents()

                    AddHandler P.ErrorDataReceived, AddressOf errDataRec
                    AddHandler P.OutputDataReceived, AddressOf outpDataRec
                    AddHandler P.Exited, AddressOf cleanup
                    P.Start()
                    P.PriorityClass = ProcessPriorityClass.High

                    P.BeginErrorReadLine()
                    P.BeginOutputReadLine()
                Catch ex3 As Exception
                    MsgBox("Couldn't start Python :" & ex.Message)

                End Try

            End Try
            end try


            ' MsgBox(strRead)
            'MsgBox(OutputStr)

    End Sub
    Public P As Process

    Public errorFlag = False
    Public pythonP$
    Sub cleanup()
        Try

            If CLEANUP_IS_OKAY Then IO.File.Delete(pythonP$)
        Catch ex As Exception

        End Try
    End Sub
    Sub errDataRec(ByVal sender As Object, ByVal e As DataReceivedEventArgs)
        log &= "!!!!" & Replace(e.Data, vbLf, vbNewLine) & vbNewLine   '("Error occurred: " & vbNewLine & e.Data)
        ' MsgBox("Error: " & vbNewLine & e.Data)
        Debug.WriteLine(e.Data)
        isBusy = False
        errorFlag = True
        'Debug.WriteLine("!" & e.Data)'
    End Sub


    Public tmp$ = "", log$ = "", prog1$ = "0 out of 0", prog2$ = "0%"

    Public applyChanges = False
    Public addD$ = ""

    Public recordingMode = False
    Public edata$ = ""

    Public lastResultFoo$ = ""
    Sub outpDataRec(ByVal sender As Object, ByVal e As DataReceivedEventArgs)

        Debug.WriteLine(e.Data)

        If InStr(e.Data, "[") > 0 Then
            edata = e.Data
            recordingMode = True
            If InStr(e.Data, "]") > 0 Then
                recordingMode = False
            Else
                Exit Sub
            End If

        ElseIf InStr(e.Data, "]") > 0 Then
            edata &= " " & e.Data
            recordingMode = False
            'Exit Sub
        ElseIf recordingMode Then
            edata &= " " & e.Data
            Exit Sub
        End If

        'clean  data
        log = ""
        edata = Trim( _
            Replace(Replace(edata, "[", ""), "]", "") _
            )

        edata = Replace(edata, Space(2), Space(1))
        edata = Replace(edata, Space(2), Space(1))
        'log = Replace(log, Space(2), Space(1))

        Dim alpha = -1
        '  Dim foo = Split(edata, Space(1))
        If edata Is Nothing Then Exit Sub
        Dim foo = System.Text.RegularExpressions.Regex.Split(edata, "\s+")

        'we are parsing the coefficients here
        If InStr(edata, "@", CompareMethod.Text) = 0 Then

            If foo.Length < 1 Or edata = "" Then Exit Sub
            lastResultFoo$ = edata & ""

            If chosenMode = 5 Then 'sin 
                If foo.Length >= 1 Then a = Double.Parse(foo(0), Globalization.CultureInfo.InvariantCulture)
                If foo.Length >= 2 Then c = Double.Parse(foo(1), Globalization.CultureInfo.InvariantCulture)
                If foo.Length >= 3 Then eoff = Double.Parse(foo(2), Globalization.CultureInfo.InvariantCulture)
            Else

                If foo.Length >= 1 Then a = Double.Parse(foo(0), Globalization.CultureInfo.InvariantCulture)
                If foo.Length >= 2 Then b = Double.Parse(foo(1), Globalization.CultureInfo.InvariantCulture)
                If foo.Length >= 3 Then c = Double.Parse(foo(2), Globalization.CultureInfo.InvariantCulture)
                If foo.Length >= 4 Then d = Double.Parse(foo(3), Globalization.CultureInfo.InvariantCulture)
                ' If foo.Length >= 5 Then eoff = Double.Parse(foo(4), Globalization.CultureInfo.InvariantCulture)
                eoff = Double.Parse(If(foo.Last = "", foo(foo.Length - 2), foo.Last), Globalization.CultureInfo.InvariantCulture)


            End If

        Else 'we are expecting external data
            '  Debug.WriteLine("!!" & e.Data & "!!")


            edata = Replace(Replace(Replace(Replace(Replace(edata, "@", ""), "'", ""), ",", ""), "[", ""), "]", "")
            foo = Split(Replace(edata, Space(2), Space(1)), Space(1))
            'list_y.Clear()
            ReDim curpts(list_x.Count - 1)
            For Each item In foo
                alpha += 1
                Try
                    curpts(alpha) = Double.Parse(item, Globalization.CultureInfo.InvariantCulture)
                Catch ex As Exception
                    Try
                        curpts(alpha) = Double.Parse(item, Globalization.CultureInfo.CreateSpecificCulture("fr-FR"))
                        'list_y.Add(Double.Parse(item, )
                    Catch ex1 As Exception

                    End Try
                Finally
                    '   curpts = list_pred_pts.ConvertAll(Function(x) x.Y).ToArray

                End Try
            Next

        End If
        '  If foo.Length = 4 Then eoff = d 'hack

        applyChanges = True
        'MainForm.isBusy = False

        'log &= Replace(e.Data, vbLf, vbNewLine) & vbNewLine
        ' log = Strings.Right(log, 1000000.0)

        ' If InStr(log, "@") Then

        ' Split(log, vbNewLine)
        ' MsgBox(edata)

        ' MainForm.isBusy = False
        ' End If



        '    If InStr(e.Data, "-" & vbNewLine & "|") > 0 And InStr(e.Data, "-" & vbNewLine & vbNewLine) Then
        '    'tmp = Split(e.Data, "-" & vbNewLine & "|")(1)
        '    tmp = Split(tmp, "----" & vbNewLine)(0)
        '    log = tmp
        '    End If


    End Sub
End Module

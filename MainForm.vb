Imports System.Reflection
Imports MathNet.Numerics
Imports MathNet.Numerics.Data
Imports MathNet.Numerics.Data.Text
Imports MathNet.Numerics.LinearAlgebra
Imports OxyPlot
Imports OxyPlot.PlotModel
Imports OxyPlot.Axes
Imports System.Globalization
Imports System.ComponentModel

Public Class MainForm

    Private isFormLoading As Boolean = True

    Private Property chosenFunction As Func(Of Integer, Double)
    Private Property chosenFunctionStr$

    Private Sub MainForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If InStr(Command(), "-sc", CompareMethod.Text) > 0 Then CLEANUP_IS_OKAY = False


        'This is darned, put it on the top
        Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.InvariantCulture


        'init plotmodel
        plotModel = New PlotModel() With {.Title = "n/a"}
        plotModel.Background = OxyColors.White
        plotModel.Axes.Add(New LinearAxis() With {.Position = AxisPosition.Bottom, .MaximumPadding = 0.1, .MinimumPadding = 0.1})
        plotModel.Axes.Add(New LinearAxis() With {.Position = AxisPosition.Left, .MaximumPadding = 0.1, .MinimumPadding = 0.1})




        Plot1.Model = plotModel 'activate


        'ComboBox3.SelectedIndex = 0

        GetType(Panel).InvokeMember("DoubleBuffered",
  BindingFlags.SetProperty Or BindingFlags.Instance Or BindingFlags.NonPublic,
  Nothing, Panel1, New Object() {True})

        vircons.Init()
        vircons.Hide()

        isFormLoading = False
    End Sub
    'TODO: tab as sep
    Private Sub Form1_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Me.DragDrop, TextBox1.DragDrop, Button1.DragDrop
        Try
            Dim v = e.Data.GetData(DataFormats.FileDrop, False)
            Dim bar = Replace(v(0), Split(v(0), ".").Last, "csv", , , CompareMethod.Text)
            'Dim
            Dim bar2 = Replace(v(0), Split(v(0), ".").Last, "log", , , CompareMethod.Text)
            Dim bar3 = Replace(v(0), Split(v(0), ".").Last, "mat", , , CompareMethod.Text)

            TextBox1.Text = ""
            If IO.File.Exists(bar) Then
                TextBox1.Text = bar
            ElseIf IO.File.Exists(bar2) Then
                TextBox1.Text = bar2
            ElseIf IO.File.Exists(bar3) Then
                TextBox1.Text = bar3
            Else
                TextBox1.Text = v(0)
            End If

        Catch ex As Exception
            MsgBox(ex.Message & vbNewLine & ex.StackTrace)
        End Try
    End Sub

    Private Sub Form1_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Me.DragEnter, TextBox1.DragEnter, Button1.DragEnter
        On Error GoTo last
        If e.Data.GetDataPresent(DataFormats.FileDrop) Or e.Data.GetDataPresent(DataFormats.Text) Then
            Dim x = e.Data.GetData(DataFormats.FileDrop)(0)
            Dim bar = Replace(x, Split(x, ".").Last, "csv", , , CompareMethod.Text)
            Dim bar2 = Replace(x, Split(x, ".").Last, "log", , , CompareMethod.Text)
            Dim bar3 = Replace(x, Split(x, ".").Last, "mat", , , CompareMethod.Text)
            If IO.File.Exists(bar) Or IO.File.Exists(bar2) Or IO.File.Exists(bar3) Then
                e.Effect = DragDropEffects.Copy
            ElseIf IO.File.Exists(bar) And Split(bar, ".").Last.ToLower = "txt" Then
                e.Effect = DragDropEffects.Copy
            Else


                e.Effect = DragDropEffects.None

            End If

        Else
last:
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If ofd.ShowDialog = Windows.Forms.DialogResult.OK Then
            TextBox1.Text = ofd.FileName
        End If
    End Sub

    Dim isLoading = False
    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        isBusy = True
        If IO.File.Exists(TextBox1.Text) Then
            p(TextBox1.Text)
        End If
        isBusy = False
        UpdateParameterWindow()

    End Sub

    Public M As Matrix(Of Double)
    Public initM As Matrix(Of Double) 'default Matrix


    Dim ios As IO.StreamReader
    Dim line$, keys$()
    Dim sep$ = ";"

    Dim allList() As List(Of Double)
    Public Sub p(ByVal str$)

        If Strings.Right(str, 3).ToLower() = "log" Then
            'SPICE log, convert it to CSV
            Dim csvPath$ = str & ".csv"
            spiceToCSV(str, csvPath)
            str = csvPath 'copy filename is case of conflict
        ElseIf Strings.Right(str, 3).ToLower() = "mat" Then
            GoTo read_mat
        End If

        'open stream
        Dim fstrr = New IO.FileStream(str, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.Read, 4096)
        ios = New IO.StreamReader(fstrr, True)

        'read first line (header)
        line = ios.ReadLine() ', vbTab, Space(1)), Space(2), " ")




        'open in math.net
        Dim hasHeader As Boolean = False


        Dim number_of_numbers = countNumbers(line)
        Dim number_of_alphabets = countAlphabets(line)
        If number_of_alphabets > number_of_numbers Then hasHeader = True

        Dim number_of_semicols = countChar(line, ";")
        Dim number_of_commas = countChar(line, ",")
        Dim number_of_semisAndTabs = countStr(line, ";" & vbTab)
        Dim number_of_Tabs = countChar(line, vbTab)
        Dim number_of_space = countChar(line, Space(1))

        If number_of_semisAndTabs > 0 Then
            sep = ";" & vbTab
        ElseIf number_of_semicols > 0 Then
            sep = ";"
        ElseIf number_of_commas > 0 And hasHeader Then
            sep = ","
            'ElseIf number_of_space + number_of_Tabs > number_of_commas Then
        Else
            sep = "\s"
        End If


        'extract headers
        If sep <> "\s" Then
            keys = Split(line, sep)
        Else
            keys = Split(Replace(Replace(line, vbTab, Space(1)), Space(2), Space(1)), " ")

        End If

        If keys.Last = "" Then ReDim Preserve keys(keys.Count - 2)

        allkeys = keys.Clone()

        ' fstrr.Position = 0 'rewind

        M = Nothing


        Dim troubles#()

        Try
            fstrr.Close()
            ios.Close()
            fstrr = Nothing
            ios = Nothing
            fstrr = New IO.FileStream(str, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.Read, 4096)
            ios = New IO.StreamReader(fstrr, True)


            M = DelimitedReader.Read(Of Double)(ios, False, sep, hasHeader, Globalization.CultureInfo.InvariantCulture)

        Catch exa As Exception
            Debug.WriteLine(exa.StackTrace)
            Debug.WriteLine(exa.Message)
        End Try



        Try
            If M Is Nothing Then
                fstrr = Nothing
                ios = Nothing
                fstrr = New IO.FileStream(str, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.Read, 4096)
                ios = New IO.StreamReader(fstrr, True)

                M = DelimitedReader.Read(Of Double)(ios, False, sep, hasHeader, Globalization.CultureInfo.InvariantCulture)
            End If
        Catch ex As Exception
            Try
                ios.BaseStream.Position = 0 'rewind
                If hasHeader Then ios.ReadLine()
                M = DelimitedReader.Read(Of Double)(ios, False, sep, hasHeader, Globalization.CultureInfo.CreateSpecificCulture("fr-FR"))
            Catch ex2 As Exception
                'Manual loading
                ios.BaseStream.Position = 0
                ' If hasHeader Then ios.ReadLine()

                Dim mm As New List(Of Double())
                Try
                    If sep = "\s" Then sep = " "
                    While ios.EndOfStream = False
                        Try
                            line = Trim(Replace(Replace(ios.ReadLine(), vbTab, Space(1)), Space(2), Space(1)))

                            troubles = Nothing
                            keys = Split(line, sep)
                            If keys.Count = 0 Then Continue While
                            ReDim troubles#(allkeys.Count - 1)
                            ' troubles# = New Double(allkeys.Count - 1) {0.0}

                            For k = 0 To Math.Min(keys.Count - 1, allkeys.Count - 1)

                                Try
                                    troubles(k) = Double.Parse(keys(k), Globalization.CultureInfo.InvariantCulture)

                                Catch ex3 As Exception
                                    Try
                                        troubles(k) = Double.Parse(keys(k), Globalization.CultureInfo.CreateSpecificCulture("de-DE"))
                                    Catch ex4 As Exception
                                        Continue While
                                    End Try
                                End Try
                            Next
                        Catch ex3 As Exception
                            Continue While
                        End Try
                        If troubles IsNot Nothing Then If mm.Contains(troubles) = False Then mm.Add(troubles)


                        '  m.Add()
                    End While

                    Try
                        'IO.File.AppendAllText("kso.txt", mm.Count & "." & allkeys.Count)
                        M = Matrix(Of Double).Build.Dense(mm.Count, allkeys.Count)
                        For row = 0 To mm.Count - 1
                            '   IO.File.AppendAllText("kso.txt", "row:" & row & ":" & mm(row).Count() & vbNewLine)
                            M.SetRow(row, mm(row))
                        Next

                    Catch exkso As Exception

                        '   IO.File.AppendAllText("kso.txt", exkso.Message & exkso.StackTrace)

                    End Try
                Catch ex3 As Exception
                    MsgBox("Fatal error, couldn't load the file", MsgBoxStyle.Critical)
                    MsgBox(ex3.Message & ex3.StackTrace)
                    Exit Sub
                End Try


                'MsgBox("has header:" & hasHeader)
                'MsgBox("sep:" & sep)
                'MsgBox(ex2.Message & vbNewLine & ex2.StackTrace, MsgBoxStyle.Critical)

                isBusy = False

                '  Exit Sub
            End Try
        End Try

        If M.ColumnCount = 1 Then
            M = M.InsertColumn(0, Vector(Of Double).Build.Dense(M.RowCount, Function(i) i))
            Dim tt = allkeys(0).Clone()
            ReDim allkeys(1)
            allkeys(0) = "index"
            If hasHeader = False Then allkeys(1) = "Value" Else allkeys(1) = tt
        End If

        'initM
        initM = M.Clone()


        ios.Close()




        'redim mins,max
        ReDim v_maxs(allkeys.Count - 1)
        ReDim v_mins(allkeys.Count - 1)
        ReDim v_excludes(allkeys.Count - 1)

        'get max/min
        For col = 0 To M.ColumnCount - 1
            v_maxs(col) = M.Column(col).Maximum
            v_mins(col) = M.Column(col).Minimum
        Next
        GoTo filldata


read_mat:  'in case it's a mat

        Dim fstrMat As New IO.FileStream(str, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.ReadWrite)
        Dim aa =
            MathNet.Numerics.Data.Matlab.MatlabReader.List(fstrMat)




        Dim dict As New Dictionary(Of String, Matrix(Of Double))
        For Each item In aa
            Try
                If Trim(item.Name) = "" Then
                    dict.Add("Ans", MathNet.Numerics.Data.Matlab.MatlabReader.Read(Of Double)(str, item.Name))

                Else
                    dict.Add(item.Name, MathNet.Numerics.Data.Matlab.MatlabReader.Read(Of Double)(str, item.Name))

                End If

            Catch ex As Exception

            End Try

        Next
        '     Dim dict As Dictionary(Of String, Matrix(Of Double)) =
        '     MathNet.Numerics.Data.Matlab.MatlabReader.ReadAll(Of Double)(str)


        allkeys = Nothing
        ReDim allkeys(dict.Count - 1)
        ReDim v_maxs(dict.Count - 1)
        ReDim v_mins(dict.Count - 1)
        ReDim v_excludes(dict.Count - 1)
        Dim curIt = 0

        Dim curMaxIdx = 0
        For Each kv In dict
            allkeys(curIt) = kv.Key

            If kv.Value.ColumnCount > kv.Value.RowCount Then
                v_maxs(curIt) = kv.Value.Row(0).Maximum
                v_mins(curIt) = kv.Value.Row(0).Minimum
                ' For col = 0 To kv.Value.ColumnCount - 1
                ' v_maxs(curIt) = kv.Value.Column(col).Maximum
                ' v_mins(curIt) = kv.Value.Column(col).Minimum
                'Next
                If curMaxIdx < kv.Value.ColumnCount Then curMaxIdx = kv.Value.ColumnCount
            Else
                v_maxs(curIt) = kv.Value.Column(0).Maximum
                v_mins(curIt) = kv.Value.Column(0).Minimum

                If curMaxIdx < kv.Value.RowCount Then curMaxIdx = kv.Value.RowCount
            End If
            curIt += 1
        Next
        M = Matrix(Of Double).Build.Dense(allkeys.Count, curMaxIdx)

        curIt = 0
        For Each kv In dict
            If kv.Value.ColumnCount > kv.Value.RowCount Then

                M.SetRow(curIt, kv.Value.Row(0).ToArray)
            Else
                M.SetRow(curIt, kv.Value.Column(0).ToArray)

            End If

            curIt += 1
        Next
        M = M.Transpose
        initM = M.Clone()


filldata:
        'Fill data
        DataGridView1.Rows.Clear()

        Dim tmp1$, tmp2$
        For k = 0 To allkeys.Count - 1
            tmp1$ = wordIt(v_mins(k))
            tmp2$ = wordIt(v_maxs(k))
            DataGridView1.Rows.Add(New String() {allkeys(k), tmp1$, tmp2$})
        Next

        ComboBox1.Items.Clear()
        ComboBox2.Items.Clear()
        ComboBox3.Items.Clear()

        ComboBox1.Items.AddRange(allkeys)
        ComboBox2.Items.AddRange(allkeys)
        ComboBox3.Items.AddRange(allkeys)

        'ios.BaseStream.Positidelimitedon = 0 'rewind
        'ios.ReadLine() 'headers
    End Sub
    Public Sub gop()

        If M Is Nothing Then Exit Sub



        list_x = New List(Of Double)
        list_y = New List(Of Double)
        list_pts = New List(Of DataPoint)
        list_pred_pts = New List(Of DataPoint)
        avg_x = 0
        avg_y = 0


        If selected_idx_x = -1 Or selected_idx_y = -1 Then Exit Sub


        'update titles etc
        XAXIS_TITLE = allkeys(selected_idx_x)
        YAXIS_TITLE = allkeys(selected_idx_y)
        MEAS_TITLE = "measurements values of " & allkeys(selected_idx_y)
        ESTIM_TITLE = "estimated values of " & allkeys(selected_idx_y)
        TITLE = allkeys(selected_idx_y) & " in function of " & allkeys(selected_idx_x)

        ' For row = 0 To M.RowCount - 1
        list_x.AddRange(M.Column(selected_idx_x))
        list_y.AddRange(M.Column(selected_idx_y))

        ' Next





        If list_x.Count > 0 Then
            Select Case chosenMode
                Case 0 'lin
                    chosenFunction = Function(i) a * list_x(i) + b
                    chosenFunctionStr = "a.x + b"

                Case 1 'exp
                    If chosenIndex = 2 Then _
                    chosenFunction = Function(i) a * Math.Exp(b * list_x(i)) + c * Math.Exp(d * list_x(i)) + eoff# _
                    Else _
                    chosenFunction = Function(i) a * Math.Exp(b * list_x(i)) + eoff

                    chosenFunctionStr = "a.exp(bx) + c.exp(dx) + e"
                Case 2 'log
                    If chosenIndex = 2 Then _
                    chosenFunction = Function(i) a * Math.Log(b * list_x(i)) + c * Math.Log(d * list_x(i)) + eoff# _
                    Else _
                    chosenFunction = Function(i) a * Math.Log(b * list_x(i)) + eoff
                    chosenFunctionStr = "a.log(bx) + c.log(dx) + e"

                Case 3 'poly
                    chosenFunction = Function(i) MathNet.Numerics.Evaluate.Polynomial(list_x(i), polyCoef)
                    chosenFunctionStr = "polynomial"
                Case 4 ' pwr
                    If chosenIndex = 2 Then _
                     chosenFunction = Function(i) a * b ^ list_x(i) + c * d ^ list_x(i) + eoff# _
                     Else _
                     chosenFunction = Function(i) a * b ^ list_x(i) + eoff
                    chosenFunctionStr = "a.b^x + c^d + e"
                Case 5 ' sin
                    chosenFunction = Function(i) a * Math.Sin(b * list_x(i) + c) + eoff
                    chosenFunctionStr = "a.sin(b.x + c) + e"

                Case 6
                    chosenFunction = Function(i) curpts(i)
                    chosenFunctionStr = "custom"
            End Select
        End If

        outp = "Running (or error state)"
        'graphics
        For k = 0 To list_x.Count - 1

            list_pts.Add(New DataPoint(list_x(k), list_y(k)))
            '    If (k > 0 AndAlso list_x(k) >= list_pred_pts.Last.X) Or k = 0 Then _
            list_pred_pts.Add(New DataPoint(list_x(k), chosenFunction(k)))
        Next


        'New vector or errors
        Dim vmap = Vector(Of Double).Build.Dense(list_x.Count, chosenFunction)
        Dim vmeas = Vector(Of Double).Build.Dense(list_x.Count, Function(i) list_y(i))

        'Evaluating
        Dim verr = vmeas - vmap
        Try
            outp = "f(x) = " & chosenFunctionStr.ToString() & vbNewLine &
                If(chosenFunctionStr = "custom", "values : " & lastResultFoo & vbNewLine, "") &
                     "Min/Max Err: " & wordIt(verr.AbsoluteMinimum) & " -  " & wordIt(verr.AbsoluteMaximum) & vbNewLine &
                     "Avg Abs Err: " & wordIt(verr.PointwiseAbs.Average) & vbNewLine &
                      vbNewLine &
                    "Pearson: " & Statistics.Correlation.Pearson(vmeas.ToArray, vmap.ToArray) & vbNewLine &
                 "StdDev (err): " & wordIt(MathNet.Numerics.Statistics.ArrayStatistics.StandardDeviation(verr.ToArray)) & vbNewLine &
                 "RMS (err): " & wordIt(MathNet.Numerics.Statistics.ArrayStatistics.RootMeanSquare(verr.ToArray)) & vbNewLine &
                 "R² : " & MathNet.Numerics.GoodnessOfFit.RSquared(vmap, vmeas) & vbNewLine &
                 "StandardError: " & wordIt(MathNet.Numerics.GoodnessOfFit.StandardError(vmap, vmeas, 2)) & vbNewLine & vbNewLine &
                   "Error Harmonic Mean: " & wordIt(MathNet.Numerics.Statistics.ArrayStatistics.HarmonicMean(verr.ToArray)) & vbNewLine &
                   "InterquatileRangeInPlace: " & wordIt(MathNet.Numerics.Statistics.ArrayStatistics.InterquartileRangeInplace(verr.ToArray)) & vbNewLine &
                   "LowerQuartileInplace : " & wordIt(MathNet.Numerics.Statistics.ArrayStatistics.LowerQuartileInplace(verr.ToArray)) & vbNewLine &
                  vbNewLine &
                         "Cov.: " & MathNet.Numerics.Statistics.ArrayStatistics.Covariance(vmeas.ToArray, vmap.ToArray) & vbNewLine &
                     "Pearson Mx: " & Statistics.Correlation.PearsonMatrix(vmeas.ToArray, vmap.ToArray).ToString & vbNewLine &
            vbNewLine &
               "Spearman: " & Statistics.Correlation.Spearman(vmeas.ToArray, vmap.ToArray) & vbNewLine &
                     "Spearman Mx: " & Statistics.Correlation.SpearmanMatrix(vmeas.ToArray, vmap.ToArray).ToString & vbNewLine
            Dim x As New Statistics.DescriptiveStatistics(vmeas, True)
            'x.Kurtosis

            outp &= "[about meas:] " & vbNewLine & "Kurtosis: " & wordIt(x.Kurtosis) & vbNewLine &
                "Skewness: " & wordIt(x.Skewness) & vbNewLine &
                "Min: " & wordIt(x.Minimum) & vbNewLine &
                "Max: " & wordIt(x.Maximum) & vbNewLine &
                "Avg: " & wordIt(x.Mean) & vbNewLine &
                "stdDev: " & wordIt(x.StandardDeviation)
            '  "Error L1 norm: " & wordIt(verr.L1Norm()) & vbNewLine & _
            '             "Error L2 norm: " & wordIt(verr.L2Norm()) & vbNewLine & _


        Catch ex As Exception
            outp = "Error" & vbNewLine & ex.Message & vbNewLine & ex.StackTrace
        End Try



        Try
            '  ViewVars.DataGridView1.Invoke(New MethodInvoker(Sub() refreshUpdatedValsInParamWindows()))
        Catch ex As Exception
            ' MsgBox(ex.Message)
        End Try




    End Sub
    Dim outp$
    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged, ComboBox1.SelectedIndexChanged, ComboBox3.SelectedIndexChanged

        selected_idx_x = ComboBox1.SelectedIndex
        selected_idx_y = ComboBox2.SelectedIndex
        selected_idx_z = ComboBox2.SelectedIndex


        Try
            'polynoms
            If NumericUpDown3.Value > M.Column(selected_idx_y).Count - 1 Then NumericUpDown3.Value = M.Column(selected_idx_y).Count - 1
            NumericUpDown3.Maximum = M.Column(selected_idx_y).Count - 1


        Catch ex As Exception

        End Try
        'chosenMode = TabControl1.SelectedIndex
        '  Try
        Try
            If BackgroundWorker1.IsBusy Then BackgroundWorker1.CancelAsync() : BackgroundWorker1.RunWorkerAsync() Else BackgroundWorker1.RunWorkerAsync() 'refresh the darned dots

        Catch ex As Exception

        End Try
        Try
            BackgroundWorker2.CancelAsync()
        Catch ex As Exception

        End Try
        Try
            BackgroundWorker2.RunWorkerAsync()

        Catch ex As Exception

        End Try
        '  Catch ex As Exception

        '  End Try

        UpdateParameterWindow()
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit

        Dim idx = DataGridView1.SelectedRows(0).Index
        M = initM.Clone()


        Dim n1 = numberIt(DataGridView1.Rows(idx).Cells(1).Value)
        Dim n2 = numberIt(DataGridView1.Rows(idx).Cells(2).Value)


        If n1 <> v_mins(idx) Then
            DataGridView1.Rows(idx).Cells(1).Style.BackColor = Color.LightCyan

        End If


        If n2 <> v_maxs(idx) Then
            DataGridView1.Rows(idx).Cells(2).Style.BackColor = Color.LightCyan

        End If
        v_mins(idx) = n1
        v_maxs(idx) = n2

        Dim linestobeRemoved = New List(Of Integer)


        'exclude list
        Dim xo$ = DataGridView1.Rows(e.RowIndex).Cells(3).Value
        If Not (xo Is Nothing OrElse xo.Length < 1) Then
            Dim xxo = Split(Trim(Replace(Replace(xo, ",", Space(1)), Space(2), Space(1))), Space(1))
            v_excludes(idx) = New List(Of Double)
            For Each kso In xxo
                v_excludes(idx).Add(Double.Parse(numberIt(kso), Globalization.CultureInfo.InvariantCulture))
            Next
        Else
            v_excludes(idx) = Nothing
        End If


        For col = 0 To M.ColumnCount - 1
            For row = M.RowCount - 1 To 0 Step -1
                If M(row, col) > v_maxs(col) Or M(row, col) < v_mins(col) Or
                    (v_excludes(col) IsNot Nothing AndAlso v_excludes(col).Contains(M(row, col)) = True) Then
                    If M.RowCount = 1 And row = 0 Then M = Nothing : Exit Sub Else M = M.RemoveRow(row)
                End If

            Next

        Next


        'Update Max polys
        If NumericUpDown3.Value > M.Column(selected_idx_y).Count - 1 Then NumericUpDown3.Value = M.Column(selected_idx_y).Count - 1
        NumericUpDown3.Maximum = M.Column(selected_idx_y).Count - 1


        If BackgroundWorker2.IsBusy Then BackgroundWorker2.CancelAsync()
        Try
            BackgroundWorker2.RunWorkerAsync()

        Catch ex As Exception

        End Try

        'Update param window 
        UpdateParameterWindow()

    End Sub


    Dim tmpp! = 0, llcnt = 0

    Dim lasta#, lastb#, lastc#, lastd#, laste#
    Dim lastpolys#(), lastpts#()
    Dim onNextTickRefresh = False
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        'Keep track of chosen Mode
        If chosenMode <> TabControl1.SelectedIndex Then

            'update chosen index
            '& update a,b,c,d
            Select Case TabControl1.SelectedIndex
                Case 0
                    a = Double.Parse(TextBox3.Text, CultureInfo.InvariantCulture)
                    b = Double.Parse(TextBox4.Text, CultureInfo.InvariantCulture)
                Case 1 'exp
                    chosenIndex = NumericUpDown1.Value
                    a = Double.Parse(TextBox6.Text, CultureInfo.InvariantCulture)
                    b = Double.Parse(TextBox7.Text, CultureInfo.InvariantCulture)
                    c = Double.Parse(TextBox5.Text, CultureInfo.InvariantCulture)
                    d = Double.Parse(TextBox8.Text, CultureInfo.InvariantCulture)
                    eoff# = Double.Parse(TextBox13.Text, CultureInfo.InvariantCulture)

                Case 2
                    chosenIndex = NumericUpDown2.Value

                    a = Double.Parse(TextBox12.Text, CultureInfo.InvariantCulture)
                    b = Double.Parse(TextBox10.Text, CultureInfo.InvariantCulture)
                    c = Double.Parse(TextBox11.Text, CultureInfo.InvariantCulture)
                    d = Double.Parse(TextBox9.Text, CultureInfo.InvariantCulture)
                    eoff# = Double.Parse(TextBox14.Text, CultureInfo.InvariantCulture)

                Case 3
                    'alone

                Case 4 'pwr
                    chosenIndex = NumericUpDown4.Value

                    a = Double.Parse(TextBox20.Text, CultureInfo.InvariantCulture)
                    b = Double.Parse(TextBox18.Text, CultureInfo.InvariantCulture)
                    c = Double.Parse(TextBox19.Text, CultureInfo.InvariantCulture)
                    d = Double.Parse(TextBox17.Text, CultureInfo.InvariantCulture)
                    eoff# = Double.Parse(TextBox16.Text, CultureInfo.InvariantCulture)

                Case 5 'sin
                    '  chosenIndex = NumericUpDown2.Value

                    a = Double.Parse(TextBox25.Text, CultureInfo.InvariantCulture)
                    b = Double.Parse(TextBox23.Text, CultureInfo.InvariantCulture)
                    c = Double.Parse(TextBox24.Text, CultureInfo.InvariantCulture)
                    d = Double.Parse(TextBox22.Text, CultureInfo.InvariantCulture)
                    eoff# = Double.Parse(TextBox22.Text, CultureInfo.InvariantCulture)
                Case 6
                    ' Try
                    ''   ReDim curpts(list_x.Count - 1)
                    ' list_pred_pts.Add(New DataPoint(list_x(i), lastpts(i)))

                    'Catch ex As Exception

                    ' ' End Try
                    '  list_pred_pts.Clear()
                    '   If lastpts IsNot Nothing Then

                    ' For i = 0 To lastpts.Count - 1
                    '  Next
                    'list_pred_pts.AddRange(lastpts)
                    '   End If
            End Select


            applyChanges = chosenMode <> TabControl1.SelectedIndex
            isBusy = False
            chosenMode = TabControl1.SelectedIndex
            ' BackgroundWorker1.RunWorkerAsync()

            '     BackgroundWorker2.RunWorkerAsync()
            '   Plot1.InvalidatePlot(True)

        End If

        'also track a,b,c etc
        If a <> lasta Or b <> lastb Or c <> lastc Or d <> lastd Or Not polyCoef.ListAlmostEqual(lastpolys, 0.000000000001) Or laste <> eoff# Or lastpts IsNot curpts Then
            onNextTickRefresh = True
            lasta = a
            lastb = b
            lastc = c
            lastd = d
            lastpolys = polyCoef.Clone()
            laste = eoff#
            lastpts = curpts
        Else
            If onNextTickRefresh = True Then
                onNextTickRefresh = False
                Try
                    BackgroundWorker2.RunWorkerAsync()

                Catch ex As Exception

                End Try
            End If
        End If





        'a,b,c or d have changed
        If applyChanges Then
            Select Case chosenMode
                Case 1 'exp
                    TextBox6.Text = a
                    TextBox7.Text = b

                    If NumericUpDown1.Value = 2 Then
                        TextBox5.Text = c
                        TextBox8.Text = d
                        TextBox13.Text = eoff
                    Else
                        TextBox5.Text = 0
                        TextBox8.Text = 0
                        TextBox13.Text = eoff
                    End If
                Case 2 'log
                    TextBox12.Text = a
                    TextBox10.Text = b

                    If NumericUpDown2.Value = 2 Then
                        TextBox11.Text = c
                        TextBox9.Text = d
                        TextBox14.Text = eoff
                    Else
                        TextBox11.Text = 1
                        TextBox9.Text = 1
                        TextBox13.Text = eoff
                    End If
                Case 3 'poly

                    '  polyCoef = DataGridView2.Rows.Cast(Of Double)()
                    If polyCoef Is Nothing Or polyCoef.Count = 0 Then polyCoef = {0.0}

                Case 4 'pwr
                    TextBox20.Text = a
                    TextBox18.Text = b

                    If NumericUpDown4.Value = 2 Then
                        TextBox19.Text = c
                        TextBox17.Text = d
                        TextBox16.Text = eoff
                    Else
                        TextBox19.Text = 1
                        TextBox17.Text = 1
                        TextBox16.Text = eoff
                    End If


                Case 5 'sin ffff
                    TextBox25.Text = a
                    TextBox24.Text = c
                    TextBox22.Text = eoff
                    'TextBox22.Text = c
                Case 6
            End Select
            Try
                BackgroundWorker2.RunWorkerAsync()
            Catch ex As Exception

            End Try


            applyChanges = False
        End If

        'Error flag
        If errorFlag Then
            errorFlag = False
            TextBox2.Text = "Operation failed!" & vbNewLine & TextBox2.Text

            TextBox2.ForeColor = Color.Red
            Dim x As New Timers.Timer(3000)

            AddHandler x.Elapsed, Function() TextBox2.ForeColor = Color.Black
            '            MsgBox("Operation failed!")
        End If

        'Row count is definite, so we have to make sure to avoid any problem
        Try
            If M IsNot Nothing Then Label10.Text = "Current cells count: " & M.RowCount() Else Label10.Text = "Current cells count: 0"
        Catch ex As Exception

        End Try
        Label11.Visible = BackgroundWorker1.IsBusy Or isBusy
        llcnt += 1

        tmpp = Math.Sin(llcnt / 100.0) * 255
        tmpp *= Math.Sign(tmpp)
        tmpp = Math.Min(tmpp, 255)

        Label11.ForeColor = Color.FromArgb(255, 255 - tmpp, 255 - tmpp)

        chosenMode = TabControl1.SelectedIndex

        TextBox2.Text = outp
        If (vircons.TextBox1 Is Nothing) Then vircons.Init()
        If (outp IsNot Nothing AndAlso outp <> loutp) Then vircons.TextBox1.AppendText("outp [" & Now.ToShortTimeString & "]" & vbNewLine & outp & vbNewLine) : loutp = outp

        If (externals.log IsNot Nothing AndAlso last_log <> externals.log) Then vircons.TextBox1.AppendText("Log [" & Now.ToShortTimeString & "]" & vbNewLine & log & vbNewLine) : last_log = externals.log

        If (externals.edata IsNot Nothing AndAlso last_edata <> externals.edata) Then vircons.TextBox1.AppendText("edata [" & Now.ToShortTimeString & "]" & vbNewLine & edata & vbNewLine) : last_edata = externals.edata

        ' Plot1.Model = plotModel 'activate



    End Sub
    Dim loutp = "", last_log = "", last_edata = """"
    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        plot_stuffs.makeGraphics()

    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint
        If G_READY Is Nothing Then Exit Sub
        e.Graphics.CompositingQuality = Drawing2D.CompositingQuality.HighQuality
        e.Graphics.DrawImage(G_READY, New Point(0, 0))
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        If e.Cancel Then Exit Sub
        gop()
        plot_stuffs.makeGraphics()

    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged

        Try
            a = CDbl(TextBox3.Text)

        Catch ex As Exception

        End Try
    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged

        Try
            b = CDbl(TextBox4.Text)


        Catch ex As Exception

        End Try
    End Sub

    Private Sub TabPage2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage2.Click

    End Sub

    Private Sub DataGridView1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles DataGridView1.MouseDoubleClick
        If DataGridView1.SelectedRows.Count = 0 Then Exit Sub

        If e.Button = Windows.Forms.MouseButtons.Right Or e.Button = Windows.Forms.MouseButtons.Middle Then
            DataGridView1.Rows(DataGridView1.SelectedRows(0).Index).Cells(1).Style.BackColor = Color.White
            DataGridView1.Rows(DataGridView1.SelectedRows(0).Index).Cells(2).Style.BackColor = Color.White
            DataGridView1.Rows(DataGridView1.SelectedRows(0).Index).SetValues(New String() {DataGridView1.SelectedRows(0).Cells(0).Value, wordIt(initM.Column(DataGridView1.SelectedRows(0).Index).Minimum),
                                                                             wordIt(initM.Column(DataGridView1.SelectedRows(0).Index).Maximum)})

            DataGridView1_CellEndEdit(sender, New DataGridViewCellEventArgs(DataGridView1.SelectedCells(0).ColumnIndex, DataGridView1.SelectedRows(0).Index))
            '   v_mins(DataGridView1.SelectedRows(0).Index) = numberIt(DataGridView1.Rows(DataGridView1.SelectedRows(0).Index).Cells(1).Value)
            '   v_max(DataGridView1.SelectedRows(0).Index) = numberIt(DataGridView1.Rows(DataGridView1.SelectedRows(0).Index).Cells(2).Value)




            '  If BackgroundWorker2.IsBusy Then BackgroundWorker2.CancelAsync()
            '  BackgroundWorker2.RunWorkerAsync()
            '  BackgroundWorker1.RunWorkerAsync()
            'Catch ex As Exception

            ' End Try
        End If
    End Sub

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If list_x Is Nothing Or list_y Is Nothing Then Exit Sub

        Dim tmp
        If Not XYZ_MODE Then
            tmp = MathNet.Numerics.Fit.Line(list_x.ToArray, list_y.ToArray)
            ' Else
            '   tmp = MathNet.Numerics.Fit.MultiDim


        End If


        TextBox3.Text = tmp.Item2
        TextBox4.Text = tmp.Item1

        Try
            BackgroundWorker2.RunWorkerAsync()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        Try
            a = CDbl(TextBox6.Text)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        Try
            b = CDbl(sender.Text)
            '   If isFormLoading Then Exit Sub
            ' BackgroundWorker2.CancelAsync() : BackgroundWorker2.RunWorkerAsync()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        Try
            c = CDbl(sender.Text)
            '   If isFormLoading Then Exit Sub
            '     BackgroundWorker2.CancelAsync() : BackgroundWorker2.RunWorkerAsync()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged
        Try
            d = CDbl(sender.Text)
            ' If isFormLoading Then Exit Sub
            ' 'BackgroundWorker2.CancelAsync() : BackgroundWorker2.RunWorkerAsync()
        Catch ex As Exception

        End Try
    End Sub
    '  log(x+y) = log(max(x,y)) + log(1+min(x,y)/max(x,y))
    ' X = a*exp(b*x) ===log===> ln(a) + b*x
    ' Y = c*exp(d*x)
    '   log[ X + Y ]
    ' log(  max[X,Y] )   + log([1 + min[x,y]/max[x,y])
    '   '-------------------------------------------------------'
    '   log(x+y) = log(x) + log(1+exp(log(y)-log(x))).
    '             = ln(a)+b*x + log(1 + exp(ln(c) + dx - ln(a) + bx))
    '           =    ln(a)+b*x + log(1+exp(ln(c)+dx)*exp(-ln(a)+bx)
    '                ln(a) + b*x + log(1 + c^dx + (1/a)^bx)
    '
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If list_x Is Nothing Or list_y Is Nothing Then Exit Sub
        '  Dim tmp = MathNet.Numerics.Fit.LinearGeneric(Of Double())(list_x.ToArray, list_y.ToArray, _
        '                                                           Function(a, b, c, d, x) a * b * x + c * d * x)


        If NumericUpDown1.Value = 2 Then
            pythonIt(
                "def func(x, a, b, c, d, e) :" & vbNewLine &
                Space(5) & "return a*np.exp(b*x)+c*np.exp(d*x) + e", ,
                If(MouseButtons And Windows.Forms.MouseButtons.Right,
                "p0=[" & a & ", " & b & ", " & c & ", " & d & ", " & eoff & "]", ""))
        ElseIf NumericUpDown1.Value = 1 Then
            pythonIt(
                "def func(x, a, b, e) :" & vbNewLine &
                Space(5) & "return a*np.exp(b*x) + e", ,
                If(MouseButtons And Windows.Forms.MouseButtons.Right,
                "p0=[" & a & ", " & b & ", " & eoff & "]", ""))

        End If



        ' TextBox3.Text = tmp.Item2
        ' TextBox4.Text = tmp.Item1
    End Sub
    Sub pythonIt(funcstr$, Optional plot As Boolean = False, Optional weights$ = "")
        isBusy = True
        Randomize()
        pythonP$ = My.Computer.FileSystem.SpecialDirectories.Temp & "\curvfit" & Int(Rnd() * 32768) & ".py"

        Dim xtmp = IO.File.CreateText(pythonP$)
        xtmp.Write( _
 _
        generateScript(
          fromListToPython(list_x),
          fromListToPython(list_y),
         funcstr$,
         plot,
         weights) _
 _
     )
        xtmp.Close()

        startPython(pythonP$)


    End Sub

    Public chosenIndex = 0
    Private Sub NumericUpDown1_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown1.ValueChanged
        If NumericUpDown1.Value = 2 Then
            Label13.Text = "f(x) = a.exp(b.x) + c.exp(d.x) + e"
            chosenIndex = 2
        Else
            Label13.Text = "f(x) = a.exp(b.x) + e"
            chosenIndex = 1

        End If

        If isLoading Then Exit Sub
        Try
            BackgroundWorker2.RunWorkerAsync()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub NumericUpDown2_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown2.ValueChanged
        If NumericUpDown2.Value = 2 Then
            Label15.Text = "f(x) = a.log(b.x) + c.log(d.x) + e"
            chosenIndex = 2
        Else
            Label15.Text = "f(x) = a.log(b.x) + e"
            chosenIndex = 1

        End If
        If isFormLoading Then Exit Sub
        Try
            BackgroundWorker2.RunWorkerAsync()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub TextBox12_TextChanged(sender As Object, e As EventArgs) Handles TextBox12.TextChanged
        Try
            a = CDbl(sender.Text)
            ' If isFormLoading Then Exit Sub
            '  BackgroundWorker2.CancelAsync() : BackgroundWorker2.RunWorkerAsync()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged
        Try
            b = CDbl(sender.Text)
            ' If isFormLoading Then Exit Sub
            '  BackgroundWorker2.CancelAsync() : BackgroundWorker2.RunWorkerAsync()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub TextBox11_TextChanged(sender As Object, e As EventArgs) Handles TextBox11.TextChanged
        Try
            c = CDbl(sender.Text)
            'If isFormLoading Then Exit Sub
            ' BackgroundWorker2.CancelAsync() : BackgroundWorker2.RunWorkerAsync()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged
        Try
            d = CDbl(sender.Text)
            ' If isFormLoading Then Exit Sub
            '  BackgroundWorker2.CancelAsync() : BackgroundWorker2.RunWorkerAsync()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If list_x Is Nothing Or list_y Is Nothing Then Exit Sub
        '  Dim tmp = MathNet.Numerics.Fit.LinearGeneric(Of Double())(list_x.ToArray, list_y.ToArray, _
        '                                                           Function(a, b, c, d, x) a * b * x + c * d * x)

        If NumericUpDown2.Value = 2 Then
            pythonIt(
                "def func(x, a, b, c, d, e) :" & vbNewLine &
                Space(5) & "return a*np.log(b*x)+c*np.log(d*x) + e", ,
                If(MouseButtons And Windows.Forms.MouseButtons.Right,
                "p0=[" & a & ", " & b & ", " & c & ", " & d & ", " & eoff & "]", ""))
        ElseIf NumericUpDown2.Value = 1 Then
            pythonIt(
                "def func(x, a, b, e) :" & vbNewLine &
                Space(5) & "return a*np.log(b*x) + e", ,
                If(MouseButtons And Windows.Forms.MouseButtons.Right,
                "p0=[" & a & ", " & b & ", " & eoff & "]", ""))

        End If



    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Plot1.InvalidatePlot(True)
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If list_x Is Nothing Or list_y Is Nothing Then Exit Sub
        Dim tmp = MathNet.Numerics.Fit.Polynomial(list_x.ToArray, list_y.ToArray, NumericUpDown3.Value)

        DataGridView2.Rows.Clear()

        For Each item In tmp
            DataGridView2.Rows.Add(item)
        Next

        'TextBox3.Text = tmp(0)
        'TextBox4.Text = tmp(1)

        Try
            BackgroundWorker2.RunWorkerAsync()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick
        'TODO: this
    End Sub

    Private Sub DataGridView2_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellEndEdit
        polyCoef = Nothing
        ReDim polyCoef(DataGridView2.RowCount - 1)
        For row = 0 To DataGridView2.RowCount - 1
            polyCoef(row) = DataGridView2.Rows(row).Cells(0).Value
        Next
    End Sub

    Private Sub DataGridView2_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs) Handles DataGridView2.RowsAdded
        polyCoef = Nothing
        ReDim polyCoef(DataGridView2.RowCount - 1)
        For row = 0 To DataGridView2.RowCount - 1
            polyCoef(row) = DataGridView2.Rows(row).Cells(0).Value
        Next
    End Sub

    Private Sub TextBox13_TextChanged(sender As Object, e As EventArgs) Handles TextBox13.TextChanged
        Try
            eoff# = CDbl(sender.Text)
            '   If isFormLoading Then Exit Sub
            ' BackgroundWorker2.CancelAsync() : BackgroundWorker2.RunWorkerAsync()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub TextBox14_TextChanged(sender As Object, e As EventArgs) Handles TextBox14.TextChanged
        Try
            eoff# = CDbl(sender.Text)
            '   If isFormLoading Then Exit Sub
            ' BackgroundWorker2.CancelAsync() : BackgroundWorker2.RunWorkerAsync()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        If list_x Is Nothing Or list_y Is Nothing Then Exit Sub


        pythonIt(TextBox15.Text, True)


    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If expgfx.ShowDialog = Windows.Forms.DialogResult.OK Then
            If Strings.Right(expgfx.FileName, 3).ToLower() = "svg" Then
                Dim fstr = New IO.FileStream(expgfx.FileName, IO.FileMode.Create, IO.FileAccess.ReadWrite)
                Dim exp As New SvgExporter() With {.Width = 2048, .Height = 2048 * Panel1.Height / Panel1.Width}
                exp.Export(Plot1.Model, fstr)
                fstr.Flush()
                fstr.Close()
                exp = Nothing

            Else
                'PNG
                Dim factor! = 1.33
                Dim fstr = New IO.FileStream(expgfx.FileName, IO.FileMode.Create, IO.FileAccess.ReadWrite, IO.FileShare.Read)
                Dim exp As New WindowsForms.PngExporter() With {.Width = Panel1.Width * factor!, .Height = Panel1.Height * factor!, .Background = OxyColors.Transparent, .Resolution = 2400}

                exp.Resolution = 1200
                Plot1.Font = New Font(Plot1.Font.FontFamily, Plot1.Font.Size * factor * 2)
                Plot1.Scale(New SizeF(factor, factor)) ' = New Font(Plot1.Font.FontFamily, Plot1.Font.Size * factor)
                Plot1.Model.LegendTitleFontSize *= factor * 2
                For Each serie In Plot1.Model.Series : serie.FontSize *= factor * 10 : Next

                Plot1.InvalidatePlot(True)
                Plot1.Update()
                exp.Export(Plot1.Model, fstr)

                Plot1.Font = New Font(Plot1.Font.FontFamily, Plot1.Font.Size / factor / 2.0)
                Plot1.Scale(New SizeF(1 / factor, 1 / factor)) '
                Plot1.Model.LegendTitleFontSize /= factor * 2
                For Each serie In Plot1.Model.Series : serie.FontSize /= factor * 2 : Next
                Plot1.InvalidatePlot(True)

                fstr.Flush()
                fstr.Close()
                exp = Nothing


            End If
        End If
    End Sub

    Private Sub NumericUpDown4_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown4.ValueChanged
        If sender.Value = 2 Then
            Label26.Text = "f(x) = a.x^b + c.x^d + e"
            chosenIndex = 2
        Else
            Label15.Text = "f(x) = a.x^b + e"
            chosenIndex = 1

        End If
        If isFormLoading Then Exit Sub
        Try
            BackgroundWorker2.RunWorkerAsync()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub NumericUpDown3_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown3.ValueChanged

    End Sub

    Private Sub TextBox20_TextChanged(sender As Object, e As EventArgs) Handles TextBox20.TextChanged
        Try
            a = CDbl(sender.Text)
            ' If isFormLoading Then Exit Sub
            '  BackgroundWorker2.CancelAsync() : BackgroundWorker2.RunWorkerAsync()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub TextBox19_TextChanged(sender As Object, e As EventArgs) Handles TextBox19.TextChanged
        Try
            c = CDbl(sender.Text)
            ' If isFormLoading Then Exit Sub
            '  BackgroundWorker2.CancelAsync() : BackgroundWorker2.RunWorkerAsync()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub TextBox18_TextChanged(sender As Object, e As EventArgs) Handles TextBox18.TextChanged
        Try
            b = CDbl(sender.Text)
            ' If isFormLoading Then Exit Sub
            '  BackgroundWorker2.CancelAsync() : BackgroundWorker2.RunWorkerAsync()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub TextBox17_TextChanged(sender As Object, e As EventArgs) Handles TextBox17.TextChanged
        Try
            d = CDbl(sender.Text)
            ' If isFormLoading Then Exit Sub
            '  BackgroundWorker2.CancelAsync() : BackgroundWorker2.RunWorkerAsync()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub TextBox16_TextChanged(sender As Object, e As EventArgs) Handles TextBox16.TextChanged
        Try
            eoff = CDbl(sender.Text)
            ' If isFormLoading Then Exit Sub
            '  BackgroundWorker2.CancelAsync() : BackgroundWorker2.RunWorkerAsync()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        If list_x Is Nothing Or list_y Is Nothing Then Exit Sub
        '  Dim tmp = MathNet.Numerics.Fit.LinearGeneric(Of Double())(list_x.ToArray, list_y.ToArray, _
        '                                                           Function(a, b, c, d, x) a * b * x + c * d * x)

        If NumericUpDown4.Value = 2 Then
            pythonIt(
                "def func(x, a, b, c, d, e) :" & vbNewLine &
                Space(5) & "return a*x**b+c*x**d + e", ,
                If(MouseButtons And Windows.Forms.MouseButtons.Right,
                "p0=[" & a & ", " & b & ", " & c & ", " & d & ", " & eoff & "]", ""))
        ElseIf NumericUpDown4.Value = 1 Then
            pythonIt(
                "def func(x, a, b, e) :" & vbNewLine &
                Space(5) & "return a*x**b + e", ,
                If(MouseButtons And Windows.Forms.MouseButtons.Right,
                "p0=[" & a & ", " & b & ", " & eoff & "]", ""))

        End If


    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        If list_x Is Nothing Or list_y Is Nothing Then Exit Sub
        '  Dim tmp = MathNet.Numerics.Fit.LinearGeneric(Of Double())(list_x.ToArray, list_y.ToArray, _
        '                                                           Function(a, b, c, d, x) a * b * x + c * d * x)
        'ffff
        Button11_Click(sender, Nothing)
        pythonIt(
            "def func(x, a, c, d) :" & vbNewLine &
            Space(5) & "return a*np.sin(" & TextBox23.Text & "*x+c)+d",
            False,
            "p0=[" & TextBox25.Text & ", " & TextBox24.Text & ", " & TextBox22.Text & "]")


    End Sub

    Private Sub TextBox25_TextChanged(sender As Object, e As EventArgs) Handles TextBox25.TextChanged
        Try
            a = CDbl(sender.Text)
            ' If isFormLoading Then Exit Sub
            '  BackgroundWorker2.CancelAsync() : BackgroundWorker2.RunWorkerAsync()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub TextBox23_TextChanged(sender As Object, e As EventArgs) Handles TextBox23.TextChanged
        Try
            b = CDbl(sender.Text)
            ' If isFormLoading Then Exit Sub
            '  BackgroundWorker2.CancelAsync() : BackgroundWorker2.RunWorkerAsync()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub TextBox24_TextChanged(sender As Object, e As EventArgs) Handles TextBox24.TextChanged
        Try
            c = CDbl(sender.Text)
            ' If isFormLoading Then Exit Sub
            '  BackgroundWorker2.CancelAsync() : BackgroundWorker2.RunWorkerAsync()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub TextBox22_TextChanged(sender As Object, e As EventArgs) Handles TextBox22.TextChanged
        Try
            eoff = CDbl(sender.Text)
            ' If isFormLoading Then Exit Sub
            '  BackgroundWorker2.CancelAsync() : BackgroundWorker2.RunWorkerAsync()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click

        If list_x Is Nothing Or list_y Is Nothing Then Exit Sub
        TextBox25.Text = (list_y.Max - list_y.Min) / 2
        Dim avg = list_y.Average
        TextBox22.Text = avg

        Dim nrises = 0
        Dim nrises2 = 0
        Dim lastrise = 0.0
        Dim firstrise = 0.0

        Dim state = 0

        Dim cur0! = -1, cur1! = -1

        For i = 1 To list_y.Count - 1
            If state = 0 Then

                If list_y(i - 1) <= avg And list_y(i) > avg Then cur0 += list_x(i) : i += 1 : state = 1 : nrises += 1 : lastrise = list_x(i) : If (nrises = 1) Then firstrise = list_x(i - 1)
            Else
                If list_y(i - 1) < avg And list_y(i) > avg Then cur1 += list_x(i) : state = 0 : nrises2 += 1
            End If
            ' If cur0 <> -1 And cur1 <> -1 Then Exit For
        Next

        If nrises2 < nrises Then nrises -= 1 : cur0 -= lastrise

        If nrises <> 0 Then
            Dim per! = 2 * Math.PI / (cur1 - cur0) * nrises
            TextBox23.Text = per

            'TextBox24.Text = Math.Asin((list_y.First - list_y.Average) / list_y.Max)
            '   TextBox24.Text = Math.Asin((list_y.First - list_y.Average) / list_y.Max)
            TextBox24.Text = -(firstrise / (cur1 - cur0) * nrises) * Math.PI * 2

        End If
        Dim lma As New LMDotNet.LMSolver

        Dim ko = lma.FitCurve(Function(x, p) p(0) * Math.Sin(p(1) * x + p(2)) + p(3),
                        New Double() {TextBox25.Text, TextBox23.Text, TextBox24.Text, TextBox22.Text},
                       list_x.ToArray,
                       list_y.ToArray)
        If ko.Iterations = 600 Then
            ko = lma.FitCurve(Function(x, p) p(0) * Math.Sin(p(1) * x + p(2)) + p(3),
                 ko.OptimizedParameters,
                     list_x.ToArray,
                     list_y.ToArray)
        End If
        TextBox25.Text = ko.OptimizedParameters(0)
        TextBox23.Text = ko.OptimizedParameters(1)
        TextBox24.Text = ko.OptimizedParameters(2)
        TextBox22.Text = ko.OptimizedParameters(3)
        '   .Text, TextBox24.Text, .Text
        Try
            If (e IsNot Nothing) Then BackgroundWorker2.RunWorkerAsync()
        Catch ex As Exception

        End Try

    End Sub
    Sub UpdateParameterWindow()



        ViewVars.DataGridView1.Rows.Clear()
        ViewVars.DataGridView1.Columns.Clear()

        If allkeys Is Nothing OrElse allkeys.Count = 0 Then Exit Sub

        ViewVars.DataGridView1.ColumnCount = allkeys.Count + 1
        For j = 0 To allkeys.Count - 1
            ViewVars.DataGridView1.Columns(j).Name = allkeys(j)
        Next
        ViewVars.DataGridView1.Columns(allkeys.Count).Name = ComboBox2.Text & "_hat"

        For k = 0 To M.RowCount - 1
            Dim AllRowCount$() = M.Row(k).ToArray().ToList().ConvertAll(Function(dbl) CStr(dbl)).ToArray
            For l = 0 To AllRowCount.Count - 1
                AllRowCount(l) = wordIt(AllRowCount(l))
            Next
            ViewVars.DataGridView1.Rows.Add(AllRowCount$)
        Next
        refreshUpdatedValsInParamWindows()


    End Sub

    Sub refreshUpdatedValsInParamWindows()
        If ViewVars.DataGridView1 Is Nothing Then Exit Sub
        If ViewVars.DataGridView1.Rows Is Nothing Then Exit Sub
        If list_pred_pts Is Nothing OrElse list_pred_pts.Count = 0 Then Exit Sub
        For k = 0 To M.RowCount - 1
            Try
     ViewVars.DataGridView1.Rows(k).Cells(allkeys.Count).Value = wordIt(list_pred_pts(k).Y)
      
            Catch ex As Exception

            End Try
        Next
    End Sub
    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click

        UpdateParameterWindow()
        ViewVars.Show()
        ViewVars.BringToFront()

    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Options.Show()
    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        Dim headCode$ = "["
        'Dim SpiceCode$ = ""
        For r = DataGridView2.Rows.Count - 1 To 0 Step -1
            headCode &= DataGridView2.Rows(r).Cells(0).Value & ", "
        Next

        headCode &= "]" ' DataGridView2.Rows(DataGridView2.Rows.Count-1).Cells(0).Value & "]"
        While headCode.IndexOf(", ]") <> -1 : headCode = Replace(headCode, ", ]", "]",,, CompareMethod.Text) : End While
        While headCode.IndexOf("[, ") <> -1 : headCode = Replace(headCode, "[, ", "[",,, CompareMethod.Text) : End While


        Dim Fcode$ = "% " & ComboBox2.Text & "_hat = np.polyval(" & headCode & ", " & ComboBox1.Text & ")"
        Fcode$ &= vbNewLine & "# " & ComboBox2.Text & "_hat = polyval(" & headCode & ", " & ComboBox1.Text & ")"

        Clipboard.SetText(Fcode)
    End Sub


    'Fetch/Hold for polynomials
    Public HeldListOfPolys As New List(Of Double)
    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        HeldListOfPolys.Clear()
        For r = 0 To DataGridView2.Rows.Count - 1
            HeldListOfPolys.Add(DataGridView2.Rows(r).Cells(0).Value)
        Next

    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        DataGridView2.Rows.Clear()
        For Each item In HeldListOfPolys
            DataGridView2.Rows.Add(item)
        Next
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Dim exp As New WindowsForms.PngExporter() With {.Resolution = 1200}
        Clipboard.SetImage(exp.ExportToBitmap(Plot1.Model))

    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        If allkeys Is Nothing OrElse allkeys.Count = 0 Then Exit Sub
        Clipboard.SetText(Strings.Join(allkeys, ", "))
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        Label32.Visible = CheckBox1.Checked
        ComboBox3.Visible = CheckBox1.Checked
        XYZ_MODE = CheckBox1.Checked


        TextBox21.Visible = XYZ_MODE
        If XYZ_MODE Then
            Label4.Text = "ax"
            Label14.Text = "f(x)=ax.x+ay.y+b"
        Else
            Label4.Text = "a"
            Label14.Text = "f(x)=a.x+b"
        End If

    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        ' Dim x As New MPFitLib.mp_func(Function(x!) x)
        Dim lma As New LMDotNet.LMSolver
        Dim pp! = list_y.Max - list_y.Min

        If NumericUpDown1.Value = 2 Then
            Dim ko = lma.FitCurve(Function(x, p) p(0) * Math.Exp(p(1) * x) + p(2) * Math.Exp(p(3) * x) + p(4),
                          New Double() {pp, 1, pp, 1, 0},
                         list_x.ToArray,
                         list_y.ToArray)
            If ko.Iterations = 600 Then
                ko = lma.FitCurve(Function(x, p) p(0) * Math.Exp(p(1) * x) + p(2) * Math.Exp(p(3) * x) + p(4),
                     ko.OptimizedParameters,
                         list_x.ToArray,
                         list_y.ToArray)
            End If
            TextBox6.Text = ko.OptimizedParameters(0)
            TextBox7.Text = ko.OptimizedParameters(1)
            TextBox5.Text = ko.OptimizedParameters(2)
            TextBox8.Text = ko.OptimizedParameters(3)
            TextBox13.Text = ko.OptimizedParameters(4)
        Else
            Dim ko = lma.FitCurve(Function(x, p) p(0) * Math.Exp(p(1) * x) + p(2),
                   New Double() {pp, 1, 0},
                   list_x.ToArray,
                   list_y.ToArray)
            If ko.Iterations = 600 Then
                ko = lma.FitCurve(Function(x, p) p(0) * Math.Exp(p(1) * x) + p(2),
                     ko.OptimizedParameters,
                         list_x.ToArray,
                         list_y.ToArray)
            End If
            TextBox6.Text = ko.OptimizedParameters(0)
            TextBox7.Text = ko.OptimizedParameters(1)
            TextBox5.Text = 0
            TextBox8.Text = 0
            TextBox13.Text = ko.OptimizedParameters(2)
        End If
        'MsgBox(ko.Message)

    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        Dim lma As New LMDotNet.LMSolver
        Dim pp! = list_y.Max - list_y.Min

        If NumericUpDown2.Value = 2 Then
            Dim ko = lma.FitCurve(Function(x, p) p(0) * Math.Log(p(1) * x) + p(2) * Math.Log(p(3) * x) + p(4),
                         New Double() {pp, 1, pp, 1, 0},
                         list_x.ToArray,
                         list_y.ToArray)
            If ko.Iterations = 600 Then
                ko = lma.FitCurve(Function(x, p) p(0) * Math.Log(p(1) * x) + p(2) * Math.Log(p(3) * x) + p(4),
                     ko.OptimizedParameters,
                         list_x.ToArray,
                         list_y.ToArray)
            End If
            If ko.Iterations = 600 Then
                ko = lma.FitCurve(Function(x, p) p(0) * Math.Log(p(1) * x) + p(2) * Math.Log(p(3) * x) + p(4),
                     ko.OptimizedParameters,
                         list_x.ToArray,
                         list_y.ToArray)
            End If
            TextBox12.Text = ko.OptimizedParameters(0)
            TextBox10.Text = ko.OptimizedParameters(1)
            TextBox11.Text = ko.OptimizedParameters(2)
            TextBox9.Text = ko.OptimizedParameters(3)
            TextBox14.Text = ko.OptimizedParameters(4)
        Else
            Dim ko = lma.FitCurve(Function(x, p) p(0) * Math.Log(p(1) * x) + p(2),
                   New Double() {pp, 1, 0},
                   list_x.ToArray,
                   list_y.ToArray)
            If ko.Iterations = 600 Then
                ko = lma.FitCurve(Function(x, p) p(0) * Math.Log(p(1) * x) + p(2),
                     ko.OptimizedParameters,
                         list_x.ToArray,
                         list_y.ToArray)
            End If
            If ko.Iterations = 600 Then
                ko = lma.FitCurve(Function(x, p) p(0) * Math.Log(p(1) * x) + p(2),
                     ko.OptimizedParameters,
                         list_x.ToArray,
                         list_y.ToArray)
            End If
            TextBox12.Text = ko.OptimizedParameters(0)
            TextBox10.Text = ko.OptimizedParameters(1)
            TextBox11.Text = 0
            TextBox9.Text = 1
            TextBox14.Text = ko.OptimizedParameters(2)
        End If
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        Dim lma As New LMDotNet.LMSolver
        Dim pp! = list_y.Max - list_y.Min

        If NumericUpDown4.Value = 2 Then
            Dim ko = lma.FitCurve(Function(x, p) p(0) * Math.Pow(p(1), x) + p(2) * Math.Pow(p(3), x) + p(4),
                       New Double() {1, pp, 1, pp, 0},
                         list_x.ToArray,
                         list_y.ToArray)
            If ko.Iterations = 600 Then
                ko = lma.FitCurve(Function(x, p) p(0) * Math.Pow(p(1), x) + p(2) * Math.Pow(p(3), x) + p(4),
                     ko.OptimizedParameters,
                         list_x.ToArray,
                         list_y.ToArray)
            End If
            TextBox20.Text = ko.OptimizedParameters(0)
            TextBox18.Text = ko.OptimizedParameters(1)
            TextBox19.Text = ko.OptimizedParameters(2)
            TextBox17.Text = ko.OptimizedParameters(3)
            TextBox16.Text = ko.OptimizedParameters(4)
        Else
            Dim ko = lma.FitCurve(Function(x, p) p(0) * Math.Pow(p(1), x) + p(2),
                     New Double() {1, pp, 0},
                   list_x.ToArray,
                   list_y.ToArray)
            If ko.Iterations = 600 Then
                ko = lma.FitCurve(Function(x, p) p(0) * Math.Pow(p(1), x) + p(2),
                     ko.OptimizedParameters,
                         list_x.ToArray,
                         list_y.ToArray)
            End If
            TextBox20.Text = ko.OptimizedParameters(0)
            TextBox18.Text = ko.OptimizedParameters(1)
            TextBox19.Text = 0
            TextBox17.Text = 1
            TextBox16.Text = ko.OptimizedParameters(2)
        End If
    End Sub

    Function isnum(str$) As Boolean
        Return str >= "0" And str <= "9"
    End Function

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        Dim adjustedName$ = TextBox1.Text.Split("\").Last
        adjustedName = Replace(adjustedName, " ", "_")
        If isnum(adjustedName(0)) Then adjustedName = "a" & adjustedName
        adjustedName = adjustedName.Replace("-", "_").Replace(".", "_")

        adjustedName = Strings.Join(TextBox1.Text.Split("\").Take(TextBox1.Text.Split("\").Count - 1).ToArray(), "\") & "\" & adjustedName
        DelimitedWriter.Write(adjustedName$ & "_selected.csv", M, ";", allkeys.ToArray())

        'Matlab
        Dim newmatlabfilename$ = adjustedName & "_selected.m" 'Strings.Join(adjustedName$.Split("\").Take(adjustedName$.Split("\").Count - 1).ToArray(), "\") & "\" & Split(adjustedName$, "\").Last & "_selected.m"
        Dim kline = 0
        Dim fw As New IO.StreamWriter(New IO.FileStream(newmatlabfilename$, IO.FileMode.Create, IO.FileAccess.Write, IO.FileShare.ReadWrite))
        For Each key In allkeys
            fw.WriteLine(key & "=[" &
                String.Join(", ", M.Column(kline).ToArray) &
                "];")
            kline += 1
        Next
        fw.Close()

        'Python
        IO.File.Copy(newmatlabfilename, adjustedName & "_selected.py", True)




    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        Dim tmp = ComboBox1.SelectedIndex
        ComboBox1.SelectedIndex = ComboBox2.SelectedIndex
        ComboBox2.SelectedIndex = tmp

    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        vircons.Show()
        vircons.Visible = True


    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        refreshUpdatedValsInParamWindows()
    End Sub
End Class

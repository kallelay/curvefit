
Module convertSpiceToCSV
    Dim sep$ = Replace(CStr(1.1), "1", "") 'get the seperator (, or .)

    Dim dBMode As Boolean = False
    Dim dBModeList As New List(Of Integer)

    Dim ays, bys As New List(Of String)
    Dim headers As New List(Of String)
    Dim firstIterForHeaders As Boolean = True

    'Convert a (LT)SPICE log into a CSV File
    Sub spiceToCSV(ByVal spicePath$, ByRef csvPath$)
        If IO.File.Exists(spicePath$) = False Then Exit Sub

        dBMode = False
        dBModeList.Clear()

        firstIterForHeaders = True
        ays.Clear() : bys.Clear() : headers.Clear()
        'Dim allLines = IO.File.ReadAllLines(spicePath$)
        Dim allLines As String()

        Dim lsf As New List(Of String)
        Dim fo = IO.File.Open(spicePath, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.ReadWrite)
        Dim fstr As New IO.StreamReader(fo)
        While Not fstr.EndOfStream
            lsf.Add(fstr.ReadLine)
            Application.DoEvents() 'slow down, but make it interactive
        End While

        fstr.Close()
        fo.Close()
        fstr = Nothing : fo = Nothing

        allLines = lsf.ToArray : lsf = Nothing


        Dim keys$(), subkeys$()

        Dim curIter = -1
        Dim skipNextLine = False
        Dim isReading = False

        For Each line In allLines

            If skipNextLine Then skipNextLine = False : Continue For
            'step
            If InStr(line, ".step", CompareMethod.Text) = 1 Then
                keys = Split(line, " ")
                ays.Add("")
                For k = 1 To keys.Count - 1
                    subkeys = Split(keys(k), "=")
                    If firstIterForHeaders Then headers.Add(subkeys.First)
                    ays(ays.Count - 1) &= subkeys.Last & ";"
                Next
                firstIterForHeaders = False
            End If

            If InStr(line, "Measurement:", CompareMethod.Text) = 1 Then
                headers.Add(Split(line, " ").Last)
                curIter = -1
                skipNextLine = True
                isReading = True 'once 
            End If

            If isReading Then
                If InStr(line, "Date:") Then isReading = False
                If Len(line) < 3 Then Continue For
                If InStr(line, vbTab) = 0 Then Continue For
                curIter += 1

                If InStr(line, "dB") = 0 Then
                    Try
                        ays(curIter) &= CStr(Double.Parse(Split(line, vbTab)(1), Globalization.CultureInfo.InvariantCulture)) & ";"

                    Catch ex As Exception
                        curIter -= 1
                        curIter = Math.Max(curIter, 0)
                    End Try

                Else
                    'dB mode
                    line = Replace(line, "(", "")
                    line = Replace(line, "°", "")
                    line = Replace(line, ")", "")
                    line = Replace(line, ",", ";")
                    line = Replace(line, "dB", "")
                    line = Replace(line, "ｰ", "")
                    line = Replace(line, "�", "")
                    ays(curIter) &= CStr(Split(line, vbTab)(1)) & ";"
                    dBMode = True
                    If dBModeList.Contains(curIter) = False Then dBModeList.Add(curIter)

                End If



            End If


        Next

        Dim ff As IO.FileStream

        Try
            ff = IO.File.Open(csvPath, IO.FileMode.Create)
        Catch ex As Exception
            Randomize()
            csvPath$ = csvPath & Int(Rnd() * 500) & ".csv"
            ff = IO.File.Open(csvPath, IO.FileMode.OpenOrCreate)
        End Try

        Dim gg = New IO.StreamWriter(ff)

        For k = 0 To headers.Count - 1
            If dBModeList.Contains(k - keys.Count) = False Then
                gg.Write(headers(k) & If(k <> headers.Count - 1, ";", ""))
            Else
                gg.Write(headers(k) & "dB;" & headers(k) & "Ph" & If(k <> headers.Count - 1, ";", ""))
            End If
        Next
        gg.Write(vbNewLine)
        For i = 0 To ays.Count - 1
            ays(i) = Replace(ays(i), ".", sep) 'global please?
            gg.WriteLine(ays(i).Substring(0, ays(i).Count - 1))
        Next
        ' If 1 = 2 Then MsgBox("hello, stupid line to make sure this recompiles, lazy mode on")
        gg.Flush()
        gg.Close()
        ff.Close()
        ff = Nothing
        gg = Nothing
    End Sub

End Module

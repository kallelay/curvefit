Imports System.Math
Imports OxyPlot

Module tools_etc

    Public CLEANUP_IS_OKAY As Boolean = True 'always cleanup

    'Hack: allkeys is now public
    Public allkeys() As String


    Public isBusy = False


    Public list_x As List(Of Double), list_y As List(Of Double)
    Public list_pts As List(Of DataPoint)
    Public list_pred_pts As List(Of DataPoint)
    'public list_z s list(of double)
    Public avg_x#, avg_y#
    Public v_mins() As Double
    Public v_maxs() As Double
    Public v_excludes() As List(Of Double)


    Public lines(), ppts() As PointF


    'Get unit symbol
    Public Function wordIt$(ByVal str#)
        If Abs(str#) < 1.0E-18 Then
            Return str
        ElseIf Abs(str#) < 0.000000000000001 Then
            Return (str# / 1.0E-18) & "a"
        ElseIf Abs(str#) < 0.000000000001 Then
            Return (str# / 0.000000000000001) & "f"
        ElseIf Abs(str#) < 0.000000001 Then
            Return (str# / 0.000000000001) & "p"
        ElseIf Abs(str#) < 0.000001 Then
            Return (str# / 0.000000001) & "n"
        ElseIf Abs(str#) < 0.001 Then
            Return (str# / 0.000001) & "u"

        ElseIf Abs(str#) < 1 Then
            Return (str# / 0.001) & "m"

        ElseIf Abs(str#) >= 1000000000000.0 Then
            Return (str# / 1000000000000.0) & "T"

        ElseIf Abs(str#) >= 1000000000.0 Then
            Return (str# / 1000000000.0) & "G"

        ElseIf Abs(str#) >= 1000000.0 Then
            Return (str# / 1000000.0) & "M"

        ElseIf Abs(str#) >= 1000.0 Then
            Return (str# / 1000.0) & "k"

        End If
        Return str
    End Function

    'Convert units to number (TODO: cleanup, check last thing, not replace!)
    Public Function numberIt#(ByVal str$)
        str = Trim(Replace(str, Space(1), ""))
        str = Replace(str, "a", "e-18")
        str = Replace(str, "f", "e-15")
        str = Replace(str, "p", "e-12")
        str = Replace(str, "n", "e-9")
        str = Replace(str, "u", "e-6")
        str = Replace(str, "µ", "e-6")
        str = Replace(str, "m", "e-3")
        str = Replace(str, "k", "e3")
        str = Replace(str, "K", "e3")
        str = Replace(str, "MEG", "e6", , , CompareMethod.Text)
        str = Replace(str, "M", "e6")
        str = Replace(str, "G", "e9")
        str = Replace(str, "T", "e12")

        Double.TryParse(str, numberIt)

    End Function


    Function countNumbers(str$) As Integer
        countNumbers = 0
        For Each ch In str
            If ch >= "0" And ch <= "9" Then countNumbers += 1
        Next
    End Function
    Function countAlphabets(str$) As Integer
        countAlphabets = 0
        For Each ch In str
            If ch >= "a" And ch <= "z" Then countAlphabets += 1
            If ch >= "A" And ch <= "Z" Then countAlphabets += 1
        Next
    End Function

    Function countChar(str$, chrr As Char) As Integer
        countChar = 0
        For Each ch In str
            If ch = chrr Then countChar += 1
        Next
    End Function
    Function countStr(str$, chrr$) As Integer
        countStr = 0
        For i = 0 To str.Length - Len(chrr) - 1
            If str.Substring(i, Len(chrr)) = chrr Then countStr += 1
        Next
    End Function
End Module

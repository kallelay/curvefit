Module mathetc
    Function listsum#(ByVal v As List(Of Double))
        listsum# = 0
        For Each item In v
            listsum# += item
        Next
    End Function

    Function listsumMinusNumber#(ByVal v As List(Of Double), ByVal n#)
        listsumMinusNumber# = 0
        For Each item In v
            listsumMinusNumber# += item - n
        Next
    End Function
    Function listsumMinusNumberMulListMinusNumber#(ByVal v1 As List(Of Double), ByVal n1#, _
                                                   ByVal v2 As List(Of Double), ByVal n2#)
        listsumMinusNumberMulListMinusNumber# = 0
        For i = 0 To v1.Count - 1
            listsumMinusNumberMulListMinusNumber# += (v1(i) - n1) * (v2(i) - n2)
        Next
    End Function
End Module

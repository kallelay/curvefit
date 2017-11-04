Imports MathNet
Imports MathNet.Numerics
Imports MathNet.Numerics.SpecialFunctions
Imports System.Math
Module funcs
    Public Class customUserVar
        Public x#(), y#(), ey#()
    End Class
    Public Function ExpFunc(p As Double(), dy As Double(), dvec As IList(Of Double), vars As Object) As Integer
        Dim i As Integer
        Dim x, y, ey As Double()
        Dim f As Double
        Dim V = DirectCast(vars, customUserVar)


        x = V.x
        y = V.y
        ey = V.ey


        For i = 0 To dy.Length - 1
            f = p(0) * Exp(p(1) * dy(i)) + p(2) * Exp(p(3) * dy(i)) + p(4)
            dy(i) = (y(i) - f) / (list_x(2) - list_x(1))

        Next
        Return 0

    End Function
  
End Module

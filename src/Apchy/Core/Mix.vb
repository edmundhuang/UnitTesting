Namespace Collection
    Public Class NoSort
        Implements System.Collections.IComparer

        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare
            Return -1
        End Function
    End Class
End Namespace
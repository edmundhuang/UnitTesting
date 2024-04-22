Public Class ConvertHelper
    Public Shared Function ToDecimal(ByVal Original As Object) As Nullable(Of Decimal)
        Try

            Dim myString As String = Original.ToString

            If IsNumeric(Original) Then
                myString = myString.Replace("￥", "")
                myString = myString.TrimStart("$", "")

                Return CDec(myString)
            End If
        Catch ex As Exception

        End Try

        Return Nothing
    End Function

    Public Shared Function ToDouble(ByVal Original As Object) As Double
        Try

            Dim myString As String = Original.ToString

            If IsNumeric(Original) Then
                myString = myString.Replace("￥", "")
                myString = myString.TrimStart("$", "")

                Return CDbl(myString)
            End If
        Catch ex As Exception

        End Try

        Return 0
    End Function

End Class

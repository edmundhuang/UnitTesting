Namespace Helper

    Public Class DatasetHelper
        Public Shared Function String2DataTable(ByVal SourceString As String, ByVal Delimiter As String, ByRef myDataTable As DataTable) As Boolean
            Dim stringArray() As String
            Dim i As Integer
            Dim myNewRow As DataRow

            myDataTable.Clear()
            '通过加入此句，可以避免字符串为空时的错误。
            If SourceString.Length = 0 Then Exit Function

            stringArray = Split(SourceString, Delimiter)

            '减2是因为数组的最后一个为空 lee
            For i = 0 To stringArray.Length - 1
                myNewRow = myDataTable.NewRow
                myNewRow.Item(0) = stringArray(i).ToString
                myDataTable.Rows.Add(myNewRow)
            Next

            Return True
        End Function
    End Class

End Namespace
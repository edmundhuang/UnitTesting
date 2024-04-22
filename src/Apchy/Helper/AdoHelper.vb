Imports System.Data.SqlClient

Namespace Helper

    Public Class AdoHelper
        Public Shared ConnectionString As String = ""

        '传递一个SQL命令，返回SQL 命令中行集的第一行第一列的值，没有值时返回
        Public Shared Function ADONLookup(ByVal ConnectionString As String, ByVal CommendText As String) As Object
            Dim MyValue As New Object

            Try
                Using myConnection As New SqlConnection(ConnectionString)
                    myConnection.Open()

                    Dim myCommand As New SqlCommand

                    myCommand.CommandText = CommendText
                    myCommand.Connection = myConnection

                    MyValue = myCommand.ExecuteScalar()
                End Using

            Catch ex As Exception
                MyValue = Nothing
            End Try

            Return MyValue
        End Function

        Public Shared Function LoadDataTable(ByVal ConnectionString As String, ByRef myDataSet As DataSet, ByVal SQLString As String, ByVal TableName As String) As Boolean
            Try
                Using myConnection As New SqlConnection(ConnectionString)

                    myConnection.Open()

                    Dim myDA As New SqlDataAdapter(SQLString, myConnection)

                    myDA.Fill(myDataSet, TableName)
                End Using

                Return True
            Catch ex As Exception
                Throw (ex)
            End Try
        End Function

    End Class

End Namespace
Imports System.Globalization
Imports System.Windows.Forms


Namespace Core

    Public Class Definition

#Region "HashTable Function"

        Public Shared Function TransferConfigToHashTable(ByVal ConfigString As String, Optional ByVal Delimeter As String = "&", Optional ByVal Equal As String = "=", Optional ByVal ValueDelimeter As String = ";") As Hashtable
            Dim myTable As New Hashtable
            ''Dim ValueDelimeter As String = ";"

            If ConfigString Is Nothing Then Return myTable
            If ConfigString.Length = 0 Then Return myTable

            Dim ConfigArray() As String

            ConfigArray = Split(ConfigString, Delimeter)

            Dim I As Integer
            Dim myFieldAndValue() As String
            Dim myField As String = ""
            Dim myValue As String = ""

            For I = 0 To ConfigArray.Length - 1
                Try
                    myFieldAndValue = Split(ConfigArray(I), Equal, 2)

                    myField = myFieldAndValue(0) & ""
                    myValue = myFieldAndValue(1) & ""

                    If myField.Length <> 0 Then
                        Try
                            myTable.Add(myField, myValue)
                        Catch ex As Exception
                            myValue = myTable(myField) & ValueDelimeter & myValue

                            myTable.Add(myField, myValue)
                        End Try
                    End If
                Catch ex As Exception

                End Try
            Next

            Return myTable
        End Function

        Public Shared Function TransferHashTableToConfig(ByVal HashTable As Hashtable, Optional ByVal Delimeter As String = "&", Optional ByVal Equal As String = "=")
            Dim ReturnString As String = ""

            Dim myField As String = ""
            Dim myValue As String = ""

            For Each key As String In HashTable.Keys
                myField = key
                myValue = HashTable.Item(myField)

                If ReturnString.Length = 0 Then
                    ReturnString = myField & Equal & myValue
                Else
                    ReturnString = ReturnString & Delimeter & myField & Equal & myValue
                End If
            Next

            Return ReturnString
        End Function
#End Region

    End Class

End Namespace
Imports System.Net
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions

Namespace Helper

    Public Class TCPHelper
        Public Shared Function LocalIPAddress() As String
            Dim strHostName As String = System.Net.Dns.GetHostName()
            Dim strIPAddress As String = System.Net.Dns.GetHostEntry(strHostName).AddressList(0).ToString()

            Return strIPAddress
        End Function

        Public Shared Function ExternalIP() As String
            Return ExternalIP("http://service.puda.net.cn/ip.asp")
        End Function

        Public Shared Function ExternalIP(ByVal URL As String) As String
            Dim wc As WebClient = New WebClient
            Dim utf8 As UTF8Encoding = New UTF8Encoding
            Dim requestHtml As String = ""
            Dim ip As IPAddress = Nothing

            Try
                requestHtml = utf8.GetString(wc.DownloadData(URL))

                Return requestHtml
            Catch ex As Exception

            End Try

            Return "127.0.0.1"
        End Function
    End Class

End Namespace
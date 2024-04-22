Namespace Helper

    Public Class DigitalHelper
        ''' <summary>
        ''' 用途：将十进制转化为二进制
        ''' 输入：Dec(十进制数)
        ''' 输入数据类型：Long
        ''' 输出：DEC_to_BIN(二进制数)
        ''' 输出数据类型：String
        ''' 输入的最大数为2147483647,输出最大数为1111111111111111111111111111111(31个1)
        ''' </summary>
        ''' <param name="Value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Decimal2BinArray(ByVal Value As Nullable(Of Long)) As Integer()
            Dim ReturnString As String = Decimal2Bin(Value)

            ReturnString = StrReverse(ReturnString)

            Dim iLength As Integer = ReturnString.Length

            Dim ReturnArray(iLength - 1) As Integer

            Dim I As Integer

            For I = 0 To iLength - 1
                ReturnArray(I) = ReturnString.Substring(I, 1)
            Next

            Return ReturnArray
        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="Value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Decimal2Bin(ByVal Value As Nullable(Of Long)) As String
            Dim ReturnString As String = ""

            If Value Is Nothing Then Return ReturnString

            Do While Value > 0
                ReturnString = Value Mod 2 & ReturnString
                Value = Value \ 2
            Loop

            Return ReturnString
        End Function
    End Class

End Namespace
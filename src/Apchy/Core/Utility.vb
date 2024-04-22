Imports System.Reflection
Imports System.Text
Imports System.Security.Cryptography
Imports System.IO

Namespace Core

    Public Class Utility

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

        Public Shared Function isFormOpened(ByVal FormType As System.Type) As Object
            For Each frmTemp In System.Windows.Forms.Application.OpenForms
                If frmTemp.Equals(FormType) Then
                    Return frmTemp
                End If
            Next

            Return Nothing
        End Function

        Public Shared Function SingleForm(ByVal FormType As System.Type) As Object
            Dim myObject As Object

            For Each frmTemp In System.Windows.Forms.Application.OpenForms

                If frmTemp.Equals(FormType) Then
                    Return frmTemp
                Else

                End If
            Next

            myObject = System.Activator.CreateInstance(FormType)
            Return myObject
        End Function

        Public Shared Function SingleForm(ByVal FormType As System.Type, ByRef MDIForm As System.Windows.Forms.Form) As Object
            Dim myObject As Object

            For Each frmTemp In MDIForm.MdiChildren
                If frmTemp.GetType.Equals(FormType) Then
                    frmTemp.Focus()
                    Return frmTemp
                End If
            Next

            myObject = System.Activator.CreateInstance(FormType)
            Return myObject
        End Function

        'Public Shared Sub ShowForm(ByVal FormType As System.Type)
        '    Dim frmTemp = isFormOpened(FormType)

        '    If frmTemp IsNot Nothing Then

        '    Else

        '    End If
        'End Sub

        Public Shared Function IsNumericData(ByVal str As Object) As Boolean
            If (IsNumeric(str)) Then
                Return True
            Else
                Return False
            End If
        End Function

        Public Shared Function IsGuidData(ByVal str As String) As Boolean
            Try
                Dim tmp As Guid = New Guid(str)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function

        Public Shared Function GetBillNo(ByVal prefix As String) As String
            Dim sID As String = prefix
            sID += DateTime.Now.Year.ToString().Substring(2)
            sID += DateTime.Now.Month.ToString("00")
            sID += DateTime.Now.Day.ToString("00")
            sID += DateTime.Now.Hour.ToString("00")
            sID += DateTime.Now.Minute.ToString("00")
            sID += DateTime.Now.Millisecond.ToString("000")
            Return sID
        End Function

        Public Shared Function GetPYCode(ByVal str As String) As String
            Dim pinyin As Char
            Dim c As Char
            Dim array() As Byte
            Dim i As Integer
            Dim sb As New System.Text.StringBuilder(str.Length)
            For Each c In str.ToCharArray
                pinyin = c
                array = System.Text.Encoding.Default.GetBytes(New Char() {c})
                If array.Length = 2 Then
                    i = array(0) * &H100 + array(1)
                    If i < &HB0A1 Then
                        pinyin = c
                    ElseIf i < &HB0C5 Then
                        pinyin = "a"
                    ElseIf i < &HB2C1 Then
                        pinyin = "b"
                    ElseIf i < &HB4EE Then
                        pinyin = "c"
                    ElseIf i < &HB6EA Then
                        pinyin = "d"
                    ElseIf i < &HB7A2 Then
                        pinyin = "e"
                    ElseIf i < &HB8C1 Then
                        pinyin = "f"
                    ElseIf i < &HB9FE Then
                        pinyin = "g"
                    ElseIf i < &HBBF7 Then
                        pinyin = "h"
                    ElseIf i < &HBFA6 Then
                        pinyin = "g"
                    ElseIf i < &HC0AC Then
                        pinyin = "k"
                    ElseIf i < &HC2E8 Then
                        pinyin = "l"
                    ElseIf i < &HC4C3 Then
                        pinyin = "m"
                    ElseIf i < &HC5B6 Then
                        pinyin = "n"
                    ElseIf i < &HC5BE Then
                        pinyin = "o"
                    ElseIf i < &HC6DA Then
                        pinyin = "p"
                    ElseIf i < &HC8BB Then
                        pinyin = "q"
                    ElseIf i < &HC8F6 Then
                        pinyin = "r"
                    ElseIf i < &HCBFA Then
                        pinyin = "s"
                    ElseIf i < &HCDDA Then
                        pinyin = "t"
                    ElseIf i < &HCEF4 Then
                        pinyin = "w"
                    ElseIf i < &HD1B9 Then
                        pinyin = "x"
                    ElseIf i < &HD4D1 Then
                        pinyin = "y"
                    ElseIf i < &HD7FA Then
                        pinyin = "z"
                    End If
                End If
                sb.Append(pinyin)
            Next
            Return sb.ToString()
        End Function


        '默认密钥向量 
        Private Shared Keys As Byte() = {18, 52, 86, 120, 144, 171, 205, 239}
        ''' <summary> 
        ''' DES加密字符串 
        ''' </summary> 
        ''' <param name="encryptString">待加密的字符串</param> 
        ''' <returns>加密成功返回加密后的字符串，失败返回源串</returns> 
        Public Shared Function EncryptDES(ByVal encryptString As String) As String
            Try
                Dim encryptKey As String = "apkzyqkf"
                Dim rgbKey As Byte() = Encoding.UTF8.GetBytes(encryptKey.Substring(0, 8))
                Dim rgbIV As Byte() = Keys
                Dim inputByteArray As Byte() = Encoding.UTF8.GetBytes(encryptString)
                Dim dCSP As New DESCryptoServiceProvider()
                Dim mStream As New MemoryStream()
                Dim cStream As New CryptoStream(mStream, dCSP.CreateEncryptor(rgbKey, rgbIV), CryptoStreamMode.Write)
                cStream.Write(inputByteArray, 0, inputByteArray.Length)
                cStream.FlushFinalBlock()
                Return Convert.ToBase64String(mStream.ToArray())
            Catch
                Return encryptString
            End Try
        End Function


        ''' <summary> 
        ''' DES解密字符串 
        ''' </summary> 
        ''' <param name="decryptString">待解密的字符串</param> 
        ''' <returns>解密成功返回解密后的字符串，失败返源串</returns> 
        Public Shared Function DecryptDES(ByVal decryptString As String) As String
            Try
                Dim decryptKey As String = "apkzyqkf"
                Dim rgbKey As Byte() = Encoding.UTF8.GetBytes(decryptKey)
                Dim rgbIV As Byte() = Keys
                Dim inputByteArray As Byte() = Convert.FromBase64String(decryptString)
                Dim DCSP As New DESCryptoServiceProvider()
                Dim mStream As New MemoryStream()
                Dim cStream As New CryptoStream(mStream, DCSP.CreateDecryptor(rgbKey, rgbIV), CryptoStreamMode.Write)
                cStream.Write(inputByteArray, 0, inputByteArray.Length)
                cStream.FlushFinalBlock()
                Return Encoding.UTF8.GetString(mStream.ToArray())
            Catch
                Return decryptString
            End Try
        End Function

    End Class

End Namespace
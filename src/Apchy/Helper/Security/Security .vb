Imports System.Security.Cryptography
Imports System.IO
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Text

Namespace Security

#Region "校验位算法"
    Public Class VerifyCode
        Public Shared Function GenerateCode(ByVal OriginalString As String) As Integer
            Dim myCode As String = MD5Codec.EncryptString(OriginalString)

            If myCode.Length >= 1 Then
                myCode = myCode.Substring(0, 1)
            End If

            Dim myASC = Asc(myCode)

            Return myASC Mod 10
        End Function
    End Class
#End Region

#Region "DES 加密与解密"

    Public Class DES
        Inherits DESCodec
    End Class

    Public Class DESCodec
        Public Shared Function CreateCrptoTransform(ByVal DesKey() As Byte, ByVal DESIV() As Byte) As ICryptoTransform
            Dim objDES As New DESCryptoServiceProvider()

            Return objDES.CreateEncryptor(DesKey, DESIV)
        End Function

        Public Shared Function CreateDecryptTransform(ByVal DesKey() As Byte, ByVal DESIV() As Byte) As ICryptoTransform
            Dim objDES As New DESCryptoServiceProvider()

            Return objDES.CreateDecryptor(DesKey, DESIV)
        End Function

        'DES加解密算法
        '   <summary>  
        '  加密算法  
        '   </summary>
        Public Shared Function EncryptString(ByVal pToEncrypt As String, ByVal sKey As String) As String
            Dim des As New DESCryptoServiceProvider
            Dim inputByteArray() As Byte
            inputByteArray = Encoding.Default.GetBytes(pToEncrypt)
            ''建立加密对象的密钥和偏移量
            ''原文使用ASCIIEncoding.ASCII方法的GetBytes方法
            ''使得输入密码必须输入英文文本
            des.Key = ASCIIEncoding.ASCII.GetBytes(sKey)
            des.IV = ASCIIEncoding.ASCII.GetBytes(sKey)
            ''写二进制数组到加密流
            ''(把内存流中的内容全部写入)
            Dim ms As New System.IO.MemoryStream
            Dim cs As New CryptoStream(ms, des.CreateEncryptor, CryptoStreamMode.Write)
            ''写二进制数组到加密流
            ''(把内存流中的内容全部写入)
            cs.Write(inputByteArray, 0, inputByteArray.Length)
            cs.FlushFinalBlock()
            ''建立输出字符串    
            Dim ret As New StringBuilder
            Dim b As Byte
            For Each b In ms.ToArray()
                ret.AppendFormat("{0:X2}", b)
            Next
            Return ret.ToString()
        End Function

        '   <summary>  
        '   解密算法  
        '   </summary>
        Public Shared Function DecryptString(ByVal pToDecrypt As String, ByVal sKey As String) As String

            If pToDecrypt.Length = 0 Then Return ""

            Dim des As New DESCryptoServiceProvider
            ''把字符串放入byte数组
            Dim len As Integer
            len = pToDecrypt.Length / 2 - 1
            Dim inputByteArray(len) As Byte
            Dim x, i As Integer
            For x = 0 To len
                i = Convert.ToInt32(pToDecrypt.Substring(x * 2, 2), 16)
                inputByteArray(x) = CType(i, Byte)
            Next
            ''建立加密对象的密钥和偏移量，此值重要，不能修改
            des.Key = ASCIIEncoding.ASCII.GetBytes(sKey)
            des.IV = ASCIIEncoding.ASCII.GetBytes(sKey)
            Dim ms As New System.IO.MemoryStream
            Dim cs As New CryptoStream(ms, des.CreateDecryptor, CryptoStreamMode.Write)
            cs.Write(inputByteArray, 0, inputByteArray.Length)
            cs.FlushFinalBlock()
            Return Encoding.Default.GetString(ms.ToArray, 0, ms.Length)
        End Function

        Public Shared Sub EncryptToBinary(ByVal myObject As Object, ByVal outFilePath As String, ByVal CryptoTransform As ICryptoTransform)
            Dim fout As New FileStream(outFilePath, FileMode.OpenOrCreate, FileAccess.Write)

            '用指定的 Key 和初始化向量 (IV) 创建对称数据加密标准 (DES) 加密器对象 
            Dim objCryptoStream As New CryptoStream(fout, CryptoTransform, CryptoStreamMode.Write)

            myObject.RemotingFormat = SerializationFormat.Binary
            Dim bf As New BinaryFormatter()

            bf.Serialize(objCryptoStream, myObject)
            objCryptoStream.Close()

            fout.Close()
        End Sub

        Public Shared Function DecryptFromBinary(ByVal inBinFilePath As String, ByVal CryptoTransform As ICryptoTransform) As Object
            Dim fin As New FileStream(inBinFilePath, FileMode.Open, FileAccess.Read)

            '用指定的 Key 和初始化向量 (IV) 创建对称数据加密标准 (DES) 加密器对象 
            Dim objCryptoStream As New CryptoStream(fin, CryptoTransform, CryptoStreamMode.Read)
            Dim myBinaryFormatter As New BinaryFormatter()

            Try
                Dim myObject As Object = myBinaryFormatter.Deserialize(objCryptoStream)
                Return myObject
            Catch ex As Exception

            Finally
                fin.Close()
            End Try

            Return Nothing
        End Function

    End Class
#End Region

#Region "MD5 加密解密"
    Public Class MD5
        Inherits MD5Codec
    End Class

    Public Class MD5Codec
        Public Shared Function EncryptString(ByVal OriginalString As String, Optional ByVal Code As Integer = 16) As String
            Dim dataToHash As Byte() = (New System.Text.ASCIIEncoding).GetBytes(OriginalString)
            Dim hashvalue As Byte() = CType(System.Security.Cryptography.CryptoConfig.CreateFromName("MD5"), System.Security.Cryptography.HashAlgorithm).ComputeHash(dataToHash)

            Dim ReturnString As String = ""

            Dim I As Integer
            Select Case Code
                Case 16  '选择16位字符的加密结果
                    For I = 4 To 11
                        ReturnString += Hex(hashvalue(I)).ToLower
                    Next
                Case 32  '选择32位字符的加密结果
                    For I = 0 To 15
                        ReturnString += Hex(hashvalue(I)).ToLower
                    Next
                Case Else   'Code错误时，返回全部字符串，即32位字符
                    For I = 0 To hashvalue.Length - 1
                        ReturnString += Hex(hashvalue(I)).ToLower
                    Next
            End Select

            Return ReturnString
        End Function
    End Class
#End Region

#Region "RSA 加密解密"
    Public Class RSACodec
        Public Shared Function BuildPrivateKeyAndPublicKey() As String()
            Dim rsa As RSACryptoServiceProvider
            rsa = New RSACryptoServiceProvider

            Dim strPrivateKey As String
            Dim strPublicKey As String
            strPrivateKey = rsa.ToXmlString(True)
            strPublicKey = rsa.ToXmlString(False)

            Dim str(1) As String
            str(0) = strPrivateKey
            str(1) = strPublicKey
            Return str
        End Function

        Public Shared Function EncryptCode(ByVal xmlKey As String, ByVal OriginalString As String) As String
            Dim rsa As RSACryptoServiceProvider
            rsa = New RSACryptoServiceProvider

            rsa.FromXmlString(xmlKey)
            Dim f As RSAPKCS1SignatureFormatter
            f = New RSAPKCS1SignatureFormatter(rsa)
            f.SetHashAlgorithm("SHA1")

            Dim source() As Byte
            source = System.Text.ASCIIEncoding.ASCII.GetBytes(OriginalString)

            Dim sha As SHA1Managed
            sha = New SHA1Managed
            Dim result() As Byte
            result = sha.ComputeHash(source)

            Dim b() As Byte
            b = f.CreateSignature(result)

            Dim strRegeditCode As String
            strRegeditCode = Convert.ToBase64String(b)
            Return strRegeditCode
        End Function

        'RSA的解密函数
        Public Shared Function DecryptCode(ByVal xmlKey As String, ByVal DecryptString As String) As String
            Dim PlainTextBArray As Byte()
            PlainTextBArray = Convert.FromBase64String(DecryptString)

            Return DecryptCode(xmlKey, PlainTextBArray)
        End Function

        'RSA的解密函数
        Public Shared Function DecryptCode(ByVal xmlKey As String, ByVal DecryptString As Byte()) As String
            Try
                Dim DypherTextBArray As Byte()
                Dim Result As String
                Dim rsa As System.Security.Cryptography.RSACryptoServiceProvider = New RSACryptoServiceProvider()
                rsa.FromXmlString(xmlKey)
                DypherTextBArray = rsa.Decrypt(DecryptString, False)
                Result = (New UnicodeEncoding()).GetString(DypherTextBArray)
                Return Result
            Catch ex As Exception
                Throw ex
            End Try
        End Function

        Public Shared Function ValidateRegeditCode(ByVal xmlKey As String, ByVal strSerialCode As String, ByVal strRegeditCode As String) As Boolean
            Dim rsa As RSACryptoServiceProvider
            rsa = New RSACryptoServiceProvider
            rsa.FromXmlString(xmlKey)
            Dim f As RSAPKCS1SignatureDeformatter
            f = New RSAPKCS1SignatureDeformatter(rsa)
            f.SetHashAlgorithm("SHA1")
            Dim key() As Byte
            key = Convert.FromBase64String(strRegeditCode)
            Dim sha As SHA1Managed
            sha = New SHA1Managed
            Dim name() As Byte
            name = sha.ComputeHash(System.Text.ASCIIEncoding.ASCII.GetBytes(strSerialCode))
            If f.VerifySignature(name, key) = True Then
                Return True   '注册码正确   
            Else
                Return False   '注册码不正确   
            End If
        End Function
    End Class

#End Region

#Region "Rijndael 加密"
    Public Class RijndaelCodec
        Private Shared mobjCryptoService As SymmetricAlgorithm = New RijndaelManaged()
        Private Shared Key As String = "Guz(%&hj7x89H$yuBI0456FtmaT5&fvHUFCy76*h%(HilJ$lhj!y6&(*jkP87jH7"

        ''' <summary>
        ''' 获得密钥
        ''' </summary>
        ''' <returns>密钥</returns>
        Private Shared Function GetLegalKey() As Byte()
            Dim sTemp As String = Key

            mobjCryptoService.GenerateKey()

            Dim bytTemp As Byte() = mobjCryptoService.Key

            Dim KeyLength As Integer = bytTemp.Length

            If sTemp.Length > KeyLength Then

                sTemp = sTemp.Substring(0, KeyLength)

            ElseIf sTemp.Length < KeyLength Then

                sTemp = sTemp.PadRight(KeyLength, " "c)
            End If

            Return ASCIIEncoding.ASCII.GetBytes(sTemp)

        End Function

        ''' <summary>
        ''' 获得初始向量IV
        ''' </summary>
        ''' <returns>初试向量IV</returns>
        Private Shared Function GetLegalIV() As Byte()
            Dim sTemp As String = "E4ghj*Ghg7!rNIfb&95GUY86GfghUb#er57HBh(u%g6HJ($jhWk7&!hg4ui%$hjk"

            mobjCryptoService.GenerateIV()

            Dim bytTemp As Byte() = mobjCryptoService.IV

            Dim IVLength As Integer = bytTemp.Length

            If sTemp.Length > IVLength Then

                sTemp = sTemp.Substring(0, IVLength)

            ElseIf sTemp.Length < IVLength Then

                sTemp = sTemp.PadRight(IVLength, " "c)
            End If

            Return ASCIIEncoding.ASCII.GetBytes(sTemp)

        End Function

        ''' <summary>
        ''' 加密方法
        ''' </summary>
        ''' <param name="Source">待加密的串</param>
        ''' <returns>经过加密的串</returns>

        Public Shared Function Encrypto(ByVal Source As String) As String
            Dim bytIn As Byte() = UTF8Encoding.UTF8.GetBytes(Source)

            Dim ms As New MemoryStream()

            mobjCryptoService.Key = GetLegalKey()

            mobjCryptoService.IV = GetLegalIV()

            Dim myEncrypto As ICryptoTransform = mobjCryptoService.CreateEncryptor()

            Dim cs As New CryptoStream(ms, myEncrypto, CryptoStreamMode.Write)

            cs.Write(bytIn, 0, bytIn.Length)

            cs.FlushFinalBlock()

            ms.Close()

            Dim bytOut As Byte() = ms.ToArray()

            Return Convert.ToBase64String(bytOut)

        End Function

        ''' <summary>
        ''' 解密方法
        ''' </summary>
        ''' <param name="Source">待解密的串</param>
        ''' <returns>经过解密的串</returns>
        Public Shared Function Decrypto(ByVal Source As String) As String
            Dim bytIn As Byte() = Convert.FromBase64String(Source)

            Dim ms As New MemoryStream(bytIn, 0, bytIn.Length)

            mobjCryptoService.Key = GetLegalKey()

            mobjCryptoService.IV = GetLegalIV()

            Dim encrypto As ICryptoTransform = mobjCryptoService.CreateDecryptor()

            Dim cs As New CryptoStream(ms, encrypto, CryptoStreamMode.Read)

            Dim sr As New StreamReader(cs)

            Return sr.ReadToEnd()

        End Function
    End Class
#End Region

End Namespace
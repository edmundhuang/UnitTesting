Imports System.Text
Imports System.Security.Cryptography
Imports System.IO
Imports System.Runtime.Serialization.Formatters.Binary

Namespace Helper

    Public Class Security

#Region "MD5"

        Public Shared Function MD5(ByVal OriginalString As String, Optional ByVal Code As Integer = 16) As String
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

#End Region

#Region "DES"

        '密钥 
        '获取或设置对称算法的机密密钥。机密密钥既用于加密，也用于解密。为了保证对称算法的安全，必须只有发送方和接收方知道该机密密钥。有效密钥大小是由特定对称算法实现指定的，密钥大小在 LegalKeySizes 中列出。 
        Private Shared DESKey As Byte() = New Byte() {11, 23, 93, 102, 72, 41, 18, 12}

        '获取或设置对称算法的初始化向量 
        Private Shared DESIV As Byte() = New Byte() {75, 158, 46, 97, 78, 57, 17, 36}



        Public Shared Sub EncryptDataSetToBinary(ByVal objDataSet As DataSet, ByVal outFilePath As String)

            Dim objDES As New DESCryptoServiceProvider()
            Dim fout As New FileStream(outFilePath, FileMode.OpenOrCreate, FileAccess.Write)
            '用指定的 Key 和初始化向量 (IV) 创建对称数据加密标准 (DES) 加密器对象 
            Dim objCryptoStream As New CryptoStream(fout, objDES.CreateEncryptor(DESKey, DESIV), CryptoStreamMode.Write)

            'StreamWriter objStreamWriter = new StreamWriter(objCryptoStream); 

            objDataSet.RemotingFormat = SerializationFormat.Binary
            Dim bf As New BinaryFormatter()


            'bf.Serialize(fout, objDataSet)

            bf.Serialize(objCryptoStream, objDataSet)
            objCryptoStream.Close()

            fout.Close()
        End Sub

        'Public Shared Sub SaveXML(ByVal objdataset As DataSet, ByVal outfilePath As String)
        '    objdataset.WriteXml(outfilePath, XmlWriteMode.WriteSchema)
        'End Sub

        Public Shared Function DecryptDataSetFromBinary(ByVal inBinFilePath As String) As DataSet
            Dim objDES As New DESCryptoServiceProvider()
            Dim fin As New FileStream(inBinFilePath, FileMode.Open, FileAccess.Read)
            '用指定的 Key 和初始化向量 (IV) 创建对称数据加密标准 (DES) 加密器对象 
            Dim objCryptoStream As New CryptoStream(fin, objDES.CreateDecryptor(DESKey, DESIV), CryptoStreamMode.Read)

            Dim myBinaryFormatter As New BinaryFormatter()

            'Dim ds As CacheData = DirectCast(myBinaryFormatter.Deserialize(fin), DataSet)

            Try
                Dim ds As DataSet = DirectCast(myBinaryFormatter.Deserialize(objCryptoStream), DataSet)
                Return ds
            Catch ex As Exception

            Finally
                fin.Close()
            End Try

            Return Nothing
        End Function

        


        'DES加解密算法
        '   <summary>  
        '  加密算法  
        '   </summary>
        Public Shared Function DesEncrypt(ByVal pToEncrypt As String, ByVal sKey As String) As String
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
        Public Shared Function DesDecrypt(ByVal pToDecrypt As String, ByVal sKey As String) As String

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

#End Region

    End Class

End Namespace
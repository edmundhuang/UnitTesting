Imports System.Security.Cryptography
Imports System.IO
Imports System.Runtime.Serialization.Formatters.Binary

Namespace Helper

    Public Class SecurityHelper

#Region "加密文件与对象的相互转换"

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

#End Region

#Region "公钥与私钥相关处理"
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

        Public Shared Function BuildRegisterCode(ByVal strPrivateKey As String, ByVal strSerialCode As String) As String
            Dim rsa As RSACryptoServiceProvider
            rsa = New RSACryptoServiceProvider

            rsa.FromXmlString(strPrivateKey)
            Dim f As RSAPKCS1SignatureFormatter
            f = New RSAPKCS1SignatureFormatter(rsa)
            f.SetHashAlgorithm("SHA1")
            Dim source() As Byte
            source = System.Text.ASCIIEncoding.ASCII.GetBytes(strSerialCode)
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

        Public Shared Function ValidateRegeditCode(ByVal strPublicKey As String, ByVal strSerialCode As String, ByVal strRegeditCode As String) As Boolean
            Dim rsa As RSACryptoServiceProvider
            rsa = New RSACryptoServiceProvider
            rsa.FromXmlString(strPublicKey)
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

#End Region



    End Class

End Namespace
Imports System.Xml.Serialization
Imports System.IO
Imports System.Text

Namespace Helper
    Public Class SerializeHelper
        Public Shared Sub Save(ByVal Model As Object, ByVal FileName As String)
            Try
                Serialize(Model, FileName)
            Catch ex As Exception
                Throw ex
            End Try
        End Sub

        Public Shared Function Load(ByVal Path As String, ByVal Type As Type) As Object
            Try
                Return Deserialize(Path, Type)
            Catch ex As Exception
                Throw ex
            End Try
        End Function

        Public Shared Sub Serialize(ByVal Model As Object, ByVal Path As String)
            Dim fs As FileStream = Nothing

            Try
                Dim xs As XmlSerializer = New XmlSerializer(Model.GetType)
                fs = New FileStream(Path, FileMode.Create, FileAccess.Write)

                xs.Serialize(fs, Model)

                fs.Close()
            Catch ex As Exception
                If fs IsNot Nothing Then
                    fs.Close()
                End If

                Throw ex
            End Try
        End Sub

        Public Shared Function Deserialize(ByVal Path As String, ByVal Type As Type) As Object
            Dim xs As XmlSerializer = New XmlSerializer(Type)

            Dim ReturnObject As Object = Nothing
            Dim fs As FileStream = Nothing

            Try
                fs = New FileStream(Path, FileMode.Open, FileAccess.Read)

                ReturnObject = xs.Deserialize(fs)
                fs.Close()
            Catch ex As Exception
                If fs IsNot Nothing Then
                    fs.Close()
                End If

                Throw ex
            End Try

            Return ReturnObject
        End Function

        ''' <summary>
        ''' 序列化 对象到字符串
        ''' </summary>
        ''' <typeparam name="T">泛型类型</typeparam>
        ''' <param name="obj">泛型对象</param>
        ''' <returns>序列化后的字符串</returns>
        ''' <remarks></remarks>
        Shared Function Serialize(Of T)(ByVal obj As T) As String
            Try
                Dim Serializer As New XmlSerializer(GetType(T))
                Dim ns As New XmlSerializerNamespaces
                Dim stream As New MemoryStream
                Dim writer As New StreamWriter(stream, Encoding.UTF8)
                Serializer.Serialize(writer, obj, ns)
                stream.Position = 0
                Dim buf(stream.Length) As Byte
                stream.Read(buf, 0, buf.Length)
                Return Encoding.UTF8.GetString(buf)
            Catch ex As Exception
                Throw New Exception("序列化失败,原因:" & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' 反序列化 字符串到对象
        ''' </summary>
        ''' <typeparam name="T">泛型类型</typeparam>
        ''' <param name="str">要转换为对象的字符串</param>
        ''' <returns>反序列化出来的对象</returns>
        ''' <remarks></remarks>
        Shared Function Desrialize(Of T)(ByVal str As String) As T
            Try
                Dim Serializer As New XmlSerializer(GetType(T))
                Dim buffer() As Byte = Encoding.UTF8.GetBytes(str)
                Dim stream As New MemoryStream(buffer)
                Dim obj As T = Serializer.Deserialize(stream)
                Return obj
            Catch ex As Exception
                Throw New Exception("反序列化失败,原因:" & ex.Message)
            End Try
        End Function
    End Class
End Namespace
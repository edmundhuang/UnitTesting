Imports System.Reflection
Imports System.IO

Public Class ReflectionHelper
    Public Shared Function DynamicType(ByVal FileName As String, ByVal TypeName As String) As Type
        Try
            If Not File.Exists(FileName) Then Return Nothing

            Dim myFileInfo As FileInfo = New FileInfo(FileName)

            Return DynamicType(myFileInfo, TypeName)
        Catch ex As Exception

        End Try

        Return Nothing
    End Function

    Public Shared Function DynamicInstance(ByVal FileName As String, ByVal TypeName As String) As Object
        Try
            If Not File.Exists(FileName) Then Return Nothing

            Dim myFileInfo As FileInfo = New FileInfo(FileName)

            Return DynamicInstance(myFileInfo, TypeName)
        Catch ex As Exception

        End Try

        Return Nothing
    End Function

    Public Shared Function DynamicType(ByVal FileInfo As FileInfo, ByVal TypeName As String) As Type
        If FileInfo Is Nothing Then Return Nothing

        Try
            Dim myAssemble As Reflection.Assembly = [Assembly].LoadFrom(FileInfo.FullName)

            If myAssemble Is Nothing Then Return Nothing

            Dim myType As Type = myAssemble.GetType(TypeName)    '"Apchy.Business.ProductClass"

            Return myType
        Catch ex As Exception

        End Try

        Return Nothing
    End Function

    Public Shared Function DynamicInstance(ByVal FileInfo As FileInfo, ByVal TypeName As String) As Object
        Try
            Dim myType As Type = DynamicType(FileInfo, TypeName)

            If myType IsNot Nothing Then
                Return System.Activator.CreateInstance(myType)
            End If
        Catch ex As Exception

        End Try

        Return Nothing
    End Function

    Public Shared Function ProcessStaticMethod(ByVal Type As Type, ByVal MethodName As String, Optional ByVal ParamArgs() As Object = Nothing) As Object
        Try
            Dim myMethod As MethodInfo = Type.GetMethod(MethodName)

            If myMethod IsNot Nothing Then
                Return myMethod.Invoke(Type, BindingFlags.Static + BindingFlags.InvokeMethod, Nothing, ParamArgs, Nothing)
            End If
        Catch ex As Exception

        End Try

        Return Nothing
    End Function

    Public Shared Function ProcessInstanceMethod(ByVal Type As Type, ByVal MethodName As String, Optional ByVal ParamArgs() As Object = Nothing) As Object
        Try
            Dim myMethod As MethodInfo = Type.GetMethod(MethodName)

            Dim myInstance As Object = System.Activator.CreateInstance(Type)

            If myMethod IsNot Nothing Then
                Return myMethod.Invoke(myInstance, BindingFlags.Instance + BindingFlags.InvokeMethod, Nothing, ParamArgs, Nothing)
            End If
        Catch ex As Exception

        End Try

        Return Nothing
    End Function
End Class


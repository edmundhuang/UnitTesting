Imports System
Imports System.Collections.Generic
Imports System.Reflection
Imports System.Text
Imports System.Diagnostics

Namespace Helper

    Public Class PropertyHelper

#Region "Set Properties"
        Public Shared Sub CopyProperties(ByVal fromFields As PropertyInfo(), ByVal toFields As PropertyInfo(), ByVal fromRecord As Object, ByVal toRecord As Object)
            Dim fromField As PropertyInfo = Nothing
            Dim toField As PropertyInfo = Nothing

            Try

                If fromFields Is Nothing Then
                    Return
                End If
                If toFields Is Nothing Then
                    Return
                End If
                For f As Integer = 0 To fromFields.Length - 1


                    fromField = DirectCast(fromFields(f), PropertyInfo)
                    For t As Integer = 0 To toFields.Length - 1
                        Debug.Print(toFields.ToString)

                        toField = DirectCast(toFields(t), PropertyInfo)

                        If fromField.Name <> toField.Name Then
                            Continue For
                        End If

                        toField.SetValue(toRecord, fromField.GetValue(fromRecord, Nothing), Nothing)

                        Exit For

                    Next

                Next
            Catch generatedExceptionName As Exception
                'Throw
            End Try
        End Sub
#End Region

#Region "Set Properties"
        Public Shared Sub CopyProperties(ByVal fromFields As PropertyInfo(), ByVal fromRecord As Object, ByVal toRecord As Object)
            Dim fromField As PropertyInfo = Nothing

            Try

                If fromFields Is Nothing Then
                    Return
                End If
                For f As Integer = 0 To fromFields.Length - 1


                    fromField = DirectCast(fromFields(f), PropertyInfo)

                    fromField.SetValue(toRecord, fromField.GetValue(fromRecord, Nothing), Nothing)

                Next
            Catch generatedExceptionName As Exception
                'Throw
            End Try
        End Sub
#End Region

        Public Shared Sub SetPropertyValue(ByRef myObject As Object, ByVal FieldName As String, ByVal Value As Object)
            Try
                 Dim myType As Type = myObject.GetType

                Dim myInfo As PropertyInfo = myType.GetProperty(FieldName)

                If myInfo IsNot Nothing Then
                    myInfo.SetValue(myObject, Value, Nothing)
                End If
            Catch ex As Exception
                Throw ex
            End Try
        End Sub

        Public Shared Function GetPropertyValue(ByRef myObject As Object, ByVal Fieldname As String) As Object
            If myObject Is Nothing Then Return Nothing

            For Each prop In myObject.GetType.GetProperties
                If prop.Name = Fieldname Then
                    Return prop.GetValue(myObject, Nothing)
                End If
            Next

            Return Nothing
        End Function

        Public Shared Function GetObjectPropertyValue(ByRef myObject As Object, ByVal Fieldname As String) As Object
            If myObject Is Nothing Then Return Nothing

            For Each prop In myObject.GetType.GetProperties
                If prop.Name = Fieldname Then
                    Return prop.GetValue(myObject, Nothing)
                End If
            Next

            Return Nothing
        End Function

        Public Shared Function GetUnderlyingType(ByVal ColumnType As Type)
            If ColumnType.IsGenericType AndAlso ColumnType.GetGenericTypeDefinition() Is GetType(Nullable(Of )) Then
                Return ColumnType.GetGenericArguments()(0)
            Else
                Return ColumnType
            End If
        End Function

        Public Shared Function GetObjectProperty(ByRef myObject As Object, ByVal Fieldname As String) As System.Reflection.PropertyInfo
            For Each prop In myObject.GetType.GetProperties
                If prop.Name.ToLower = Fieldname.ToLower Then
                    Return prop
                End If
            Next

            Return Nothing
        End Function

        'PropertyHandler Sample For Identical Classes 
        'MyClass record = new MyClass();
        'MyClass newRecord = new MyClass();
        'PropertyInfo[] fromFields = null;

        'fromFields = typeof(MyClass).GetProperties();

        'PropertyHandler.SetProperties(fromFields, record, newRecord);


        'PropertyHandler Sample For Similar Classes 

        'MyClass record = new MyClass();
        'MyOtherClass newRecord = new MyOtherClass();
        'PropertyInfo[] fromFields = null;
        'PropertyInfo[] toFields = null;

        'fromFields = typeof(MyClass).GetProperties();
        'toFields = typeof(MyOtherClass).GetProperties();

        'PropertyHandler.SetProperties(fromFields,toFields,record, newRecord);



    End Class

End Namespace
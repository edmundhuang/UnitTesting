Imports System.Windows.Forms

Public Class FormHelper
    Public Shared Function GetActiveControl(ByVal ActiveContainer As ContainerControl)
        Dim myControl As Control = ActiveContainer.ActiveControl

        If Not myControl.Parent.Equals(ActiveContainer) Then
            myControl = myControl.Parent
        End If

        Try
            If myControl.HasChildren Then
                myControl = GetActiveControl(myControl)
            End If
        Catch ex As Exception

        End Try

        Return myControl
    End Function
End Class

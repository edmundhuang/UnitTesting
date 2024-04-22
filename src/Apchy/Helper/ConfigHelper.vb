
Namespace Helper

    Public Class ConfigHelper
        Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Int32, ByVal lpFileName As String) As Int32
        Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Int32

        Public Shared Function GetINI(ByVal Section As String, ByVal AppName As String, ByVal lpDefault As String, ByVal FileName As String) As String
            Dim ReturnString As String = ""

            ReturnString = LSet(ReturnString, 256)

            GetPrivateProfileString(Section, AppName, lpDefault, ReturnString, Len(ReturnString), FileName)

            Return Left(ReturnString, InStr(ReturnString, Chr(0)) - 1)
        End Function

        Public Shared Function WriteINI(ByVal Section As String, ByVal AppName As String, ByVal lpDefault As String, ByVal FileName As String) As Long
            WriteINI = WritePrivateProfileString(Section, AppName, lpDefault, FileName)
        End Function
    End Class

    'Dim path As String
    '    path = Application.StartupPath + "\Send.ini"
    '    TextBox1.Text = GetINI("Send", "Send1", "", path)
    '    TextBox2.Text = GetINI("Send", "Send2", "", path)
    'Dim IsSms As Integer = GetINI("Send", "IsSms", "", path)
    '    If (IsSms = 1) Then
    '        Me.RadioButton1.Checked = True
    '    ElseIf (IsSms = 0) Then
    '        Me.RadioButton2.Checked = True
    '    End If

End Namespace
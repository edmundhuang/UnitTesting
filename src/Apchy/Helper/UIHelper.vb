Imports System.Windows.Forms

Namespace Helper
    Public Class UIHelper
        Public Shared Sub TrimEditor(ByVal Sender As System.Object)
            Try
                Dim myEditor As Control = TryCast(Sender, Control)

                If myEditor IsNot Nothing Then
                    myEditor.Text = myEditor.Text.Trim
                End If
            Catch ex As Exception

            End Try
        End Sub

        Public Shared Sub StoreException(ByVal e As Exception)
            Dim sStr As String

            sStr = DateTime.Now.ToShortTimeString & " -- 错误原因：" & e.Message & System.Environment.NewLine

            Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory

            If sPath.Substring(sPath.Length - 1, 1) = "\" Then
                sPath = sPath.Substring(0, sPath.Length - 1)
            End If

            Dim sFolderPath As String = sPath + "\Log"
            If Not System.IO.Directory.Exists(sFolderPath) Then
                System.IO.Directory.CreateDirectory(sFolderPath)
            End If

            Dim _sFilePath As String = sPath + "\Log\Err.log"
            Dim stream As System.IO.StreamWriter = System.IO.File.AppendText(_sFilePath)
            stream.Write(sStr)
            stream.Close()
        End Sub

        Public Shared Sub StoreLog(ByVal LogString As String)

            LogString = DateTime.Now.ToLongTimeString & "       " & LogString & System.Environment.NewLine

            Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
            If sPath.Substring(sPath.Length - 1, 1) = "\" Then
                sPath = sPath.Substring(0, sPath.Length - 1)
            End If

            Dim sFolderPath As String = sPath + "\Log"
            If Not System.IO.Directory.Exists(sFolderPath) Then
                System.IO.Directory.CreateDirectory(sFolderPath)
            End If

            Dim sDateString As String = DateTime.Now.Year & "-" & DateTime.Now.Month & "-" & DateTime.Now.Day

            Dim _sFilePath As String = sPath + "\Log\Cash" & sDateString & ".log"

            Dim stream As System.IO.StreamWriter = System.IO.File.AppendText(_sFilePath)
            stream.Write(LogString)
            stream.Close()
        End Sub

        Public Shared Sub BindingSourceEndEdit(ByRef BindingSourceCollection As List(Of BindingSource))

            Dim I As Integer

            For I = BindingSourceCollection.Count - 1 To 0 Step -1
                Try
                    TryCast(BindingSourceCollection(I), BindingSource).EndEdit()
                Catch ex As Exception

                End Try
            Next
        End Sub

        Public Shared Function GetActiveControl(ByVal Winform As Form) As Control
            Dim myControl As Control

            myControl = Winform.ActiveControl

            If myControl Is Nothing Then Return Nothing

            While True
                Dim myContainer As IContainerControl = TryCast(myControl, IContainerControl)

                If myContainer Is Nothing Then
                    Return myControl
                Else
                    myControl = myContainer.ActiveControl
                End If

            End While

            Return myControl
        End Function

        Public Shared Function GetPYCode(ByVal OriginalString As String) As String

            Dim sOne As String = ""
            Dim sChar As String = ""
            Dim sRet As String = ""
            Dim I As Integer

            Try

                For I = 1 To Len(OriginalString)

                    sChar = Mid(OriginalString, I, 1)

                    If Asc(sChar) > 0 Then

                        If UCase(sChar) <= "Z" And UCase(sChar) >= "A" Then

                            sOne = UCase(sChar)

                        Else

                            sOne = sChar

                        End If

                    ElseIf Asc(sChar) >= Asc("啊") And Asc(sChar) < Asc("八") Then

                        sOne = "A"

                    ElseIf Asc(sChar) >= Asc("八") And Asc(sChar) < Asc("擦") Then

                        sOne = "B"

                    ElseIf Asc(sChar) >= Asc("擦") And Asc(sChar) < Asc("搭") Then

                        sOne = "C"

                    ElseIf Asc(sChar) >= Asc("搭") And Asc(sChar) < Asc("蛾") Then

                        sOne = "D"

                    ElseIf Asc(sChar) >= Asc("蛾") And Asc(sChar) < Asc("发") Then

                        sOne = "E"

                    ElseIf Asc(sChar) >= Asc("发") And Asc(sChar) < Asc("噶") Then

                        sOne = "F"

                    ElseIf Asc(sChar) >= Asc("噶") And Asc(sChar) < Asc("哈") Then

                        sOne = "G"

                    ElseIf Asc(sChar) >= Asc("哈") And Asc(sChar) < Asc("击") Then

                        sOne = "H"

                    ElseIf Asc(sChar) >= Asc("击") And Asc(sChar) < Asc("喀") Then

                        sOne = "J"

                    ElseIf Asc(sChar) >= Asc("喀") And Asc(sChar) < Asc("垃") Then

                        sOne = "K"

                    ElseIf Asc(sChar) >= Asc("垃") And Asc(sChar) < Asc("妈") Then

                        sOne = "L"

                    ElseIf Asc(sChar) >= Asc("妈") And Asc(sChar) < Asc("拿") Then

                        sOne = "M"

                    ElseIf Asc(sChar) >= Asc("拿") And Asc(sChar) < Asc("哦") Then

                        sOne = "N"

                    ElseIf Asc(sChar) >= Asc("哦") And Asc(sChar) < Asc("啪") Then

                        sOne = "O"

                    ElseIf Asc(sChar) >= Asc("啪") And Asc(sChar) < Asc("期") Then

                        sOne = "P"

                    ElseIf Asc(sChar) >= Asc("期") And Asc(sChar) < Asc("然") Then

                        sOne = "Q"

                    ElseIf Asc(sChar) >= Asc("然") And Asc(sChar) < Asc("撒") Then

                        sOne = "R"

                    ElseIf Asc(sChar) >= Asc("撒") And Asc(sChar) < Asc("塌") Then

                        sOne = "S"

                    ElseIf Asc(sChar) >= Asc("塌") And Asc(sChar) < Asc("挖") Then

                        sOne = "T"

                    ElseIf Asc(sChar) >= Asc("挖") And Asc(sChar) < Asc("昔") Then

                        sOne = "W"

                    ElseIf Asc(sChar) >= Asc("昔") And Asc(sChar) < Asc("压") Then

                        sOne = "X"

                    ElseIf Asc(sChar) >= Asc("压") And Asc(sChar) < Asc("匝") Then

                        sOne = "Y"

                    ElseIf Asc(sChar) >= Asc("匝") And Asc(sChar) <= Asc("做") Then

                        sOne = "Z"

                    End If

                    sRet += sOne

                Next I

                If sRet.Length = 0 Or IsDBNull(sRet) Then
                    Return "AAA"
                Else
                    Return sRet
                End If

            Catch ex As Exception
                MsgBox(sRet & sChar & ex.ToString())
                Return ""
            End Try
        End Function

        ''' <summary>
        ''' 获取枚举类型的列表
        ''' </summary>
        ''' <param name="type">枚举类型</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetEnumList(ByVal type As Type) As IList(Of EnumList)
            Dim list As IList(Of EnumList) = New List(Of EnumList)()
            Dim str As String() = System.[Enum].GetNames(type)
            For Each s As String In str
                Dim en As New EnumList()
                en.ID = CInt(([Enum].Parse(type, s)))
                en.Name = s
                list.Add(en)
            Next
            Return list
        End Function

    End Class

    ''' <summary>
    ''' 枚举列表类
    ''' </summary>
    ''' <remarks></remarks>
    Public Class EnumList
        Private _ID As Integer = -999
        Private _Name As String = ""

        Public Property ID() As Integer
            Get
                Return _ID
            End Get
            Set(ByVal value As Integer)
                _ID = value
            End Set
        End Property
        Public Property Name() As String
            Get
                Return _Name
            End Get
            Set(ByVal value As String)
                _Name = value
            End Set
        End Property
    End Class



End Namespace

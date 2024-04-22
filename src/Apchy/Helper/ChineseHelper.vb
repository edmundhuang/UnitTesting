Imports System.Text
Imports Microsoft.International.Converters.PinYinConverter
Imports Microsoft.International.Formatters

Namespace Helper
    Public Class ChineseHelper
        Public Shared Function MixPYCode(ByVal ParamArray OriginalString() As String)
            If OriginalString.Length = 0 Then Return ""

            Dim I As Integer
            Dim ReturnString As String = ""
            Dim myString As String = ""

            For I = 0 To OriginalString.Length - 1
                myString = "" & OriginalString(I)

                If ReturnString.Length = 0 Then
                    ReturnString = myString
                Else
                    ReturnString = ReturnString & "|||" & myString
                End If
            Next

            ReturnString = GetPYCode(ReturnString)

            Return ReturnString
        End Function

        Public Shared Function GetPYCode(ByVal SourceString As String) As String
            Dim mySentense As New ChineseSentense(SourceString)

            Return mySentense.InitialString
        End Function

        Public Shared Function ConvertNumber(ByVal Number As Nullable(Of Decimal)) As String
            Dim ReturnString As String = Nothing

            If Number Is Nothing Then Return ReturnString

            ReturnString = String.Format(New EastAsiaNumericFormatter, "{0:L}", Number)

            Return ReturnString

            'East Asia Numeric Formatting Library - 支持将小写的数字字符串格式化成简体中文，繁体中文，日文和韩文的大写数字字符串。
            '【例如：】
            '--将数字转换为大写简体中文(拾贰亿叁仟肆佰伍拾陆万柒仟捌佰玖拾点肆伍)
            'string temp_4 = string.Format(new Microsoft.International.Formatters.EastAsiaNumericFormatter(), "{0:L}", 1234567890.45);
            '--将数字转换为小写(十二亿三千四百五十六万七千八百九十点四五)
            'string temp_6 = string.Format(new Microsoft.International.Formatters.EastAsiaNumericFormatter(), "{0:Ln}", 1234567890.45);
            '--将数字转换为货币(拾贰亿叁仟肆佰伍拾陆万柒仟捌佰玖拾点肆伍)
            'string temp_7 = string.Format(new Microsoft.International.Formatters.EastAsiaNumericFormatter(), "{0:Lc}", 1234567890.45);
        End Function

        'Public Shared Function GetPYCode(ByVal OriginalString As String) As String
        '    Dim sOne As String = ""
        '    Dim sChar As String = ""
        '    Dim sRet As String = ""
        '    Dim I As Integer

        '    Try
        '        For I = 1 To Len(OriginalString)
        '            sChar = Mid(OriginalString, I, 1)
        '            If Asc(sChar) > 0 Then
        '                If UCase(sChar) <= "Z" And UCase(sChar) >= "A" Then
        '                    sOne = UCase(sChar)
        '                Else
        '                    sOne = sChar
        '                End If
        '            ElseIf Asc(sChar) >= Asc("啊") And Asc(sChar) < Asc("八") Then
        '                sOne = "A"
        '            ElseIf Asc(sChar) >= Asc("八") And Asc(sChar) < Asc("擦") Then
        '                sOne = "B"
        '            ElseIf Asc(sChar) >= Asc("擦") And Asc(sChar) < Asc("搭") Then
        '                sOne = "C"
        '            ElseIf Asc(sChar) >= Asc("搭") And Asc(sChar) < Asc("蛾") Then
        '                sOne = "D"
        '            ElseIf Asc(sChar) >= Asc("蛾") And Asc(sChar) < Asc("发") Then
        '                sOne = "E"
        '            ElseIf Asc(sChar) >= Asc("发") And Asc(sChar) < Asc("噶") Then
        '                sOne = "F"
        '            ElseIf Asc(sChar) >= Asc("噶") And Asc(sChar) < Asc("哈") Then
        '                sOne = "G"
        '            ElseIf Asc(sChar) >= Asc("哈") And Asc(sChar) < Asc("击") Then
        '                sOne = "H"
        '            ElseIf Asc(sChar) >= Asc("击") And Asc(sChar) < Asc("喀") Then
        '                sOne = "J"
        '            ElseIf Asc(sChar) >= Asc("喀") And Asc(sChar) < Asc("垃") Then
        '                sOne = "K"
        '            ElseIf Asc(sChar) >= Asc("垃") And Asc(sChar) < Asc("妈") Then
        '                sOne = "L"
        '            ElseIf Asc(sChar) >= Asc("妈") And Asc(sChar) < Asc("拿") Then
        '                sOne = "M"
        '            ElseIf Asc(sChar) >= Asc("拿") And Asc(sChar) < Asc("哦") Then
        '                sOne = "N"
        '            ElseIf Asc(sChar) >= Asc("哦") And Asc(sChar) < Asc("啪") Then
        '                sOne = "O"
        '            ElseIf Asc(sChar) >= Asc("啪") And Asc(sChar) < Asc("期") Then
        '                sOne = "P"
        '            ElseIf Asc(sChar) >= Asc("期") And Asc(sChar) < Asc("然") Then
        '                sOne = "Q"
        '            ElseIf Asc(sChar) >= Asc("然") And Asc(sChar) < Asc("撒") Then
        '                sOne = "R"
        '            ElseIf Asc(sChar) >= Asc("撒") And Asc(sChar) < Asc("塌") Then
        '                sOne = "S"
        '            ElseIf Asc(sChar) >= Asc("塌") And Asc(sChar) < Asc("挖") Then
        '                sOne = "T"
        '            ElseIf Asc(sChar) >= Asc("挖") And Asc(sChar) < Asc("昔") Then
        '                sOne = "W"
        '            ElseIf Asc(sChar) >= Asc("昔") And Asc(sChar) < Asc("压") Then
        '                sOne = "X"
        '            ElseIf Asc(sChar) >= Asc("压") And Asc(sChar) < Asc("匝") Then
        '                sOne = "Y"
        '            ElseIf Asc(sChar) >= Asc("匝") And Asc(sChar) <= Asc("做") Then
        '                sOne = "Z"
        '            End If
        '            sRet += sOne
        '        Next I

        '        If sRet.Length = 0 Or IsDBNull(sRet) Then
        '            Return "AAA"
        '        Else
        '            Return sRet
        '        End If
        '    Catch ex As Exception
        '        MsgBox(sRet & sChar & ex.ToString())
        '        Return ""
        '    End Try
        'End Function

        'Public Shared Function GetAmountString(ByVal TotalAmount As Nullable(Of Decimal)) As String
        '    If TotalAmount Is Nothing Then
        '        Return ""
        '    End If

        '    Dim x As String, y As String
        '    Const zimu = ".sbqwsbqysbqwsbq" '定义位置代码
        '    Const letter = "0123456789sbqwy.zjf" '定义汉字缩写
        '    Const upcase = "零壹贰叁肆伍陆柒捌玖拾佰仟萬億圆整角分" '定义大写汉字
        '    Dim temp As String
        '    temp = TotalAmount

        '    If InStr(temp, ".") > 0 Then temp = Left(temp, InStr(temp, ".") - 1)
        '    If Len(temp) > 16 Then MsgBox("数目太大，无法换算！请输入一亿亿以下的数字", 64, "错误提示") : Return Nothing '只能转换一亿亿元以下数目的货币！
        '    x = Format(TotalAmount, "0.00") '格式化货币
        '    y = ""
        '    For i = 1 To Len(x) - 3
        '        y = y & Mid(x, i, 1) & Mid(zimu, Len(x) - 2 - i, 1)
        '    Next
        '    If Right(x, 3) = ".00" Then
        '        y = y & "z" '***元整
        '    Else
        '        y = y & Left(Right(x, 2), 1) & "j" & Right(x, 1) & "f" '*元*角*分
        '    End If
        '    y = Replace(y, "0q", "0") '避免零千(如：40200肆萬零千零贰佰)
        '    y = Replace(y, "0b", "0") '避免零百(如：41000肆萬壹千零佰)
        '    y = Replace(y, "0s", "0") '避免零十(如：204贰佰零拾零肆)
        '    Do While y <> Replace(y, "00", "0")
        '        y = Replace(y, "00", "0") '避免双零(如：1004壹仟零零肆)
        '    Loop
        '    y = Replace(y, "0y", "y") '避免零億(如：210億 贰佰壹十零億)
        '    y = Replace(y, "0w", "w") '避免零萬(如：210萬 贰佰壹十零萬)
        '    y = IIf(Len(x) = 5 And Left(y, 1) = "1", Right(y, Len(y) - 1), y) '避免壹十(如：14壹拾肆；10壹拾)
        '    y = IIf(Len(x) = 4, Replace(y, "0.", ""), Replace(y, "0.", ".")) '避免零元(如：20.00贰拾零圆；0.12零圆壹角贰分)
        '    For i = 1 To 19
        '        y = Replace(y, Mid(letter, i, 1), Mid(upcase, i, 1)) '大写汉字
        '    Next

        '    Return y
        'End Function
    End Class

    Public Class ChineseSentense
        Private m_SourceString As String
        Public Sub New()

        End Sub

        Public Sub New(ByVal SourceString As String)
            m_SourceString = SourceString
        End Sub

        Public Property Source As String
            Get
                Return m_SourceString
            End Get
            Set(value As String)
                If Not m_SourceString.Equals(value) Then
                    m_SourceString = value

                    m_PinYins = Nothing
                End If
            End Set
        End Property

        Private m_PinYins As List(Of String)()
        Public ReadOnly Property PinYins As List(Of String)()
            Get
                If m_PinYins IsNot Nothing Then
                    Return m_PinYins
                End If

                Dim iLength As Integer = Source.Length
                ReDim m_PinYins(iLength - 1)

                Dim I As Integer = 0
                For Each c In Source
                    Dim myList As New List(Of String)
                    If ChineseChar.IsValidChar(c) Then
                        Dim myChineseChar As New ChineseChar(c)

                        myList.AddRange(myChineseChar.Pinyins)
                    Else
                        myList.Add(c)
                    End If

                    m_PinYins(I) = myList
                    I += 1
                Next

                Return m_PinYins
            End Get
        End Property

        Private m_InitialString As String
        Public ReadOnly Property InitialString() As String
            Get
                If m_InitialString IsNot Nothing Then
                    Return m_InitialString
                End If

                Dim myPinYins() As List(Of String) = PinYins

                Dim mySB As New StringBuilder

                For I = 0 To myPinYins.Length - 1
                    Dim myObj As List(Of String) = myPinYins(I)

                    mySB.Append(myObj(0).Substring(0, 1))
                Next

                m_InitialString = mySB.ToString

                Return m_InitialString
            End Get
        End Property
    End Class

End Namespace
Imports System.Globalization
Namespace Helper

    Public Class DateTimeHelper
        Public Shared StartTickCount As Long
        Public Shared StartDateTime As DateTime

        Public Shared Sub RegisterServerTime(ByVal ServerTime As DateTime)
            StartTickCount = System.Environment.TickCount
            StartDateTime = ServerTime
        End Sub

        Public Shared Function ServerNow() As DateTime
            Dim iNum As Integer = System.Environment.TickCount
            iNum = iNum - StartTickCount

            Return StartDateTime.AddMilliseconds(iNum)
        End Function

        Public Shared Function ServerToday() As DateTime
            Return ServerNow.Date
        End Function

        Public Shared Function WeekNumber(ByVal SelectDate As DateTime,
                                          Optional ByVal SelectWeekRule As CalendarWeekRule = CalendarWeekRule.FirstDay,
                                          Optional ByVal SelectDayOfWeek As DayOfWeek = DayOfWeek.Monday) As Integer
            Dim culture As System.Globalization.CultureInfo = CultureInfo.CurrentCulture

            Dim intWeek As Integer = culture.Calendar.GetWeekOfYear(SelectDate, SelectWeekRule, SelectDayOfWeek)

            Return intWeek
        End Function

        Public Shared Function TimeSpan(ByVal StartTime As String, ByVal EndTime As String, ByVal SpanType As DateInterval) As Long

            Dim myDate As DateTime = Today
            Dim myStartTime As DateTime

            'dim culture as IFormatProvider= new CultureInfo("zh-CN", true);

            Try
                'DateTime.ParseExact(StartTime, "HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture)
                myStartTime = myDate & " " & StartTime
            Catch ex As Exception

            End Try

            Dim myEndTime As DateTime
            Try
                'DateTime.ParseExact(EndTime, "HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture)
                myEndTime = myDate & " " & EndTime

            Catch ex As Exception

            End Try

            '如果结束时间小于开始时间，结束时间加一天
            If DateTime.Compare(myStartTime, myEndTime) > 0 Then
                myEndTime = myEndTime.AddDays(1)
            End If

            Dim myTimeSpan As TimeSpan = myEndTime - myStartTime

            Select Case SpanType
                Case DateInterval.Hour
                    Return myTimeSpan.Hours
                Case DateInterval.Minute
                    Return myTimeSpan.Hours * 60 + myTimeSpan.Minutes
                Case DateInterval.Second
                    Return myTimeSpan.Hours * 3600 + myTimeSpan.Minutes * 60 + myTimeSpan.Seconds
            End Select
        End Function
    End Class

End Namespace
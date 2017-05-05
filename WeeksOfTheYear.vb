'Imports pbs.BO.EXT.VUS
Imports pbs.Helper
Imports System.Globalization


Namespace Data

    Public Class WeeksOfTheYear
        Inherits BO.Data.BaseUserTable

        Public Overrides ReadOnly Property DataStore As String
            Get
                'Return GetType(pbs.BO.EXT.VUS.).ToString
            End Get
        End Property

        Public Overrides ReadOnly Property Description As String
            Get
                Return "This user table return all week of the year"
            End Get
        End Property

        Public Overrides ReadOnly Property Group As String
            Get
                Return "EXT.VUS"
            End Get
        End Property

        Public Overrides ReadOnly Property Syntax As String
            Get
                Return <text>
                           Syntax: user table(pbs.BO.Data.WeeksOfTheYear?Year=...[&amp;WeekType=...])
                           WeekType - Optional as integer. 1 - get Consecutive weeks of the year. 2 - get Week of the year by month
                       </text>
            End Get
        End Property

        Public Overloads Overrides Function GetTable(pSource As String, pName As String) As DataTable
            Dim dt = GetTable(New pbsCmdArgs(pSource))
            dt.TableName = pName

            Return dt
        End Function

        Public Overrides Function GetTable(Args As pbsCmdArgs) As DataTable
            'get input parameters
            Dim year = Args.GetValueByKey("Year", Now.Year)
            Dim weekType = Args.GetValueByKey("WeekType", 1)
            Dim ret As New DataTable

            Select Case weekType
                Case 1
                    ret = ConsecutiveWeeksOfTheYear(year)
                Case 2
                    ret = WeeksOfTheYearByMonth(year)
            End Select

            Return ret
        End Function


        Private Function ConsecutiveWeeksOfTheYear(pYear As Integer) As DataTable
            'Create DataTable dt
            Dim dt As New DataTable
            dt.Columns.Add("Month")
            dt.Columns.Add("Week")
            dt.Columns.Add("WeekStart")
            dt.Columns.Add("WeekEnd")
            dt.Columns(0).AutoIncrement = True


            Dim dfi = DateTimeFormatInfo.CurrentInfo
            Dim calendar = dfi.Calendar
            Dim DayOfTheYear = New DateTime(pYear, 1, 1)
            Dim i = 1

            'Calculate
            While i < 366

                Dim week = calendar.GetWeekOfYear(DayOfTheYear.AddDays(i - 1), CalendarWeekRule.FirstDay, DayOfWeek.Monday)
                Dim month = DayOfTheYear.AddDays(i - 1).Month
                Dim weekStart = DayOfTheYear.AddDays(i - 1)
                Dim weekEnd As Date

                If weekStart.DayOfWeek = DayOfWeek.Sunday Then
                    weekEnd = weekStart

                    i += 1
                Else
                    weekEnd = weekStart.Add(New TimeSpan(7 - DayOfTheYear.AddDays(i - 1).DayOfWeek, 0, 0, 0))

                    i += 7 - DayOfTheYear.AddDays(i - 1).DayOfWeek + 1
                End If

                'Add to DataTable
                Dim newRow = dt.NewRow
                newRow("Week") = week
                newRow("Month") = month
                newRow("WeekStart") = weekStart.Date.ToString("dd/MM/yyyy")
                newRow("WeekEnd") = weekEnd.Date.ToString("dd/MM/yyyy")
                dt.Rows.Add(newRow)
            End While

            Return dt
        End Function

        Private Function WeeksOfTheYearByMonth(pYear As Integer) As DataTable
            'Create DataTable dt
            Dim dt As New DataTable
            dt.Columns.Add("Month")
            dt.Columns.Add("Week")
            dt.Columns.Add("WeekStart")
            dt.Columns.Add("WeekEnd")
            dt.Columns(0).AutoIncrement = True


            Dim dfi = DateTimeFormatInfo.CurrentInfo
            Dim calendar = dfi.Calendar
            Dim DayOfTheYear = New DateTime(pYear, 1, 1)
            Dim i = 1

            'Calculate
            While i < 366

                Dim week = calendar.GetWeekOfYear(DayOfTheYear.AddDays(i - 1), CalendarWeekRule.FirstDay, DayOfWeek.Monday)
                Dim month = DayOfTheYear.AddDays(i - 1).Month
                Dim weekStart = DayOfTheYear.AddDays(i - 1)
                Dim weekEnd As Date

                If weekStart.DayOfWeek = DayOfWeek.Sunday Then
                    weekEnd = weekStart

                    i += 1
                Else
                    weekEnd = weekStart.Add(New TimeSpan(7 - DayOfTheYear.AddDays(i - 1).DayOfWeek, 0, 0, 0))
                    Dim lastDayOfMonth = New DateTime(pYear, month, DateTime.DaysInMonth(pYear, month))

                    If weekEnd > lastDayOfMonth Then
                        weekEnd = lastDayOfMonth

                        i += DateDiff(DateInterval.Day, weekStart, lastDayOfMonth) + 1
                    Else

                        i += 7 - DayOfTheYear.AddDays(i - 1).DayOfWeek + 1
                    End If
                End If

                'Add to DataTable
                Dim newRow = dt.NewRow
                newRow("Week") = week
                newRow("Month") = month
                newRow("WeekStart") = weekStart.Date.ToString("dd/MM/yyyy")
                newRow("WeekEnd") = weekEnd.Date.ToString("dd/MM/yyyy")
                dt.Rows.Add(newRow)
            End While

            Return dt
        End Function


    End Class
End Namespace
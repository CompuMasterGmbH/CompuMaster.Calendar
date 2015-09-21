Imports NUnit.Framework

Namespace CompuMaster.Test.Calendar

    <TestFixture()> Public Class DateInformation

        <Test()> Public Sub LastDateOfWeekday()
            Dim Today0 As Date = Now
            Dim TodayM1 As Date = Now.AddDays(-1)
            Dim TodayM2 As Date = Now.AddDays(-2)
            Dim TodayM3 As Date = Now.AddDays(-3)
            Dim TodayM4 As Date = Now.AddDays(-4)
            Dim TodayM5 As Date = Now.AddDays(-5)
            Dim TodayM6 As Date = Now.AddDays(-6)
            Dim TodayM7 As Date = Now.AddDays(-7)
            Assert.AreEqual(Today0.Date, CompuMaster.Calendar.DateInformation.LastDateOfWeekday(Today0.DayOfWeek))
            Assert.AreEqual(TodayM1.Date, CompuMaster.Calendar.DateInformation.LastDateOfWeekday(TodayM1.DayOfWeek))
            Assert.AreEqual(TodayM2.Date, CompuMaster.Calendar.DateInformation.LastDateOfWeekday(TodayM2.DayOfWeek))
            Assert.AreEqual(TodayM3.Date, CompuMaster.Calendar.DateInformation.LastDateOfWeekday(TodayM3.DayOfWeek))
            Assert.AreEqual(TodayM4.Date, CompuMaster.Calendar.DateInformation.LastDateOfWeekday(TodayM4.DayOfWeek))
            Assert.AreEqual(TodayM5.Date, CompuMaster.Calendar.DateInformation.LastDateOfWeekday(TodayM5.DayOfWeek))
            Assert.AreEqual(TodayM6.Date, CompuMaster.Calendar.DateInformation.LastDateOfWeekday(TodayM6.DayOfWeek))
            Assert.AreEqual(Today0.Date, CompuMaster.Calendar.DateInformation.LastDateOfWeekday(TodayM7.DayOfWeek)) 'TodayM7 must result again in Today0
        End Sub

        <Test()> Public Sub LastDateOfWeek()

            Dim Jahr As Integer = 2005
            Dim TestCulture As System.Globalization.CultureInfo = System.Globalization.CultureInfo.CreateSpecificCulture("de-DE")
            For MyCounter As Integer = 1 To 52 '2005 contains 52 weeks
                Dim LastDate As Date
                Try
                    LastDate = CompuMaster.Calendar.DateInformation.LastDateOfWeek(Jahr, MyCounter, TestCulture)
                Catch ex As Exception
                    Throw New Exception("Error found for week " & MyCounter & "/" & Jahr, ex)
                End Try
                If LastDate.Year <> Jahr And lastdate.Year <> Jahr + 1 Then
                    Throw New Exception("Invalid date " & lastdate.ToString & " (requested year: " & Jahr & ")")
                End If
            Next

            Jahr = 2004
            For MyCounter As Integer = 1 To 53 '2004 contains 53 weeks
                Dim LastDate As Date
                Try
                    LastDate = CompuMaster.Calendar.DateInformation.LastDateOfWeek(Jahr, MyCounter, TestCulture)
                Catch ex As Exception
                    Throw New Exception("Error found for week " & MyCounter & "/" & Jahr, ex)
                End Try
                If LastDate.Year <> Jahr And lastdate.Year <> Jahr + 1 Then
                    Throw New Exception("Invalid date " & lastdate.ToString & " (requested year: " & Jahr & ")")
                End If
            Next

        End Sub

        <Test()> Public Sub FirstDateOfWeek()

            Dim Jahr As Integer = 2005
            Dim TestCulture As System.Globalization.CultureInfo = System.Globalization.CultureInfo.CreateSpecificCulture("de-DE")
            For MyCounter As Integer = 1 To 52 '2005 contains 52 weeks
                Dim LastDate As Date
                Try
                    LastDate = CompuMaster.Calendar.DateInformation.FirstDateOfWeek(Jahr, MyCounter, TestCulture)
                Catch ex As Exception
                    Throw New Exception("Error found for week " & MyCounter & "/" & Jahr, ex)
                End Try
                If LastDate.Year <> Jahr Then
                    Throw New Exception("Invalid date (another year)")
                End If
            Next

            Jahr = 2004
            For MyCounter As Integer = 1 To 53 '2004 contains 53 weeks
                Dim LastDate As Date
                Try
                    LastDate = CompuMaster.Calendar.DateInformation.FirstDateOfWeek(Jahr, MyCounter, TestCulture)
                Catch ex As Exception
                    Throw New Exception("Error found for week " & MyCounter & "/" & Jahr, ex)
                End Try
                If LastDate.Year <> Jahr Then
                    Throw New Exception("Invalid date (another year)")
                End If
            Next

        End Sub

        <Test()> Public Sub WeekOfYear()

            Dim Jahr As Integer = 2005
            Dim TestCulture As System.Globalization.CultureInfo = System.Globalization.CultureInfo.CreateSpecificCulture("de-DE")
            For MyCounter As Integer = 0 To 364 '365 days per year
                Dim DateOfYear As Date = New Date(Jahr, 1, 1).AddDays(MyCounter)
                Dim week As CompuMaster.Calendar.DateInformation.WeekNumber = CompuMaster.Calendar.DateInformation.WeekOfYear(DateOfYear, TestCulture)
                If week.Week < 1 Or week.Week > 53 Then
                    Throw New Exception("Invalid week number")
                End If
                If MyCounter < 15 And week.Week >= 15 Then
                    'Week year should be of 2004
                    Assert.AreEqual(Jahr - 1, week.Year)
                Else
                    Assert.AreEqual(Jahr, week.Year)
                End If
            Next

        End Sub

        <Test()> Public Sub CrossCheckWeekOfYear()

            Dim Jahr As Integer = 2005
            For MyCounter As Integer = 0 To 364 '365 days per year
                Dim DateOfYear As Date = New Date(Jahr, 1, 1).AddDays(MyCounter)
                Dim WeekNo As CompuMaster.Calendar.DateInformation.WeekNumber
                WeekNo = CompuMaster.Calendar.DateInformation.WeekOfYear(DateOfYear)
                Dim FirstDayOfWeek, LastDayOfWeek As Date
                FirstDayOfWeek = CompuMaster.Calendar.DateInformation.FirstDateOfWeek(WeekNo.Year, WeekNo.Week)
                LastDayOfWeek = CompuMaster.Calendar.DateInformation.LastDateOfWeek(WeekNo.Year, WeekNo.Week)
                If FirstDayOfWeek <= DateOfYear AndAlso DateOfYear <= LastDayOfWeek Then
                    'Absolutely fine
                Else
                    Assert.Fail("Week not in expected date range: " & DateOfYear.ToString, FirstDayOfWeek, LastDayOfWeek)
                End If
            Next

            Jahr = 2004
            For MyCounter As Integer = 0 To 365 '366 days per year
                Dim DateOfYear As Date = New Date(Jahr, 1, 1).AddDays(MyCounter)
                Dim WeekNo As CompuMaster.Calendar.DateInformation.WeekNumber
                WeekNo = CompuMaster.Calendar.DateInformation.WeekOfYear(DateOfYear)
                Dim FirstDayOfWeek, LastDayOfWeek As Date
                FirstDayOfWeek = CompuMaster.Calendar.DateInformation.FirstDateOfWeek(WeekNo.Year, WeekNo.Week)
                LastDayOfWeek = CompuMaster.Calendar.DateInformation.LastDateOfWeek(WeekNo.Year, WeekNo.Week)
                If FirstDayOfWeek <= DateOfYear AndAlso DateOfYear <= LastDayOfWeek Then
                    'Absolutely fine
                Else
                    Assert.Fail("Week not in expected date range: " & DateOfYear.ToString, FirstDayOfWeek, LastDayOfWeek)
                End If
            Next

        End Sub

        <Test()> Public Sub Min()
            Dim TimeSpanT1 As New TimeSpan(1&)
            Dim TimeSpanT0 As New TimeSpan(0&)
            Dim TimeSpanTM1 As New TimeSpan(-1&)
            Dim TimeSpanMS1 As New TimeSpan(0, 0, 0, 0, 1)
            Dim TimeSpanMS0 As New TimeSpan(0, 0, 0, 0, 0)
            Dim TimeSpanMSM1 As New TimeSpan(0, 0, 0, 0, -1)
            Dim TimeSpanSuperhigh As New TimeSpan(365 * 200, 0, 0, 0, 1)

            Assert.AreEqual(TimeSpanT0, CompuMaster.Calendar.DateInformation.Min(TimeSpanT0, TimeSpanT1))
            Assert.AreEqual(TimeSpanT0, CompuMaster.Calendar.DateInformation.Min(TimeSpanT1, TimeSpanT0))
            Assert.AreEqual(TimeSpanTM1, CompuMaster.Calendar.DateInformation.Min(TimeSpanTM1, TimeSpanT0))
            Assert.AreEqual(TimeSpanMS0, CompuMaster.Calendar.DateInformation.Min(TimeSpanMS1, TimeSpanMS0))
            Assert.AreEqual(TimeSpanMSM1, CompuMaster.Calendar.DateInformation.Min(TimeSpanMSM1, TimeSpanMS0))
            Assert.AreEqual(TimeSpanMS1, CompuMaster.Calendar.DateInformation.Min(TimeSpanMS1, TimeSpanSuperhigh))

        End Sub

        <Test()> Public Sub Max()
            Dim TimeSpanT1 As New TimeSpan(1&)
            Dim TimeSpanT0 As New TimeSpan(0&)
            Dim TimeSpanTM1 As New TimeSpan(-1&)
            Dim TimeSpanMS1 As New TimeSpan(0, 0, 0, 0, 1)
            Dim TimeSpanMS0 As New TimeSpan(0, 0, 0, 0, 0)
            Dim TimeSpanMSM1 As New TimeSpan(0, 0, 0, 0, -1)
            Dim TimeSpanSuperhigh As New TimeSpan(365 * 200, 0, 0, 0, 1)

            Assert.AreEqual(TimeSpanT1, CompuMaster.Calendar.DateInformation.Max(TimeSpanT0, TimeSpanT1))
            Assert.AreEqual(TimeSpanT1, CompuMaster.Calendar.DateInformation.Max(TimeSpanT1, TimeSpanT0))
            Assert.AreEqual(TimeSpanT0, CompuMaster.Calendar.DateInformation.Max(TimeSpanTM1, TimeSpanT0))
            Assert.AreEqual(TimeSpanMS1, CompuMaster.Calendar.DateInformation.Max(TimeSpanMS1, TimeSpanMS0))
            Assert.AreEqual(TimeSpanMS0, CompuMaster.Calendar.DateInformation.Max(TimeSpanMSM1, TimeSpanMS0))
            Assert.AreEqual(TimeSpanSuperhigh, CompuMaster.Calendar.DateInformation.Max(TimeSpanMS1, TimeSpanSuperhigh))

        End Sub

    End Class

End Namespace
Imports NUnit.Framework
Imports CompuMaster.Calendar

Namespace CompuMaster.Test.Calendar

    <TestFixture()> Public Class DateTimeRangeTest

        Private Shared Function LookupTestCulture(cultureName As String) As System.Globalization.CultureInfo
            Dim Result As System.Globalization.CultureInfo
            If cultureName <> Nothing Then
                Result = System.Globalization.CultureInfo.GetCultureInfo(cultureName)
            Else
                Result = System.Globalization.CultureInfo.InvariantCulture
            End If
            Return Result
        End Function

        <Test()> Public Sub ToStringTest(<Values("", "de-DE", "en-US", "en-UK", "fr-FR")> cultureName As String)
            Dim TestCulture As System.Globalization.CultureInfo = LookupTestCulture(cultureName)
            Dim TestRange As New DateTimeRange
            Assert.AreEqual("", TestRange.ToString)
            TestRange.From = New Date(2005, 1, 1)
            Select Case cultureName
                Case "de-DE"
                    Assert.AreEqual("01.01.2005 00:00:00 - *", TestRange.ToString(TestCulture))
                Case "en-US"
                    Assert.AreEqual("1/1/2005 12:00:00 AM - *", TestRange.ToString(TestCulture))
                Case "", Nothing
                    Assert.AreEqual(TestRange.From.ToString & " - *", TestRange.ToString)
                    Assert.AreEqual("2005-01-01 00:00:00 - *", TestRange.ToString(TestCulture))
                Case Else
                    Assert.AreEqual(TestRange.From.Value.ToString(TestCulture) & " - *", TestRange.ToString(TestCulture))
            End Select
            TestRange.Till = New Date(2005, 12, 31, 23, 59, 59, 999, DateTimeKind.Utc)
            Select Case cultureName
                Case "de-DE"
                    Assert.AreEqual("01.01.2005 00:00:00 - 31.12.2005 23:59:59", TestRange.ToString(TestCulture))
                Case "en-US"
                    Assert.AreEqual("1/1/2005 12:00:00 AM - 12/31/2005 11:59:59 PM", TestRange.ToString(TestCulture))
                Case "", Nothing
                    Assert.AreEqual(TestRange.From.ToString & " - " & TestRange.Till.ToString, TestRange.ToString)
                    Assert.AreEqual("2005-01-01 00:00:00 - 2005-12-31 23:59:59", TestRange.ToString(TestCulture))
                Case Else
                    Assert.AreEqual(TestRange.From.Value.ToString(TestCulture) & " - " & TestRange.Till.Value.ToString(TestCulture), TestRange.ToString(TestCulture))
            End Select
            TestRange.From = Nothing
            Select Case cultureName
                Case "de-DE"
                    Assert.AreEqual("* - 31.12.2005 23:59:59", TestRange.ToString(TestCulture))
                Case "en-US"
                    Assert.AreEqual("* - 12/31/2005 11:59:59 PM", TestRange.ToString(TestCulture))
                Case "", Nothing
                    Assert.AreEqual("* - " & TestRange.Till.ToString, TestRange.ToString)
                    Assert.AreEqual("* - 2005-12-31 23:59:59", TestRange.ToString(TestCulture))
                Case Else
                    Assert.AreEqual("* - " & TestRange.Till.Value.ToString(TestCulture), TestRange.ToString(TestCulture))
            End Select
            TestRange.Till = Nothing
            Assert.AreEqual("", TestRange.ToString)
        End Sub

        <Test()> Public Sub Parse(<Values("", "de-DE", "en-US", "en-UK", "fr-FR")> cultureName As String)
            Dim TestCulture As System.Globalization.CultureInfo = LookupTestCulture(cultureName)

            Dim ExpectedResult As New DateTimeRange()

            Dim Jahr As Integer = 2005
            For MyCounter As Integer = 1 To 52 '2005 contains 52 weeks
                Dim LastDate As Date
                Try
                    LastDate = DateInformation.LastDateOfWeek(Jahr, MyCounter, TestCulture)
                Catch ex As Exception
                    Throw New Exception("Error found for week " & MyCounter & "/" & Jahr, ex)
                End Try
                If LastDate.Year <> Jahr And LastDate.Year <> Jahr + 1 Then
                    Throw New Exception("Invalid date " & LastDate.ToString & " (requested year: " & Jahr & ")")
                End If
            Next

            Jahr = 2004
            For MyCounter As Integer = 1 To 53 '2004 contains 53 weeks
                Dim LastDate As Date
                Try
                    LastDate = DateInformation.LastDateOfWeek(Jahr, MyCounter, TestCulture)
                Catch ex As Exception
                    Throw New Exception("Error found for week " & MyCounter & "/" & Jahr, ex)
                End Try
                If LastDate.Year <> Jahr And LastDate.Year <> Jahr + 1 Then
                    Throw New Exception("Invalid date " & LastDate.ToString & " (requested year: " & Jahr & ")")
                End If
            Next

        End Sub

        <Test()> Public Sub EqualsTest()

            Dim Range1 As DateTimeRange
            Dim Range2 As DateTimeRange
            Dim ExpectedAsEqual As Boolean
            Dim CurrentDate As Date = Now

            Range1 = New DateTimeRange()
            Range2 = New DateTimeRange()
            ExpectedAsEqual = True
            Assert.AreEqual(ExpectedAsEqual, Range1 = Range2)
            Assert.AreEqual(Not ExpectedAsEqual, Range1 <> Range2)
            Assert.AreEqual(0, Range1.GetHashCode)
            Assert.AreEqual(0, Range2.GetHashCode)
            Assert.AreEqual(0, Range1.CompareTo(Range2))
            Assert.AreEqual(ExpectedAsEqual, Range1.Equals(Range2))
            Assert.AreEqual(ExpectedAsEqual, DateTimeRange.Equals(Range1, Range2))

            Range1 = New DateTimeRange()
            Range2 = New DateTimeRange(Nothing, Nothing)
            ExpectedAsEqual = True
            Assert.AreEqual(ExpectedAsEqual, DateTimeRange.Equals(Range1, Range2))
            Assert.AreEqual(ExpectedAsEqual, Range1.Equals(Range2))
            Assert.AreEqual(ExpectedAsEqual, Range1 = Range2)
            Assert.AreEqual(Not ExpectedAsEqual, Range1 <> Range2)

            Range1 = New DateTimeRange()
            Range2 = Nothing
            ExpectedAsEqual = True
            Assert.AreEqual(ExpectedAsEqual, DateTimeRange.Equals(Range1, Range2))
            Assert.AreEqual(ExpectedAsEqual, Range1.Equals(Range2))
            Assert.AreEqual(ExpectedAsEqual, Range1 = Range2)
            Assert.AreEqual(Not ExpectedAsEqual, Range1 <> Range2)

            Range1 = Nothing
            Range2 = Nothing
            ExpectedAsEqual = True
            Assert.AreEqual(ExpectedAsEqual, DateTimeRange.Equals(Range1, Range2))
            Assert.AreEqual(ExpectedAsEqual, Range1 = Range2)
            Assert.AreEqual(Not ExpectedAsEqual, Range1 <> Range2)

            Range1 = New DateTimeRange(DateTime.MinValue, DateTime.MaxValue)
            Range2 = New DateTimeRange()
            ExpectedAsEqual = True
            Assert.AreEqual(ExpectedAsEqual, DateTimeRange.Equals(Range1, Range2))
            Assert.AreEqual(ExpectedAsEqual, Range1.Equals(Range2))
            Assert.AreEqual(ExpectedAsEqual, Range1 = Range2)
            Assert.AreEqual(Not ExpectedAsEqual, Range1 <> Range2)

            Range1 = New DateTimeRange(DateTime.MinValue, Nothing)
            Range2 = New DateTimeRange()
            ExpectedAsEqual = True
            Assert.AreEqual(ExpectedAsEqual, DateTimeRange.Equals(Range1, Range2))
            Assert.AreEqual(ExpectedAsEqual, Range1.Equals(Range2))
            Assert.AreEqual(ExpectedAsEqual, Range1 = Range2)
            Assert.AreEqual(Not ExpectedAsEqual, Range1 <> Range2)

            Range1 = New DateTimeRange(DateTime.MinValue, CurrentDate)
            Range2 = New DateTimeRange()
            ExpectedAsEqual = False
            Assert.AreNotEqual(0, Range1.GetHashCode)
            Assert.AreEqual(0, Range2.GetHashCode)
            Assert.AreEqual(-1, Range1.CompareTo(Range2))
            Assert.AreEqual(ExpectedAsEqual, DateTimeRange.Equals(Range1, Range2))
            Assert.AreEqual(ExpectedAsEqual, Range1.Equals(Range2))
            Assert.AreEqual(ExpectedAsEqual, Range1 = Range2)
            Assert.AreEqual(Not ExpectedAsEqual, Range1 <> Range2)

            Range1 = New DateTimeRange(CurrentDate, DateTime.MaxValue)
            Range2 = New DateTimeRange()
            ExpectedAsEqual = False
            Assert.AreEqual(ExpectedAsEqual, DateTimeRange.Equals(Range1, Range2))
            Assert.AreEqual(ExpectedAsEqual, Range1.Equals(Range2))
            Assert.AreEqual(ExpectedAsEqual, Range1 = Range2)
            Assert.AreEqual(Not ExpectedAsEqual, Range1 <> Range2)

            Range1 = New DateTimeRange(DateTime.MinValue, CurrentDate)
            Range2 = New DateTimeRange(DateTime.MinValue, CurrentDate)
            ExpectedAsEqual = True
            Assert.AreNotEqual(0, Range1.GetHashCode)
            Assert.AreNotEqual(0, Range2.GetHashCode)
            Assert.AreEqual(0, Range1.CompareTo(Range2))
            Assert.AreEqual(ExpectedAsEqual, DateTimeRange.Equals(Range1, Range2))
            Assert.AreEqual(ExpectedAsEqual, Range1.Equals(Range2))
            Assert.AreEqual(ExpectedAsEqual, Range1 = Range2)
            Assert.AreEqual(Not ExpectedAsEqual, Range1 <> Range2)

        End Sub

    End Class

End Namespace
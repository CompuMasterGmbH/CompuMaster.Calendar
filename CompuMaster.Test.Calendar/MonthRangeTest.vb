Imports NUnit.Framework

Namespace CompuMaster.Test.Calendar

    <TestFixture()> Public Class MonthRangeTest

        <Test> Public Sub ValueExchange()
            Dim Value As CompuMaster.Calendar.MonthRange

            Value = New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 12))
            Assert.AreEqual(New CompuMaster.Calendar.Month(2020, 1), Value.FirstMonth)
            Assert.AreEqual(New CompuMaster.Calendar.Month(2020, 12), Value.LastMonth)

            Value = New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 12), New CompuMaster.Calendar.Month(2020, 1))
            Assert.AreEqual(New CompuMaster.Calendar.Month(2020, 1), Value.FirstMonth)
            Assert.AreEqual(New CompuMaster.Calendar.Month(2020, 12), Value.LastMonth)
        End Sub

        <Test> Public Sub ToStringTest()
            Dim Value As CompuMaster.Calendar.MonthRange
            Value = New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 12), New CompuMaster.Calendar.Month(2020, 1))
            Assert.AreEqual("2020-01 - 2020-12", Value.ToString)
        End Sub

        <Test> Public Sub OperatorsTest()
            Assert.IsTrue(CType(Nothing, CompuMaster.Calendar.MonthRange) = CType(Nothing, CompuMaster.Calendar.MonthRange))
            Assert.IsTrue(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 12), New CompuMaster.Calendar.Month(2020, 1)) <> CType(Nothing, CompuMaster.Calendar.MonthRange))
            Assert.IsTrue(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 12), New CompuMaster.Calendar.Month(2020, 1)) > CType(Nothing, CompuMaster.Calendar.MonthRange))
            Assert.IsTrue(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 12)) = New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 12)))
            Assert.IsTrue(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 12)) <= New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 12)))
            Assert.IsTrue(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 12)) >= New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 12)))
            Assert.IsTrue(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 11)) <> New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 12)))
            Assert.IsTrue(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 12)) < New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 2), New CompuMaster.Calendar.Month(2020, 12)))
            Assert.IsTrue(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 12)) <= New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 2), New CompuMaster.Calendar.Month(2020, 12)))
            Assert.IsTrue(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 11)) < New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 12)))
            Assert.IsTrue(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 11)) <= New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 12)))
            Assert.IsTrue(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 2), New CompuMaster.Calendar.Month(2020, 12)) > New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 12)))
            Assert.IsTrue(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 2), New CompuMaster.Calendar.Month(2020, 12)) >= New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 12)))
            Assert.IsTrue(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 12)) > New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 11)))
            Assert.IsTrue(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 12)) >= New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 11)))
        End Sub

        <Test> Public Sub Issue_6_ClassReferencesInsteadOfValueTypeCopies()
            Dim A As New CompuMaster.Calendar.Month(2020, 1)
            Dim B As New CompuMaster.Calendar.MonthRange(A, A)
            Assert.AreEqual(2020, B.FirstMonth.Year) 'success
            A.Year = 2021

            Assert.AreEqual(2020, B.FirstMonth.Year) 'failed: B.FirstYear = 2021 instead of expected 2020
            'expected by typical developer: 2020, but is now 2021 since Month is a class and 
            'MonthRange was created using a pointer/reference instead of a copy from A
        End Sub

    End Class

End Namespace

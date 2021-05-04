Imports NUnit.Framework

Namespace CompuMaster.Test.Calendar

    <TestFixture()> Public Class MonthRangeTest

        <Test> Public Sub ValueExchange()
            Dim Value As CompuMaster.Calendar.MonthRange

            Value = New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 12))
            Assert.AreEqual(New CompuMaster.Calendar.Month(2020, 1), Value.FirstPeriod)
            Assert.AreEqual(New CompuMaster.Calendar.Month(2020, 12), Value.LastPeriod)

            Value = New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 12), New CompuMaster.Calendar.Month(2020, 1))
            Assert.AreEqual(New CompuMaster.Calendar.Month(2020, 1), Value.FirstPeriod)
            Assert.AreEqual(New CompuMaster.Calendar.Month(2020, 12), Value.LastPeriod)
        End Sub

        <Test> Public Sub ToStringTest()
            Dim Value As CompuMaster.Calendar.MonthRange
            Value = New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 12), New CompuMaster.Calendar.Month(2020, 1))
            Assert.AreEqual("2020-01 - 2020-12", Value.ToString)
        End Sub

    End Class

End Namespace

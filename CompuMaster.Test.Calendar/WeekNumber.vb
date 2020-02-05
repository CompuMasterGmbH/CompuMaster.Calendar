Imports NUnit.Framework

Namespace CompuMaster.Test.Calendar

    <TestFixture()> Public Class WeekNumberTest

        <Test> Public Sub EqualsTest()
            Dim w0 As CompuMaster.Calendar.WeekNumber = Nothing
            Assert.AreEqual(Nothing, w0)
            Dim w0n As New CompuMaster.Calendar.WeekNumber
            Assert.AreNotEqual(Nothing, w0n)
            Assert.AreEqual(0, w0n.Year)
            Assert.AreEqual(0, w0n.Week)
            Dim w1 As New CompuMaster.Calendar.WeekNumber With {.Year = 2000, .Week = 14}
            Assert.AreNotEqual(w1, w0)
            Dim w1a As New CompuMaster.Calendar.WeekNumber With {.Year = 2000, .Week = 14}
            Assert.AreEqual(w1, w1a)
            Dim w2 As New CompuMaster.Calendar.WeekNumber With {.Year = 2000, .Week = 15}
            Assert.AreNotEqual(w1, w2)
            Assert.AreNotEqual(w0, w2)
            Dim w3 As New CompuMaster.Calendar.WeekNumber With {.Year = 2001, .Week = 14}
            Assert.AreNotEqual(w1, w3)
            Assert.AreNotEqual(w2, w3)
        End Sub

        <Test> Public Sub WeekNumberToString()
            Assert.AreEqual("2020/WK10", New CompuMaster.Calendar.WeekNumber(2020, 10).ToString)
            Assert.AreEqual("2020/WK01", New CompuMaster.Calendar.WeekNumber(2020, 1).ToString)
            Assert.AreEqual("0020/WK01", New CompuMaster.Calendar.WeekNumber(20, 1).ToString)
        End Sub
        <Test> Public Sub Conversions()
            Assert.AreEqual("2020/WK10", CType(New CompuMaster.Calendar.WeekNumber(2020, 10), String))
            Assert.AreEqual(Nothing, CType(CType(Nothing, CompuMaster.Calendar.WeekNumber), String))
            Assert.AreEqual(Nothing, CType(CType(Nothing, String), CompuMaster.Calendar.WeekNumber))
            Assert.AreEqual(Nothing, CType(String.Empty, CompuMaster.Calendar.WeekNumber))
            Assert.AreEqual("2020/WK10", CType("2020/WK10", CompuMaster.Calendar.WeekNumber).ToString)
        End Sub

    End Class

End Namespace
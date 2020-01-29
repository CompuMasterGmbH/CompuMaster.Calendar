Imports NUnit.Framework

Namespace CompuMaster.Test.Calendar

    <TestFixture()> Public Class WeekNumberTest

        <Test> Public Sub EqualsTest()
            Dim w0 As CompuMaster.Calendar.WeekNumber
            Assert.AreNotEqual(Nothing, w0)
            Assert.AreEqual(0, w0.Year)
            Assert.AreEqual(0, w0.Week)
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

    End Class

End Namespace
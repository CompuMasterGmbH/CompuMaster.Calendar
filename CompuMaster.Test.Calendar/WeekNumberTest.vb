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

        <Test> Public Sub FirstWeek()
            Assert.AreEqual("2010/WK01", New CompuMaster.Calendar.WeekNumber(2010, 0).FirstWeek.ToString)
            Assert.AreEqual("0001/WK01", New CompuMaster.Calendar.WeekNumber(1, 0).FirstWeek.ToString)
            Assert.AreEqual("0000/WK01", New CompuMaster.Calendar.WeekNumber(0, 0).FirstWeek.ToString)
        End Sub

        <Test> Public Sub NextPeriod()
            Assert.Throws(Of NotSupportedException)(Sub()
                                                        Dim Value As New CompuMaster.Calendar.WeekNumber(2010, 99)
                                                        Value.NextWeek.ToString()
                                                    End Sub)
            Assert.AreEqual("2010/WK01", New CompuMaster.Calendar.WeekNumber(2010, 0).NextWeek.ToString)
            Assert.AreEqual("2010/WK02", New CompuMaster.Calendar.WeekNumber(2010, 1).NextWeek.ToString)
            Assert.AreEqual("0001/WK01", New CompuMaster.Calendar.WeekNumber(1, 0).NextWeek.ToString)
            Assert.AreEqual("0000/WK01", New CompuMaster.Calendar.WeekNumber(0, 0).NextWeek.ToString)
        End Sub

        <Test> Public Sub PreviousPeriod()
            Assert.AreEqual("2010/WK11", New CompuMaster.Calendar.WeekNumber(2010, 12).PreviousWeek.ToString)
            Assert.Throws(Of NotSupportedException)(Sub()
                                                        Assert.AreEqual("2009/WK12", New CompuMaster.Calendar.WeekNumber(2010, 0).PreviousWeek.ToString)
                                                    End Sub)
            Assert.Throws(Of NotSupportedException)(Sub()
                                                        Assert.AreEqual("2009/WK12", New CompuMaster.Calendar.WeekNumber(2010, 1).PreviousWeek.ToString)
                                                    End Sub)
            Assert.Throws(Of NotSupportedException)(Sub()
                                                        Assert.AreEqual("0000/WK12", New CompuMaster.Calendar.WeekNumber(1, 0).PreviousWeek.ToString)
                                                    End Sub)
            Assert.Catch(Of NotSupportedException)(Sub()
                                                       Console.WriteLine(New CompuMaster.Calendar.WeekNumber(0, 0).PreviousWeek.ToString)
                                                   End Sub)
        End Sub

        <Test()> Public Sub OperatorMinus()
            Assert.Multiple(Sub()
                                Assert.AreEqual(New CompuMaster.Calendar.WeekNumberSpan(), New CompuMaster.Calendar.WeekNumber(2012, 5) - New CompuMaster.Calendar.WeekNumber(2012, 5))
                                Assert.AreEqual(New CompuMaster.Calendar.WeekNumberSpan(0, 1), New CompuMaster.Calendar.WeekNumber(2012, 5) - New CompuMaster.Calendar.WeekNumber(2012, 4))
                                Assert.AreEqual(New CompuMaster.Calendar.WeekNumberSpan(0, 2), New CompuMaster.Calendar.WeekNumber(2012, 5) - New CompuMaster.Calendar.WeekNumber(2012, 3))
                                Assert.AreEqual(New CompuMaster.Calendar.WeekNumberSpan(1, -11), New CompuMaster.Calendar.WeekNumber(2012, 1) - New CompuMaster.Calendar.WeekNumber(2011, 12))
                                Assert.AreEqual(New CompuMaster.Calendar.WeekNumberSpan(1, -9), New CompuMaster.Calendar.WeekNumber(2012, 2) - New CompuMaster.Calendar.WeekNumber(2011, 11))
                                Assert.AreEqual(New CompuMaster.Calendar.WeekNumberSpan(2, -9), New CompuMaster.Calendar.WeekNumber(2012, 2) - New CompuMaster.Calendar.WeekNumber(2010, 11))
                                Assert.AreEqual(New CompuMaster.Calendar.WeekNumberSpan(0, -1), New CompuMaster.Calendar.WeekNumber(2012, 4) - New CompuMaster.Calendar.WeekNumber(2012, 5))
                                Assert.AreEqual(New CompuMaster.Calendar.WeekNumberSpan(0, -2), New CompuMaster.Calendar.WeekNumber(2012, 3) - New CompuMaster.Calendar.WeekNumber(2012, 5))
                                Assert.AreEqual(New CompuMaster.Calendar.WeekNumberSpan(-1, 11), New CompuMaster.Calendar.WeekNumber(2011, 12) - New CompuMaster.Calendar.WeekNumber(2012, 1))
                                Assert.AreEqual(New CompuMaster.Calendar.WeekNumberSpan(-1, 9), New CompuMaster.Calendar.WeekNumber(2011, 11) - New CompuMaster.Calendar.WeekNumber(2012, 2))
                                Assert.AreEqual(New CompuMaster.Calendar.WeekNumberSpan(-2, 9), New CompuMaster.Calendar.WeekNumber(2010, 11) - New CompuMaster.Calendar.WeekNumber(2012, 2))
                                Assert.AreEqual(New CompuMaster.Calendar.WeekNumberSpan(1, 0), New CompuMaster.Calendar.WeekNumber(2020, 1) - New CompuMaster.Calendar.WeekNumber(2019, 1))
                                Assert.AreEqual(New CompuMaster.Calendar.WeekNumberSpan(1, 0), New CompuMaster.Calendar.WeekNumber(2020, 0) - New CompuMaster.Calendar.WeekNumber(2019, 0))
                                Assert.AreEqual(New CompuMaster.Calendar.WeekNumberSpan(1, 1), New CompuMaster.Calendar.WeekNumber(2020, 1) - New CompuMaster.Calendar.WeekNumber(2019, 0))
                                Assert.AreEqual(New CompuMaster.Calendar.WeekNumberSpan(1, 2), New CompuMaster.Calendar.WeekNumber(2020, 2) - New CompuMaster.Calendar.WeekNumber(2019, 0))
                                Assert.AreEqual(New CompuMaster.Calendar.WeekNumberSpan(1, -1), New CompuMaster.Calendar.WeekNumber(2020, 0) - New CompuMaster.Calendar.WeekNumber(2019, 1))
                                Assert.AreEqual(New CompuMaster.Calendar.WeekNumberSpan(0, 11), New CompuMaster.Calendar.WeekNumber(2019, 12) - New CompuMaster.Calendar.WeekNumber(2019, 1))
                            End Sub)
        End Sub

        <Test> Public Sub SmallerBiggerEquals()
            Assert.Multiple(Sub()
                                Assert.IsTrue(New CompuMaster.Calendar.WeekNumber(2010, 9) = New CompuMaster.Calendar.WeekNumber(2010, 9))
                                Assert.IsTrue(New CompuMaster.Calendar.WeekNumber(2010, 1) <> New CompuMaster.Calendar.WeekNumber(2010, 9))
                                Assert.IsTrue(New CompuMaster.Calendar.WeekNumber(2010, 1) < New CompuMaster.Calendar.WeekNumber(2010, 9))
                                Assert.IsTrue(New CompuMaster.Calendar.WeekNumber(2010, 9) > New CompuMaster.Calendar.WeekNumber(2010, 1))
                                Assert.IsFalse(New CompuMaster.Calendar.WeekNumber(2010, 9) <> New CompuMaster.Calendar.WeekNumber(2010, 9))
                                Assert.IsFalse(New CompuMaster.Calendar.WeekNumber(2010, 1) = New CompuMaster.Calendar.WeekNumber(2010, 9))
                                Assert.IsFalse(New CompuMaster.Calendar.WeekNumber(2010, 1) > New CompuMaster.Calendar.WeekNumber(2010, 9))
                                Assert.IsFalse(New CompuMaster.Calendar.WeekNumber(2010, 9) < New CompuMaster.Calendar.WeekNumber(2010, 1))

                                Assert.AreEqual(New CompuMaster.Calendar.WeekNumber(2010, 9), New CompuMaster.Calendar.WeekNumber(2010, 9))
                                Assert.AreNotEqual(New CompuMaster.Calendar.WeekNumber(2010, 1), New CompuMaster.Calendar.WeekNumber(2010, 9))
                                Assert.Less(New CompuMaster.Calendar.WeekNumber(2010, 1), New CompuMaster.Calendar.WeekNumber(2010, 9))
                                Assert.Greater(New CompuMaster.Calendar.WeekNumber(2010, 9), New CompuMaster.Calendar.WeekNumber(2010, 1))

                                Assert.IsTrue(New CompuMaster.Calendar.WeekNumber(2010, 0) = New CompuMaster.Calendar.WeekNumber(2010, 0))
                                Assert.IsTrue(New CompuMaster.Calendar.WeekNumber(2010, 0) <> New CompuMaster.Calendar.WeekNumber(2010, 1))
                                Assert.IsTrue(New CompuMaster.Calendar.WeekNumber(2010, 0) < New CompuMaster.Calendar.WeekNumber(2010, 1))
                                Assert.IsTrue(New CompuMaster.Calendar.WeekNumber(2010, 0) > New CompuMaster.Calendar.WeekNumber(2009, 12))
                                Assert.IsFalse(New CompuMaster.Calendar.WeekNumber(2010, 0) <> New CompuMaster.Calendar.WeekNumber(2010, 0))
                                Assert.IsFalse(New CompuMaster.Calendar.WeekNumber(2010, 0) = New CompuMaster.Calendar.WeekNumber(2010, 1))
                                Assert.IsFalse(New CompuMaster.Calendar.WeekNumber(2010, 0) > New CompuMaster.Calendar.WeekNumber(2010, 1))
                                Assert.IsFalse(New CompuMaster.Calendar.WeekNumber(2010, 0) < New CompuMaster.Calendar.WeekNumber(2009, 12))
                            End Sub)
        End Sub

        <Test> Public Sub Sorting()
            Dim Values As New Generic.List(Of CompuMaster.Calendar.WeekNumber)
            Values.Add(New CompuMaster.Calendar.WeekNumber(2010, 51))
            Values.Add(New CompuMaster.Calendar.WeekNumber(1999, 33))
            Values.Add(New CompuMaster.Calendar.WeekNumber(2010, 0))
            Values.Add(New CompuMaster.Calendar.WeekNumber(2010, 20))
            Values.Add(New CompuMaster.Calendar.WeekNumber)
            Values.Add(New CompuMaster.Calendar.WeekNumber(2010, 9))

            Values.Sort()

            Assert.AreEqual(New CompuMaster.Calendar.WeekNumber, Values(0))
            Assert.AreEqual(New CompuMaster.Calendar.WeekNumber(1999, 33), Values(1))
            Assert.AreEqual(New CompuMaster.Calendar.WeekNumber(2010, 0), Values(2))
            Assert.AreEqual(New CompuMaster.Calendar.WeekNumber(2010, 9), Values(3))
            Assert.AreEqual(New CompuMaster.Calendar.WeekNumber(2010, 20), Values(4))
            Assert.AreEqual(New CompuMaster.Calendar.WeekNumber(2010, 51), Values(5))
        End Sub

    End Class

End Namespace
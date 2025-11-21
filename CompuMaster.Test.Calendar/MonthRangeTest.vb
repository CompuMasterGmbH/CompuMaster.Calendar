Option Explicit On
Option Strict Off

Imports NUnit.Framework

Namespace CompuMaster.Test.Calendar

    <TestFixture()> Public Class MonthRangeTest

        <Test> Public Sub ValueAssignments()
            Dim Value As CompuMaster.Calendar.MonthRange

            Value = New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 12))
            Assert.AreEqual(New CompuMaster.Calendar.Month(2020, 1), Value.FirstMonth)
            Assert.AreEqual(New CompuMaster.Calendar.Month(2020, 12), Value.LastMonth)
            Assert.False(Value.IsEmpty)

            Assert.Catch(Of ArgumentException)(Sub()
                                                   Value = New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 12), New CompuMaster.Calendar.Month(2020, 1))
                                               End Sub)
            Assert.Catch(Of ArgumentException)(Sub()
                                                   Value = New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 12), New CompuMaster.Calendar.Month)
                                               End Sub)
            Assert.Catch(Of ArgumentNullException)(Sub()
                                                       Value = New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 12), CType(Nothing, CompuMaster.Calendar.Month))
                                                   End Sub)
            Assert.Catch(Of ArgumentNullException)(Sub()
                                                       Value = New CompuMaster.Calendar.MonthRange(CType(Nothing, CompuMaster.Calendar.Month), New CompuMaster.Calendar.Month(2020, 12))
                                                   End Sub)
        End Sub

        <Test> Public Sub ToStringTest()
            Dim Value As CompuMaster.Calendar.MonthRange
            Value = New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 12))
            Assert.AreEqual("2020-01 - 2020-12", Value.ToString)

            Value = CompuMaster.Calendar.MonthRange.Empty
            Assert.AreEqual("", Value.ToString)
        End Sub

        <Test> Public Sub OperatorsTest()
            Assert.IsTrue(CType(Nothing, CompuMaster.Calendar.MonthRange) = CType(Nothing, CompuMaster.Calendar.MonthRange))
            Assert.IsTrue(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 12)) <> CType(Nothing, CompuMaster.Calendar.MonthRange))
            Assert.IsTrue(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 12)) > CType(Nothing, CompuMaster.Calendar.MonthRange))
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

        <Test> Public Sub MonthCount()
            Dim Value As CompuMaster.Calendar.MonthRange

            Value = New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 12), New CompuMaster.Calendar.Month(2020, 12))
            Assert.AreEqual(13, Value.MonthCount)

            Value = New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 1))
            Assert.AreEqual(1, Value.MonthCount)

            Value = New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month, New CompuMaster.Calendar.Month)
            Assert.AreEqual("0001-01", Value.FirstMonth.ToString)
            Assert.AreEqual(1, Value.MonthCount)

            Value = CompuMaster.Calendar.MonthRange.Empty
            Assert.AreEqual(0, Value.MonthCount)

            Value = New CompuMaster.Calendar.MonthRange()
            Assert.AreEqual(0, Value.MonthCount)
        End Sub

        <Test> Public Sub Months()
            Dim Value As CompuMaster.Calendar.MonthRange
            Dim Expected As CompuMaster.Calendar.Month()

            Value = New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 12), New CompuMaster.Calendar.Month(2020, 3))
            Expected = New CompuMaster.Calendar.Month() {New CompuMaster.Calendar.Month(2019, 12), New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 2), New CompuMaster.Calendar.Month(2020, 3)}
            Assert.AreEqual(Expected, Value.Months)

            Value = New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 1))
            Expected = New CompuMaster.Calendar.Month() {New CompuMaster.Calendar.Month(2020, 1)}
            Assert.AreEqual(Expected, Value.Months)

            Value = New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month, New CompuMaster.Calendar.Month)
            Expected = New CompuMaster.Calendar.Month() {New CompuMaster.Calendar.Month(1, 1)}
            Assert.AreEqual(Expected, Value.Months)

            Value = CompuMaster.Calendar.MonthRange.Empty
            Expected = New CompuMaster.Calendar.Month() {}
            Assert.AreEqual(Expected, Value.Months)

            Value = New CompuMaster.Calendar.MonthRange()
            Expected = New CompuMaster.Calendar.Month() {}
            Assert.AreEqual(Expected, Value.Months)
        End Sub

        <Test> Public Sub Contains_Month()
            Dim Value As CompuMaster.Calendar.MonthRange

            Value = New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 12), New CompuMaster.Calendar.Month(2020, 3))
            Assert.AreEqual(False, Value.Contains(CType(Nothing, CompuMaster.Calendar.Month)))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.Month()))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.Month(2019, 11)))
            Assert.AreEqual(True, Value.Contains(New CompuMaster.Calendar.Month(2019, 12)))
            Assert.AreEqual(True, Value.Contains(New CompuMaster.Calendar.Month(2020, 1)))
            Assert.AreEqual(True, Value.Contains(New CompuMaster.Calendar.Month(2020, 2)))
            Assert.AreEqual(True, Value.Contains(New CompuMaster.Calendar.Month(2020, 3)))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.Month(2020, 4)))

            Value = New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 1))
            Assert.AreEqual(False, Value.Contains(CType(Nothing, CompuMaster.Calendar.Month)))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.Month()))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.Month(2019, 11)))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.Month(2019, 12)))
            Assert.AreEqual(True, Value.Contains(New CompuMaster.Calendar.Month(2020, 1)))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.Month(2020, 2)))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.Month(2020, 3)))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.Month(2020, 4)))

            Value = New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month, New CompuMaster.Calendar.Month)
            Assert.AreEqual(False, Value.Contains(CType(Nothing, CompuMaster.Calendar.Month)))
            Assert.AreEqual(True, Value.Contains(New CompuMaster.Calendar.Month()))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.Month(2019, 11)))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.Month(2019, 12)))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.Month(2020, 1)))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.Month(2020, 2)))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.Month(2020, 3)))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.Month(2020, 4)))

            Value = CompuMaster.Calendar.MonthRange.Empty
            Assert.AreEqual(False, Value.Contains(CType(Nothing, CompuMaster.Calendar.Month)))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.Month()))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.Month(2019, 11)))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.Month(2019, 12)))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.Month(2020, 1)))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.Month(2020, 2)))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.Month(2020, 3)))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.Month(2020, 4)))

        End Sub

        <Test> Public Sub Contains_MonthRange()
            Dim Value As CompuMaster.Calendar.MonthRange

            Value = New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 12), New CompuMaster.Calendar.Month(2020, 3))
            Assert.Throws(Of ArgumentNullException)(Sub()
                                                        Value.Contains(CType(Nothing, CompuMaster.Calendar.MonthRange))
                                                    End Sub)

            Value = New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 12), New CompuMaster.Calendar.Month(2020, 3))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 11), New CompuMaster.Calendar.Month(2020, 1))))
            Assert.AreEqual(True, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 12), New CompuMaster.Calendar.Month(2020, 1))))
            Assert.AreEqual(True, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 1))))
            Assert.AreEqual(True, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 3))))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 4))))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 11), New CompuMaster.Calendar.Month(2020, 4))))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 1), New CompuMaster.Calendar.Month(2019, 11))))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 4), New CompuMaster.Calendar.Month(2020, 12))))
            Assert.AreEqual(False, Value.Contains(CompuMaster.Calendar.MonthRange.Empty))

            Value = New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 1))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 11), New CompuMaster.Calendar.Month(2020, 1))))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 12), New CompuMaster.Calendar.Month(2020, 1))))
            Assert.AreEqual(True, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 1))))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 3))))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 4))))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 11), New CompuMaster.Calendar.Month(2020, 4))))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 1), New CompuMaster.Calendar.Month(2019, 11))))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 4), New CompuMaster.Calendar.Month(2020, 12))))
            Assert.AreEqual(False, Value.Contains(CompuMaster.Calendar.MonthRange.Empty))

            Value = New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month, New CompuMaster.Calendar.Month)
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 11), New CompuMaster.Calendar.Month(2020, 1))))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 12), New CompuMaster.Calendar.Month(2020, 1))))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 1))))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 3))))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 4))))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 11), New CompuMaster.Calendar.Month(2020, 4))))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 1), New CompuMaster.Calendar.Month(2019, 11))))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 4), New CompuMaster.Calendar.Month(2020, 12))))
            Assert.AreEqual(False, Value.Contains(CompuMaster.Calendar.MonthRange.Empty))

            Value = CompuMaster.Calendar.MonthRange.Empty
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 11), New CompuMaster.Calendar.Month(2020, 1))))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 12), New CompuMaster.Calendar.Month(2020, 1))))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 1))))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 3))))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 4))))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 11), New CompuMaster.Calendar.Month(2020, 4))))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 1), New CompuMaster.Calendar.Month(2019, 11))))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 4), New CompuMaster.Calendar.Month(2020, 12))))
            Assert.AreEqual(False, Value.Contains(CompuMaster.Calendar.MonthRange.Empty))
        End Sub

        <Test> Public Sub Overlaps()
            Dim Value As CompuMaster.Calendar.MonthRange

            Value = New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 12), New CompuMaster.Calendar.Month(2020, 3))
            Assert.Throws(Of ArgumentNullException)(Sub()
                                                        Value.Overlaps(CType(Nothing, CompuMaster.Calendar.MonthRange))
                                                    End Sub)

            Value = New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 12), New CompuMaster.Calendar.Month(2020, 3))
            Assert.AreEqual(True, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 11), New CompuMaster.Calendar.Month(2020, 1))))
            Assert.AreEqual(True, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 12), New CompuMaster.Calendar.Month(2020, 1))))
            Assert.AreEqual(True, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 1))))
            Assert.AreEqual(True, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 3))))
            Assert.AreEqual(True, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 4))))
            Assert.AreEqual(True, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 11), New CompuMaster.Calendar.Month(2020, 4))))
            Assert.AreEqual(False, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 1), New CompuMaster.Calendar.Month(2019, 11))))
            Assert.AreEqual(False, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 4), New CompuMaster.Calendar.Month(2020, 12))))
            Assert.AreEqual(False, Value.Overlaps(CompuMaster.Calendar.MonthRange.Empty))

            Value = New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 1))
            Assert.AreEqual(True, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 11), New CompuMaster.Calendar.Month(2020, 1))))
            Assert.AreEqual(True, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 12), New CompuMaster.Calendar.Month(2020, 1))))
            Assert.AreEqual(True, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 1))))
            Assert.AreEqual(True, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 3))))
            Assert.AreEqual(True, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 4))))
            Assert.AreEqual(True, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 11), New CompuMaster.Calendar.Month(2020, 4))))
            Assert.AreEqual(False, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 1), New CompuMaster.Calendar.Month(2019, 11))))
            Assert.AreEqual(False, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 4), New CompuMaster.Calendar.Month(2020, 12))))
            Assert.AreEqual(False, Value.Overlaps(CompuMaster.Calendar.MonthRange.Empty))

            Value = New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month, New CompuMaster.Calendar.Month)
            Assert.AreEqual(False, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 11), New CompuMaster.Calendar.Month(2020, 1))))
            Assert.AreEqual(False, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 12), New CompuMaster.Calendar.Month(2020, 1))))
            Assert.AreEqual(False, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 1))))
            Assert.AreEqual(False, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 3))))
            Assert.AreEqual(False, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 4))))
            Assert.AreEqual(False, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 11), New CompuMaster.Calendar.Month(2020, 4))))
            Assert.AreEqual(False, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 1), New CompuMaster.Calendar.Month(2019, 11))))
            Assert.AreEqual(False, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 4), New CompuMaster.Calendar.Month(2020, 12))))
            Assert.AreEqual(False, Value.Overlaps(CompuMaster.Calendar.MonthRange.Empty))

            Value = CompuMaster.Calendar.MonthRange.Empty
            Assert.AreEqual(False, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 11), New CompuMaster.Calendar.Month(2020, 1))))
            Assert.AreEqual(False, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 12), New CompuMaster.Calendar.Month(2020, 1))))
            Assert.AreEqual(False, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 1))))
            Assert.AreEqual(False, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 3))))
            Assert.AreEqual(False, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 4))))
            Assert.AreEqual(False, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 11), New CompuMaster.Calendar.Month(2020, 4))))
            Assert.AreEqual(False, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 1), New CompuMaster.Calendar.Month(2019, 11))))
            Assert.AreEqual(False, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 4), New CompuMaster.Calendar.Month(2020, 12))))
            Assert.AreEqual(False, Value.Overlaps(CompuMaster.Calendar.MonthRange.Empty))
        End Sub

        <Test> Public Sub Parse()
            Assert.True(CompuMaster.Calendar.MonthRange.Parse(Nothing).IsEmpty)
            Assert.True(CompuMaster.Calendar.MonthRange.Parse("").IsEmpty)
            Assert.Catch(Of FormatException)(Sub()
                                                 CompuMaster.Calendar.MonthRange.Parse("2020-04 2020-05")
                                             End Sub)

            Parse_AssertToStringAndReParseResultIsEqual(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 11), New CompuMaster.Calendar.Month(2020, 1)))
            Parse_AssertToStringAndReParseResultIsEqual(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2024, 1), New CompuMaster.Calendar.Month(2024, 12)))
            Parse_AssertToStringAndReParseResultIsEqual(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 11), New CompuMaster.Calendar.Month(2019, 11)))

            Parse_AssertToStringAndReParseResultIsEqual(CompuMaster.Calendar.MonthRange.Parse("2020-04 - 2020-05"))
            Parse_AssertToStringAndReParseResultIsEqual(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2024, 1), New CompuMaster.Calendar.Month(2024, 12)))

            Parse_AssertToStringAndReParseResultIsEqual(CompuMaster.Calendar.MonthRange.Parse(""))
        End Sub

        Private Sub Parse_AssertToStringAndReParseResultIsEqual(value As CompuMaster.Calendar.MonthRange)
            Dim TextRepresentation As String = value.ToString
            If value.IsEmpty Then
                Assert.AreEqual("", TextRepresentation)
            Else
                Assert.AreEqual(value.FirstMonth.ToString & " - " & value.LastMonth.ToString, TextRepresentation)
            End If
            Dim Reparsed As CompuMaster.Calendar.MonthRange = CompuMaster.Calendar.MonthRange.Parse(TextRepresentation)
            Assert.AreEqual(value, Reparsed)
        End Sub

        <Test> Public Sub TryParse()
            Dim Buffer As CompuMaster.Calendar.MonthRange
            Assert.True(CompuMaster.Calendar.MonthRange.TryParse(Nothing, Buffer))
            Assert.True(Buffer.IsEmpty)
            Assert.True(CompuMaster.Calendar.MonthRange.TryParse("", Buffer))
            Assert.True(Buffer.IsEmpty)
            Assert.False(CompuMaster.Calendar.MonthRange.TryParse("2020-04 2020-05", Buffer))
            Assert.That(Buffer, [Is].Null)
            Assert.True(CompuMaster.Calendar.MonthRange.TryParse("2020-04 - 2020-05", Buffer))
            Assert.That(Buffer, [Is].Not.Null)
            Assert.False(CompuMaster.Calendar.MonthRange.TryParse("2020-00 - 2020-05", Buffer))
            Assert.That(Buffer, [Is].Null)
            Assert.False(CompuMaster.Calendar.MonthRange.TryParse("A020-04 - 2020-05", Buffer))
            Assert.That(Buffer, [Is].Null)

            TryParse_AssertToStringAndReParseResultIsEqual(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 11), New CompuMaster.Calendar.Month(2020, 1)))
            TryParse_AssertToStringAndReParseResultIsEqual(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2024, 1), New CompuMaster.Calendar.Month(2024, 12)))
            TryParse_AssertToStringAndReParseResultIsEqual(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 11), New CompuMaster.Calendar.Month(2019, 11)))

            TryParse_AssertToStringAndReParseResultIsEqual(CompuMaster.Calendar.MonthRange.Parse("2020-04 - 2020-05"))
            TryParse_AssertToStringAndReParseResultIsEqual(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2024, 1), New CompuMaster.Calendar.Month(2024, 12)))

            TryParse_AssertToStringAndReParseResultIsEqual(CompuMaster.Calendar.MonthRange.Parse(""))
        End Sub

        Private Sub TryParse_AssertToStringAndReParseResultIsEqual(value As CompuMaster.Calendar.MonthRange)
            Dim TextRepresentation As String = value.ToString
            If value.IsEmpty Then
                Assert.AreEqual("", TextRepresentation)
            Else
                Assert.AreEqual(value.FirstMonth.ToString & " - " & value.LastMonth.ToString, TextRepresentation)
            End If
            Dim Reparsed As CompuMaster.Calendar.MonthRange = Nothing
            Assert.True(CompuMaster.Calendar.MonthRange.TryParse(TextRepresentation, Reparsed))
            Assert.AreEqual(value, Reparsed)
        End Sub

        <Test> Public Sub Add()
            Dim OriginRange As CompuMaster.Calendar.MonthRange = CompuMaster.Calendar.MonthRange.Parse("2000-01 - 2020-12")
            Assert.AreEqual("2000-01 - 2020-12", OriginRange.Add(0, 0).ToString)
            Assert.AreEqual("2010-01 - 2030-12", OriginRange.Add(10, 0).ToString)
            Assert.AreEqual("2001-01 - 2021-12", OriginRange.Add(0, 12).ToString)
            Assert.AreEqual("2005-02 - 2026-01", OriginRange.Add(4, 13).ToString)
            Assert.AreEqual("1994-12 - 2015-11", OriginRange.Add(-4, -13).ToString)
            Assert.Catch(Of InvalidOperationException)(Sub()
                                                           CompuMaster.Calendar.MonthRange.Empty.Add(0, 0)
                                                       End Sub)
        End Sub

        <Test> Public Sub AddMonths()
            Dim OriginRange As CompuMaster.Calendar.MonthRange = CompuMaster.Calendar.MonthRange.Parse("2000-01 - 2020-12")
            Assert.AreEqual("2000-01 - 2020-12", OriginRange.AddMonths(0).ToString)
            Assert.AreEqual("2000-05 - 2021-04", OriginRange.AddMonths(4).ToString)
            Assert.AreEqual("1999-09 - 2020-08", OriginRange.AddMonths(-4).ToString)
            Assert.Catch(Of InvalidOperationException)(Sub()
                                                           CompuMaster.Calendar.MonthRange.Empty.AddMonths(0)
                                                       End Sub)
        End Sub

        <Test> Public Sub AddYears()
            Dim OriginRange As CompuMaster.Calendar.MonthRange = CompuMaster.Calendar.MonthRange.Parse("2000-01 - 2020-12")
            Assert.AreEqual("2000-01 - 2020-12", OriginRange.AddYears(0).ToString)
            Assert.AreEqual("2004-01 - 2024-12", OriginRange.AddYears(4).ToString)
            Assert.AreEqual("1996-01 - 2016-12", OriginRange.AddYears(-4).ToString)
            Assert.Catch(Of InvalidOperationException)(Sub()
                                                           CompuMaster.Calendar.MonthRange.Empty.AddYears(0)
                                                       End Sub)
        End Sub

        <Test> Public Sub Empty()
            Dim Value As CompuMaster.Calendar.MonthRange

            Value = CompuMaster.Calendar.MonthRange.Empty
            Assert.IsTrue(Value.IsEmpty)
            Assert.IsNull(Value.FirstMonth)
            Assert.IsNull(Value.LastMonth)
            Assert.Zero(Value.MonthCount)
            Assert.AreEqual(New CompuMaster.Calendar.Month() {}, Value.Months)

            Value = New CompuMaster.Calendar.MonthRange()
            Assert.IsTrue(Value.IsEmpty)
            Assert.IsNull(Value.FirstMonth)
            Assert.IsNull(Value.LastMonth)
            Assert.Zero(Value.MonthCount)
            Assert.AreEqual(New CompuMaster.Calendar.Month() {}, Value.Months)
        End Sub

        <Test> Public Sub EmptyComparisons()
            Dim Value As CompuMaster.Calendar.MonthRange

            Value = CompuMaster.Calendar.MonthRange.Empty
            Assert.IsTrue(Value.Equals(New CompuMaster.Calendar.MonthRange()))
            Assert.IsTrue(Value = New CompuMaster.Calendar.MonthRange())
            Assert.AreEqual(0, Value.CompareTo(New CompuMaster.Calendar.MonthRange()))

            Assert.AreEqual(New CompuMaster.Calendar.MonthRange(), CompuMaster.Calendar.MonthRange.Empty)
        End Sub

        <Test> Public Sub SimplifiedName()
            Assert.AreEqual("", New CompuMaster.Calendar.MonthRange().SimplifiedName)
            Assert.AreEqual("2000", New CompuMaster.Calendar.MonthRange(2000, 1, 2000, 12).SimplifiedName)
            Assert.AreEqual("2000-01 - 2020-12", New CompuMaster.Calendar.MonthRange(2000, 1, 2020, 12).SimplifiedName)
            Assert.AreEqual("2000-02 - 2000-12", New CompuMaster.Calendar.MonthRange(2000, 2, 2000, 12).SimplifiedName)
            Assert.AreEqual("2000-01 - 2000-11", New CompuMaster.Calendar.MonthRange(2000, 1, 2000, 11).SimplifiedName)
        End Sub

    End Class

End Namespace

Imports NUnit.Framework

Namespace CompuMaster.Test.Calendar

    <TestFixture()> Public Class MonthRangeTest

        <Test> Public Sub ValueAssignments()
            Dim Value As CompuMaster.Calendar.MonthRange

            Value = New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 12))
            Assert.AreEqual(New CompuMaster.Calendar.Month(2020, 1), Value.FirstMonth)
            Assert.AreEqual(New CompuMaster.Calendar.Month(2020, 12), Value.LastMonth)

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

            Value = New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 1))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 11), New CompuMaster.Calendar.Month(2020, 1))))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 12), New CompuMaster.Calendar.Month(2020, 1))))
            Assert.AreEqual(True, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 1))))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 3))))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 4))))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 11), New CompuMaster.Calendar.Month(2020, 4))))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 1), New CompuMaster.Calendar.Month(2019, 11))))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 4), New CompuMaster.Calendar.Month(2020, 12))))

            Value = New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month, New CompuMaster.Calendar.Month)
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 11), New CompuMaster.Calendar.Month(2020, 1))))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 12), New CompuMaster.Calendar.Month(2020, 1))))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 1))))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 3))))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 4))))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 11), New CompuMaster.Calendar.Month(2020, 4))))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 1), New CompuMaster.Calendar.Month(2019, 11))))
            Assert.AreEqual(False, Value.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 4), New CompuMaster.Calendar.Month(2020, 12))))
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

            Value = New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 1))
            Assert.AreEqual(True, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 11), New CompuMaster.Calendar.Month(2020, 1))))
            Assert.AreEqual(True, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 12), New CompuMaster.Calendar.Month(2020, 1))))
            Assert.AreEqual(True, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 1))))
            Assert.AreEqual(True, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 3))))
            Assert.AreEqual(True, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 4))))
            Assert.AreEqual(True, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 11), New CompuMaster.Calendar.Month(2020, 4))))
            Assert.AreEqual(False, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 1), New CompuMaster.Calendar.Month(2019, 11))))
            Assert.AreEqual(False, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 4), New CompuMaster.Calendar.Month(2020, 12))))

            Value = New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month, New CompuMaster.Calendar.Month)
            Assert.AreEqual(False, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 11), New CompuMaster.Calendar.Month(2020, 1))))
            Assert.AreEqual(False, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 12), New CompuMaster.Calendar.Month(2020, 1))))
            Assert.AreEqual(False, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 1))))
            Assert.AreEqual(False, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 3))))
            Assert.AreEqual(False, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 1), New CompuMaster.Calendar.Month(2020, 4))))
            Assert.AreEqual(False, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 11), New CompuMaster.Calendar.Month(2020, 4))))
            Assert.AreEqual(False, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 1), New CompuMaster.Calendar.Month(2019, 11))))
            Assert.AreEqual(False, Value.Overlaps(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2020, 4), New CompuMaster.Calendar.Month(2020, 12))))
        End Sub

        <Test> Public Sub Parse()
            Assert.Catch(Of ArgumentNullException)(Sub()
                                                       CompuMaster.Calendar.MonthRange.Parse(Nothing)
                                                   End Sub)
            Assert.Catch(Of ArgumentNullException)(Sub()
                                                       CompuMaster.Calendar.MonthRange.Parse("")
                                                   End Sub)
            Assert.Catch(Of FormatException)(Sub()
                                                 CompuMaster.Calendar.MonthRange.Parse("2020-04 2020-05")
                                             End Sub)

            Parse_AssertToStringAndReParseResultIsEqual(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 11), New CompuMaster.Calendar.Month(2020, 1)))
            Parse_AssertToStringAndReParseResultIsEqual(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2024, 1), New CompuMaster.Calendar.Month(2024, 12)))
            Parse_AssertToStringAndReParseResultIsEqual(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2019, 11), New CompuMaster.Calendar.Month(2019, 11)))

            Parse_AssertToStringAndReParseResultIsEqual(CompuMaster.Calendar.MonthRange.Parse("2020-04 - 2020-05"))
            Parse_AssertToStringAndReParseResultIsEqual(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2024, 1), New CompuMaster.Calendar.Month(2024, 12)))

        End Sub

        Private Sub Parse_AssertToStringAndReParseResultIsEqual(value As CompuMaster.Calendar.MonthRange)
            Dim TextRepresentation As String = value.ToString
            Assert.AreEqual(value.FirstMonth.ToString & " - " & value.LastMonth.ToString, TextRepresentation)
            Dim Reparsed As CompuMaster.Calendar.MonthRange = CompuMaster.Calendar.MonthRange.Parse(TextRepresentation)
            Assert.AreEqual(value, Reparsed)
        End Sub

    End Class

End Namespace

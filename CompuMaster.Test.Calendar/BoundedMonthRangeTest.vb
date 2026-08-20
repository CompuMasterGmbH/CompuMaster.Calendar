Option Explicit On
Option Strict On

Imports NUnit.Framework

Namespace CompuMaster.Test.Calendar

    <TestFixture()> Public Class BoundedMonthRangeTest

        <Test> Public Sub OptionallyBoundedConstructionAndFormatting()
            Assert.AreEqual("* - *", New CompuMaster.Calendar.OptionallyBoundedMonthRange(Nothing, Nothing).ToString())
            Assert.AreEqual("* - 2026-12", New CompuMaster.Calendar.OptionallyBoundedMonthRange(Nothing, New CompuMaster.Calendar.Month(2026, 12)).ToString())
            Assert.AreEqual("2027-01 - *", New CompuMaster.Calendar.OptionallyBoundedMonthRange(New CompuMaster.Calendar.Month(2027, 1), Nothing).ToString())
            Assert.AreEqual("2027-01 - 2027-12", New CompuMaster.Calendar.OptionallyBoundedMonthRange(New CompuMaster.Calendar.Month(2027, 1), New CompuMaster.Calendar.Month(2027, 12)).ToString())
            Assert.Throws(Of ArgumentException)(Sub()
                                                    Dim InvalidValue As New CompuMaster.Calendar.OptionallyBoundedMonthRange(New CompuMaster.Calendar.Month(2027, 2), New CompuMaster.Calendar.Month(2027, 1))
                                                End Sub)
        End Sub

        <Test> Public Sub PartiallyBoundedConstruction()
            Assert.Throws(Of ArgumentException)(Sub()
                                                    Dim InvalidValue As New CompuMaster.Calendar.PartiallyBoundedMonthRange(Nothing, Nothing)
                                                End Sub)
            Assert.AreEqual("* - 2026-12", New CompuMaster.Calendar.PartiallyBoundedMonthRange(Nothing, New CompuMaster.Calendar.Month(2026, 12)).ToString())
            Assert.AreEqual("2027-01 - *", New CompuMaster.Calendar.PartiallyBoundedMonthRange(New CompuMaster.Calendar.Month(2027, 1), Nothing).ToString())
            Assert.AreEqual("2027-01 - 2027-12", New CompuMaster.Calendar.PartiallyBoundedMonthRange(New CompuMaster.Calendar.Month(2027, 1), New CompuMaster.Calendar.Month(2027, 12)).ToString())
        End Sub

        <Test> Public Sub ConstructorClonesBoundaries()
            Dim First As New CompuMaster.Calendar.Month(2027, 1)
            Dim Value As New CompuMaster.Calendar.OptionallyBoundedMonthRange(First, Nothing)
            First.Year = 2030
            Assert.AreEqual("2027-01 - *", Value.ToString())
        End Sub

        <Test> Public Sub Parsing()
            Assert.AreEqual("* - *", CompuMaster.Calendar.OptionallyBoundedMonthRange.Parse("* - *").ToString())
            Assert.AreEqual("* - 2026-12", CompuMaster.Calendar.PartiallyBoundedMonthRange.Parse("* - 2026-12").ToString())
            Assert.AreEqual("2027-01 - 2027-12", CompuMaster.Calendar.PartiallyBoundedMonthRange.Parse("2027-01 - 2027-12").ToString())
            Assert.Throws(Of ArgumentException)(Sub() CompuMaster.Calendar.PartiallyBoundedMonthRange.Parse("* - *"))
            Assert.Throws(Of FormatException)(Sub() CompuMaster.Calendar.OptionallyBoundedMonthRange.Parse(""))
            Assert.Throws(Of FormatException)(Sub() CompuMaster.Calendar.OptionallyBoundedMonthRange.Parse(Nothing))

            Dim OptionalResult As CompuMaster.Calendar.OptionallyBoundedMonthRange = Nothing
            Assert.False(CompuMaster.Calendar.OptionallyBoundedMonthRange.TryParse("", OptionalResult))
            Assert.IsNull(OptionalResult)
            Assert.False(CompuMaster.Calendar.OptionallyBoundedMonthRange.TryParse("2027-00 - *", OptionalResult))
            Assert.True(CompuMaster.Calendar.OptionallyBoundedMonthRange.TryParse("2027-01 - *", OptionalResult))

            Dim PartialResult As CompuMaster.Calendar.PartiallyBoundedMonthRange = Nothing
            Assert.False(CompuMaster.Calendar.PartiallyBoundedMonthRange.TryParse("* - *", PartialResult))
            Assert.IsNull(PartialResult)
        End Sub

        <Test> Public Sub ContainsMonthsAndRanges()
            Dim UntilDecember As New CompuMaster.Calendar.OptionallyBoundedMonthRange(Nothing, New CompuMaster.Calendar.Month(2026, 12))
            Assert.True(UntilDecember.Contains(New CompuMaster.Calendar.Month(1, 1)))
            Assert.True(UntilDecember.Contains(New CompuMaster.Calendar.Month(2026, 12)))
            Assert.False(UntilDecember.Contains(New CompuMaster.Calendar.Month(2027, 1)))
            Assert.False(UntilDecember.Contains(CType(Nothing, CompuMaster.Calendar.Month)))

            Dim AllMonths As New CompuMaster.Calendar.OptionallyBoundedMonthRange(Nothing, Nothing)
            Assert.True(AllMonths.Contains(UntilDecember))
            Assert.False(UntilDecember.Contains(AllMonths))
            Assert.True(UntilDecember.Contains(New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2026, 1), New CompuMaster.Calendar.Month(2026, 12))))
        End Sub

        <Test> Public Sub OverlapsRanges()
            Dim Historical As New CompuMaster.Calendar.OptionallyBoundedMonthRange(Nothing, New CompuMaster.Calendar.Month(2026, 12))
            Dim Current As New CompuMaster.Calendar.OptionallyBoundedMonthRange(New CompuMaster.Calendar.Month(2027, 1), Nothing)
            Dim Touching As New CompuMaster.Calendar.OptionallyBoundedMonthRange(New CompuMaster.Calendar.Month(2026, 12), Nothing)
            Assert.False(Historical.Overlaps(Current))
            Assert.True(Historical.Overlaps(Touching))
            Assert.True(New CompuMaster.Calendar.OptionallyBoundedMonthRange(Nothing, Nothing).Overlaps(Current))
        End Sub

        <Test> Public Sub Equality()
            Dim A As New CompuMaster.Calendar.OptionallyBoundedMonthRange(Nothing, New CompuMaster.Calendar.Month(2026, 12))
            Dim B As New CompuMaster.Calendar.OptionallyBoundedMonthRange(Nothing, New CompuMaster.Calendar.Month(2026, 12))
            Dim C As New CompuMaster.Calendar.OptionallyBoundedMonthRange(Nothing, New CompuMaster.Calendar.Month(2027, 1))
            Assert.True(A = B)
            Assert.False(A <> B)
            Assert.False(A = C)
            Assert.AreEqual(A.GetHashCode(), B.GetHashCode())
        End Sub

        <Test> Public Sub MonthRangeConversions()
            Dim Original As New CompuMaster.Calendar.MonthRange(New CompuMaster.Calendar.Month(2027, 1), New CompuMaster.Calendar.Month(2027, 12))
            Dim OptionalValue As CompuMaster.Calendar.OptionallyBoundedMonthRange = CType(Original, CompuMaster.Calendar.OptionallyBoundedMonthRange)
            Dim PartialValue As CompuMaster.Calendar.PartiallyBoundedMonthRange = CType(Original, CompuMaster.Calendar.PartiallyBoundedMonthRange)
            Assert.AreEqual(Original.ToString(), OptionalValue.ToString())
            Assert.AreEqual(Original.ToString(), PartialValue.ToString())
            Assert.AreEqual(Original, OptionalValue.ToMonthRange())
            Assert.AreEqual(Original, CType(PartialValue, CompuMaster.Calendar.MonthRange))

            Dim OpenValue As New CompuMaster.Calendar.OptionallyBoundedMonthRange(Nothing, New CompuMaster.Calendar.Month(2026, 12))
            Dim Converted As CompuMaster.Calendar.MonthRange = Nothing
            Assert.False(OpenValue.TryToMonthRange(Converted))
            Assert.IsNull(Converted)
            Assert.Throws(Of InvalidOperationException)(Sub() OpenValue.ToMonthRange())
            Assert.Throws(Of InvalidCastException)(Sub()
                                                       Dim InvalidValue As CompuMaster.Calendar.OptionallyBoundedMonthRange = CType(CompuMaster.Calendar.MonthRange.Empty, CompuMaster.Calendar.OptionallyBoundedMonthRange)
                                                   End Sub)
        End Sub

    End Class

End Namespace

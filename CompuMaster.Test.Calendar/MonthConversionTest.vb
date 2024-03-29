﻿Option Explicit On
Option Strict On

Imports NUnit.Framework

Namespace CompuMaster.Test.Calendar

    <TestFixture()> Public Class MonthConversionTest

        <Test> Public Sub ZeroableMonthToMonth()
            Assert.Catch(Of InvalidCastException)(Sub()
                                                      Dim Value As CompuMaster.Calendar.Month = CType(New CompuMaster.Calendar.ZeroableMonth(2000, 0), CompuMaster.Calendar.Month)
                                                  End Sub)
            Assert.AreEqual(New CompuMaster.Calendar.Month(2000, 1), CType(New CompuMaster.Calendar.ZeroableMonth(2000, 1), CompuMaster.Calendar.Month))
        End Sub

        <Test> Public Sub ZeroableMonthToString()
            Assert.AreEqual(Nothing, CType(Nothing, CompuMaster.Calendar.ZeroableMonth))
            Assert.AreEqual("0000-00", (New CompuMaster.Calendar.ZeroableMonth(0, 0).ToString))
            Assert.AreEqual("0001-01", (New CompuMaster.Calendar.ZeroableMonth(1, 1).ToString))
            Assert.AreEqual("2000-00", (New CompuMaster.Calendar.ZeroableMonth(2000, 0).ToString))
            Assert.AreEqual("2000-01", (New CompuMaster.Calendar.ZeroableMonth(2000, 1).ToString))
        End Sub

        <Test> Public Sub ZeroableMonthToInt32()
            Assert.AreEqual(200000, (New CompuMaster.Calendar.ZeroableMonth(2000, 0).ToInteger))
            Assert.AreEqual(200001, (New CompuMaster.Calendar.ZeroableMonth(2000, 1).ToInteger))
        End Sub

        <Test> Public Sub MonthToZeroableMonth()
            Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth(2000, 1), CType(New CompuMaster.Calendar.Month(2000, 1), CompuMaster.Calendar.ZeroableMonth))
            Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth(2000, 1), (New CompuMaster.Calendar.Month(2000, 1)).ToZeroableMonth())

            'Ensure value type behaviour
            Dim Origin As New CompuMaster.Calendar.Month(2000, 1)
            Dim ConvertedInitialValue As CompuMaster.Calendar.ZeroableMonth = Origin.ToZeroableMonth()
            Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth(2000, 1), ConvertedInitialValue)
            Origin.Year = 3000
            Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth(2000, 1), ConvertedInitialValue) 'must stay unchanged
        End Sub

        <Test> Public Sub MonthToString()
            Assert.AreEqual(Nothing, CType(Nothing, CompuMaster.Calendar.Month))
            Assert.AreEqual("0001-01", (New CompuMaster.Calendar.Month(1, 1).ToString))
            Assert.AreEqual("2000-01", (New CompuMaster.Calendar.Month(2000, 1).ToString))
        End Sub

        <Test> Public Sub MonthToInt32()
            Assert.AreEqual(200001, (New CompuMaster.Calendar.Month(2000, 1).ToInteger))
        End Sub

        <Test> Public Sub Int32ToMonth()
            Assert.AreEqual(New CompuMaster.Calendar.Month(2000, 1), CType(200001, CompuMaster.Calendar.Month))
            Assert.Catch(Of InvalidCastException)(Sub()
                                                      Dim Dummy As CompuMaster.Calendar.Month = (CType(-1, CompuMaster.Calendar.Month))
                                                  End Sub)
            Assert.Catch(Of InvalidCastException)(Sub()
                                                      Dim Dummy As CompuMaster.Calendar.Month = (CType(0, CompuMaster.Calendar.Month))
                                                  End Sub)
            Assert.Catch(Of InvalidCastException)(Sub()
                                                      Dim Dummy As CompuMaster.Calendar.Month = (CType(100, CompuMaster.Calendar.Month))
                                                  End Sub)
            Assert.AreEqual(101, CType(101, CompuMaster.Calendar.Month).ToInteger)
        End Sub

        <Test> Public Sub Int32ToZeroableMonth()
            Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth(2000, 0), CType(200000, CompuMaster.Calendar.ZeroableMonth))
            Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth(2000, 1), CType(200001, CompuMaster.Calendar.ZeroableMonth))
            Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth(9999, 12), CType(999912, CompuMaster.Calendar.ZeroableMonth))
            Assert.Catch(Of InvalidCastException)(Sub()
                                                      Dim Dummy As CompuMaster.Calendar.ZeroableMonth = (CType(-1, CompuMaster.Calendar.ZeroableMonth))
                                                  End Sub)
            Assert.AreEqual(0, CType(0, CompuMaster.Calendar.ZeroableMonth).ToInteger)
            Assert.Catch(Of InvalidCastException)(Sub()
                                                      Dim Dummy As CompuMaster.Calendar.ZeroableMonth = (CType(999999, CompuMaster.Calendar.ZeroableMonth))
                                                  End Sub)
        End Sub

        <Test> Public Sub Int64ToMonth()
            Assert.AreEqual(New CompuMaster.Calendar.Month(2000, 1), CType(200001L, CompuMaster.Calendar.Month))
        End Sub

        <Test> Public Sub Int64ToZeroableMonth()
            Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth(2000, 0), CType(200000L, CompuMaster.Calendar.ZeroableMonth))
            Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth(2000, 1), CType(200001L, CompuMaster.Calendar.ZeroableMonth))
        End Sub

        '<Test> Public Sub UInt32ToMonth()
        '    Assert.AreEqual(New CompuMaster.Calendar.Month(2000, 1), CType(CType(200001, UInteger), CompuMaster.Calendar.Month))
        'End Sub
        '
        '<Test> Public Sub UInt32ToZeroableMonth()
        '    Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth(2000, 0), CType(CType(200000, UInteger), CompuMaster.Calendar.ZeroableMonth))
        '    Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth(2000, 1), CType(CType(200001, UInteger), CompuMaster.Calendar.ZeroableMonth))
        'End Sub
        '
        '<Test> Public Sub UInt64ToMonth()
        '    Assert.AreEqual(New CompuMaster.Calendar.Month(2000, 1), CType(CType(200001, UInt64), CompuMaster.Calendar.Month))
        'End Sub
        '
        '<Test> Public Sub UInt64ToZeroableMonth()
        '    Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth(2000, 0), CType(CType(200000, UInt64), CompuMaster.Calendar.ZeroableMonth))
        '    Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth(2000, 1), CType(CType(200001, UInt64), CompuMaster.Calendar.ZeroableMonth))
        'End Sub

        <Test> Public Sub StringToMonth()
            Assert.AreEqual(New CompuMaster.Calendar.Month(2000, 1), CompuMaster.Calendar.Month.Parse("2000-01"))
            Assert.AreEqual(New CompuMaster.Calendar.Month(2000, 1), CType("2000-01", CompuMaster.Calendar.Month))
        End Sub

        <Test> Public Sub StringToZeroableMonth()
            Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth(2000, 0), CompuMaster.Calendar.ZeroableMonth.Parse("2000-00"))
            Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth(2000, 0), CType("2000-00", CompuMaster.Calendar.ZeroableMonth))
            Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth(2000, 1), CompuMaster.Calendar.ZeroableMonth.Parse("2000-01"))
            Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth(2000, 1), CType("2000-01", CompuMaster.Calendar.ZeroableMonth))
        End Sub

        <Test> Public Sub ImplicitConversionsOnCompileTime()
            Assert.Ignore("Test for compiler requires manual code check")
#If False Then
            Assert.AreEqual(New CompuMaster.Calendar.Month(2000, 1), ImplicitConversionTestHelperToMonth(200001))
            Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth(2000, 1), ImplicitConversionTestHelperToZeroableMonth(200001))
            Assert.AreEqual(New CompuMaster.Calendar.Month(2000, 1), ImplicitConversionTestHelperToMonth(200001L))
            Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth(2000, 1), ImplicitConversionTestHelperToZeroableMonth(200001L))
            'Assert.AreEqual(New CompuMaster.Calendar.Month(2000, 1), ImplicitConversionTestHelperToMonth(CType(200001, UInteger)))
            'Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth(2000, 1), ImplicitConversionTestHelperToZeroableMonth(CType(200001, UInteger)))
            'Assert.AreEqual(New CompuMaster.Calendar.Month(2000, 1), ImplicitConversionTestHelperToMonth(CType(200001, ULong)))
            'Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth(2000, 1), ImplicitConversionTestHelperToZeroableMonth(CType(200001, ULong)))
#End If
        End Sub

        Private Function ImplicitConversionTestHelperToMonth(value As CompuMaster.Calendar.Month) As CompuMaster.Calendar.Month
            Return value
        End Function

        Private Function ImplicitConversionTestHelperToZeroableMonth(value As CompuMaster.Calendar.ZeroableMonth) As CompuMaster.Calendar.ZeroableMonth
            Return value
        End Function

    End Class

End Namespace
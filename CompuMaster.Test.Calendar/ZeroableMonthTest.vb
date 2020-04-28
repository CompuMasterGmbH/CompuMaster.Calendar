﻿Imports NUnit.Framework

Namespace CompuMaster.Test.Calendar

    <TestFixture()> Public Class ZeroableMonthTest

        <Test> Public Sub Parse()
            Assert.AreEqual("2010-10", CompuMaster.Calendar.Month.Parse("2010-10").ToString)
            Assert.AreEqual("1900-01", CompuMaster.Calendar.Month.Parse("1900-01").ToString)
            Assert.AreEqual("2010-00", CompuMaster.Calendar.Month.Parse("2010-00").ToString)
            Assert.AreEqual("2010-10", CompuMaster.Calendar.ZeroableMonth.Parse("Oct/2010", "MMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US")).ToString)
            Assert.AreEqual("1900-01", CompuMaster.Calendar.ZeroableMonth.Parse("Jan/1900", "MMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US")).ToString)
            Assert.AreEqual("9999-12", CompuMaster.Calendar.ZeroableMonth.Parse("Dec/9999", "MMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US")).ToString)
            Assert.AreEqual("2010-00", CompuMaster.Calendar.ZeroableMonth.Parse("???/2010", "MMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US")).ToString)
            Assert.Throws(Of ArgumentNullException)(Sub()
                                                        CompuMaster.Calendar.ZeroableMonth.Parse(Nothing, "MMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US"))
                                                    End Sub)
            Assert.Throws(Of ArgumentNullException)(Sub()
                                                        CompuMaster.Calendar.ZeroableMonth.Parse("", "MMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US"))
                                                    End Sub)
            Assert.Throws(Of ArgumentException)(Sub()
                                                    CompuMaster.Calendar.ZeroableMonth.Parse("invalid-value", "MMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US"))
                                                End Sub)
            Assert.AreEqual("2010-01", CompuMaster.Calendar.ZeroableMonth.Parse("01/2010", "MM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US")).ToString)
            Assert.AreEqual("2010-10", CompuMaster.Calendar.ZeroableMonth.Parse("10/2010", "MM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US")).ToString)
            Assert.AreEqual("2010-01", CompuMaster.Calendar.ZeroableMonth.Parse("1/2010", "M/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US")).ToString)
            Assert.AreEqual("2010-01", CompuMaster.Calendar.ZeroableMonth.Parse("01/2010", "M/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US")).ToString)
            Assert.AreEqual("2010-10", CompuMaster.Calendar.ZeroableMonth.Parse("10/2010", "MM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US")).ToString)
            Assert.AreEqual("2001-10", CompuMaster.Calendar.ZeroableMonth.Parse("10/01", "MM/yy", System.Globalization.CultureInfo.GetCultureInfo("en-US")).ToString)
            Assert.AreEqual("1999-10", CompuMaster.Calendar.ZeroableMonth.Parse("10/99", "MM/yy", System.Globalization.CultureInfo.GetCultureInfo("en-US")).ToString)
            Assert.AreEqual("1999-00", CompuMaster.Calendar.ZeroableMonth.Parse("00/99", "MM/yy", System.Globalization.CultureInfo.GetCultureInfo("en-US")).ToString)
        End Sub

        <Test> Public Sub TryParse()
            Dim Buffer As CompuMaster.Calendar.ZeroableMonth = Nothing
            Assert.AreEqual(True, CompuMaster.Calendar.ZeroableMonth.TryParse("Oct/2010", "MMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US"), Buffer))
            Assert.AreEqual("2010-10", Buffer.ToString)
            Assert.AreEqual(True, CompuMaster.Calendar.ZeroableMonth.TryParse("Jan/1900", "MMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US"), Buffer))
            Assert.AreEqual("1900-01", Buffer.ToString)
            Assert.AreEqual(True, CompuMaster.Calendar.ZeroableMonth.TryParse("Dec/9999", "MMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US"), Buffer))
            Assert.AreEqual("9999-12", Buffer.ToString)
            Assert.AreEqual(True, CompuMaster.Calendar.ZeroableMonth.TryParse("Mrz/2010", "MMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("de-DE"), Buffer))
            Assert.AreEqual("2010-03", Buffer.ToString)
            Assert.AreEqual(True, CompuMaster.Calendar.ZeroableMonth.TryParse("Okt/2010", "MMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("de-DE"), Buffer))
            Assert.AreEqual("2010-10", Buffer.ToString)
            Assert.AreEqual(True, CompuMaster.Calendar.ZeroableMonth.TryParse("Jan/1900", "MMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("de-DE"), Buffer))
            Assert.AreEqual("1900-01", Buffer.ToString)
            Assert.AreEqual(True, CompuMaster.Calendar.ZeroableMonth.TryParse("Dez/9999", "MMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("de-DE"), Buffer))
            Assert.AreEqual("9999-12", Buffer.ToString)
            Assert.AreEqual(True, CompuMaster.Calendar.ZeroableMonth.TryParse("???/9999", "MMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("de-DE"), Buffer))
            Assert.AreEqual("9999-00", Buffer.ToString)
            Buffer = Nothing
            Assert.AreEqual(False, CompuMaster.Calendar.ZeroableMonth.TryParse(Nothing, "MMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US"), Buffer))
            Assert.AreEqual(Nothing, Buffer)
            Assert.AreEqual(False, CompuMaster.Calendar.ZeroableMonth.TryParse("", "MMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US"), Buffer))
            Assert.AreEqual(Nothing, Buffer)
            Assert.AreEqual(False, CompuMaster.Calendar.ZeroableMonth.TryParse("invalid-value", "MMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US"), Buffer))
            Assert.AreEqual(Nothing, Buffer)
            Assert.AreEqual(True, CompuMaster.Calendar.ZeroableMonth.TryParse("January/1900", "MMMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US"), Buffer))
            Assert.AreEqual("1900-01", Buffer.ToString)
            Assert.AreEqual(True, CompuMaster.Calendar.ZeroableMonth.TryParse("October/2010", "MMMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US"), Buffer))
            Assert.AreEqual("2010-10", Buffer.ToString)
            Assert.AreEqual(True, CompuMaster.Calendar.ZeroableMonth.TryParse("1900 January", "yyyy MMMM", System.Globalization.CultureInfo.GetCultureInfo("en-US"), Buffer))
            Assert.AreEqual("1900-01", Buffer.ToString)
            Assert.AreEqual(True, CompuMaster.Calendar.ZeroableMonth.TryParse("2010 October", "yyyy MMMM", System.Globalization.CultureInfo.GetCultureInfo("en-US"), Buffer))
            Assert.AreEqual("2010-10", Buffer.ToString)
            Assert.AreEqual(True, CompuMaster.Calendar.ZeroableMonth.TryParse("01/1900", "MM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US"), Buffer))
            Assert.AreEqual("1900-01", Buffer.ToString)
            Assert.AreEqual(True, CompuMaster.Calendar.ZeroableMonth.TryParse("10/2010", "MM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US"), Buffer))
            Assert.AreEqual("2010-10", Buffer.ToString)
            Assert.AreEqual(True, CompuMaster.Calendar.ZeroableMonth.TryParse("1/1900", "M/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US"), Buffer))
            Assert.AreEqual("1900-01", Buffer.ToString)
            Assert.AreEqual(True, CompuMaster.Calendar.ZeroableMonth.TryParse("10/2010", "M/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US"), Buffer))
            Assert.AreEqual("2010-10", Buffer.ToString)
            Assert.AreEqual(True, CompuMaster.Calendar.ZeroableMonth.TryParse("2010 ???", "yyyy MMMM", System.Globalization.CultureInfo.GetCultureInfo("en-US"), Buffer))
            Assert.AreEqual("2010-00", Buffer.ToString)
        End Sub

        <Test> Public Sub ParseFromUniqueShortName()
            Assert.AreEqual("2010-10", CompuMaster.Calendar.ZeroableMonth.ParseFromUniqueShortName("Oct/2010").ToString)
            Assert.AreEqual("1900-01", CompuMaster.Calendar.ZeroableMonth.ParseFromUniqueShortName("Jan/1900").ToString)
            Assert.AreEqual("9999-12", CompuMaster.Calendar.ZeroableMonth.ParseFromUniqueShortName("Dec/9999").ToString)
            Assert.AreEqual("1900-00", CompuMaster.Calendar.ZeroableMonth.ParseFromUniqueShortName("???/1900").ToString)
        End Sub

        <Test> Public Sub ZeroPeriod()
            Assert.AreEqual("2010-00", New CompuMaster.Calendar.ZeroableMonth(2010, 0).ToString)
            Assert.AreEqual("0001-00", New CompuMaster.Calendar.ZeroableMonth(1, 0).ToString)
            Assert.AreEqual("2010-00", New CompuMaster.Calendar.ZeroableMonth(2010, 4).ZeroPeriod.ToString)
            Assert.AreEqual("0001-00", New CompuMaster.Calendar.ZeroableMonth(1, 4).ZeroPeriod.ToString)
            Assert.AreEqual("0000-00", New CompuMaster.Calendar.ZeroableMonth(0, 4).ZeroPeriod.ToString)
        End Sub

        <Test> Public Sub FirstPeriod()
            Assert.AreEqual("2010-01", New CompuMaster.Calendar.ZeroableMonth(2010, 0).FirstPeriod.ToString)
            Assert.AreEqual("0001-01", New CompuMaster.Calendar.ZeroableMonth(1, 0).FirstPeriod.ToString)
            Assert.AreEqual("0000-01", New CompuMaster.Calendar.ZeroableMonth(0, 0).FirstPeriod.ToString)
        End Sub

        <Test> Public Sub LastPeriod()
            Assert.AreEqual("2010-12", New CompuMaster.Calendar.ZeroableMonth(2010, 0).LastPeriod.ToString)
            Assert.AreEqual("0001-12", New CompuMaster.Calendar.ZeroableMonth(1, 0).LastPeriod.ToString)
            Assert.AreEqual("0000-12", New CompuMaster.Calendar.ZeroableMonth(0, 0).LastPeriod.ToString)
        End Sub

        <Test> Public Sub NextPeriod()
            Assert.AreEqual("2011-01", New CompuMaster.Calendar.ZeroableMonth(2010, 12).NextPeriod.ToString)
            Assert.AreEqual("2010-01", New CompuMaster.Calendar.ZeroableMonth(2010, 0).NextPeriod.ToString)
            Assert.AreEqual("2010-02", New CompuMaster.Calendar.ZeroableMonth(2010, 1).NextPeriod.ToString)
            Assert.AreEqual("0001-01", New CompuMaster.Calendar.ZeroableMonth(1, 0).NextPeriod.ToString)
            Assert.AreEqual("0000-01", New CompuMaster.Calendar.ZeroableMonth(0, 0).NextPeriod.ToString)
        End Sub

        <Test> Public Sub PreviousPeriod()
            Assert.AreEqual("2010-11", New CompuMaster.Calendar.ZeroableMonth(2010, 12).PreviousPeriod.ToString)
            Assert.AreEqual("2009-12", New CompuMaster.Calendar.ZeroableMonth(2010, 0).PreviousPeriod.ToString)
            Assert.AreEqual("2009-12", New CompuMaster.Calendar.ZeroableMonth(2010, 1).PreviousPeriod.ToString)
            Assert.AreEqual("0000-12", New CompuMaster.Calendar.ZeroableMonth(1, 0).PreviousPeriod.ToString)
            Assert.Catch(Of ArgumentOutOfRangeException)(Sub()
                                                             Console.WriteLine(New CompuMaster.Calendar.ZeroableMonth(0, 0).PreviousPeriod.ToString)
                                                         End Sub)
        End Sub

        <Test()> Public Sub OperatorMinus()
            Assert.AreEqual(0, New CompuMaster.Calendar.ZeroableMonth(2012, 5) - New CompuMaster.Calendar.ZeroableMonth(2012, 5))
            Assert.AreEqual(1, New CompuMaster.Calendar.ZeroableMonth(2012, 5) - New CompuMaster.Calendar.ZeroableMonth(2012, 4))
            Assert.AreEqual(2, New CompuMaster.Calendar.ZeroableMonth(2012, 5) - New CompuMaster.Calendar.ZeroableMonth(2012, 3))
            Assert.AreEqual(1, New CompuMaster.Calendar.ZeroableMonth(2012, 1) - New CompuMaster.Calendar.ZeroableMonth(2011, 12))
            Assert.AreEqual(3, New CompuMaster.Calendar.ZeroableMonth(2012, 2) - New CompuMaster.Calendar.ZeroableMonth(2011, 11))
            Assert.AreEqual(15, New CompuMaster.Calendar.ZeroableMonth(2012, 2) - New CompuMaster.Calendar.ZeroableMonth(2010, 11))
            Assert.AreEqual(-1, New CompuMaster.Calendar.ZeroableMonth(2012, 4) - New CompuMaster.Calendar.ZeroableMonth(2012, 5))
            Assert.AreEqual(-2, New CompuMaster.Calendar.ZeroableMonth(2012, 3) - New CompuMaster.Calendar.ZeroableMonth(2012, 5))
            Assert.AreEqual(-1, New CompuMaster.Calendar.ZeroableMonth(2011, 12) - New CompuMaster.Calendar.ZeroableMonth(2012, 1))
            Assert.AreEqual(-3, New CompuMaster.Calendar.ZeroableMonth(2011, 11) - New CompuMaster.Calendar.ZeroableMonth(2012, 2))
            Assert.AreEqual(-15, New CompuMaster.Calendar.ZeroableMonth(2010, 11) - New CompuMaster.Calendar.ZeroableMonth(2012, 2))
            Assert.AreEqual(12, New CompuMaster.Calendar.ZeroableMonth(2020, 1) - New CompuMaster.Calendar.ZeroableMonth(2019, 1))
            Assert.AreEqual(12, New CompuMaster.Calendar.ZeroableMonth(2020, 0) - New CompuMaster.Calendar.ZeroableMonth(2019, 0))
            Assert.AreEqual(13, New CompuMaster.Calendar.ZeroableMonth(2020, 1) - New CompuMaster.Calendar.ZeroableMonth(2019, 0))
            Assert.AreEqual(14, New CompuMaster.Calendar.ZeroableMonth(2020, 2) - New CompuMaster.Calendar.ZeroableMonth(2019, 0))
            Assert.AreEqual(11, New CompuMaster.Calendar.ZeroableMonth(2020, 0) - New CompuMaster.Calendar.ZeroableMonth(2019, 1))
            Assert.AreEqual(11, New CompuMaster.Calendar.ZeroableMonth(2019, 12) - New CompuMaster.Calendar.ZeroableMonth(2019, 1))
        End Sub

        <Test()> Public Sub Compares()
            Dim value1 As New CompuMaster.Calendar.ZeroableMonth(2010, 1)
            Dim value2 As New CompuMaster.Calendar.ZeroableMonth(2009, 12)
            Dim value3 As New CompuMaster.Calendar.ZeroableMonth(2010, 1)
            Dim value3nulled As New CompuMaster.Calendar.ZeroableMonth(2010, 0)
            Assert.AreEqual(True, value1 > value2)
            Assert.AreEqual(False, value1 < value2)
            Assert.AreEqual(True, value1 >= value2)
            Assert.AreEqual(False, value1 <= value2)
            Assert.AreEqual(True, value1 >= value3)
            Assert.AreEqual(True, value1 <= value3)
            Assert.AreEqual(True, value1 = value3)
            Assert.AreEqual(value1, value3)
            Assert.AreEqual(False, value1 <> value3)
            Assert.AreEqual(False, value1 > value3)
            Assert.AreEqual(False, value1 < value3)
            Assert.AreNotEqual(Nothing, value3)
            Assert.AreNotEqual(value3, Nothing)
            Assert.AreNotEqual(DBNull.Value, value3)
            Assert.AreNotEqual(value3, DBNull.Value)
            Assert.AreNotEqual(New System.Text.StringBuilder, value3)
            Assert.AreNotEqual(value3, New System.Text.StringBuilder)
            Assert.AreNotEqual(String.Empty, value3)
            Assert.AreNotEqual(value3, String.Empty)
            Assert.AreNotEqual(5, value3)
            Assert.AreNotEqual(value3, 5)
            'nulled comparisons
            Assert.AreEqual(True, value2 <> value3nulled)
            Assert.AreEqual(True, value3 <> value3nulled)
            Assert.AreEqual(False, value3nulled <> value3nulled)
            Assert.AreEqual(False, value2 = value3nulled)
            Assert.AreEqual(False, value3 = value3nulled)
            Assert.AreEqual(True, value3nulled = value3nulled)

            Assert.AreEqual(True, value2 < value3nulled)
            Assert.AreEqual(False, value3 < value3nulled)
            Assert.AreEqual(False, value3nulled < value3nulled)
            Assert.AreEqual(False, value2 > value3nulled)
            Assert.AreEqual(True, value3 > value3nulled)
            Assert.AreEqual(False, value3nulled > value3nulled)

            Assert.AreEqual(True, value2 <= value3nulled)
            Assert.AreEqual(False, value3 <= value3nulled)
            Assert.AreEqual(True, value3nulled <= value3nulled)
            Assert.AreEqual(False, value2 >= value3nulled)
            Assert.AreEqual(True, value3 >= value3nulled)
            Assert.AreEqual(True, value3nulled >= value3nulled)

            value1 = Nothing
            value2 = Nothing
            Assert.AreEqual(True, value1 = value2)
            Assert.AreEqual(False, value1 <> value2)
            Assert.AreEqual(False, value1 < value2)
            Assert.AreEqual(False, value1 > value2)
            Assert.AreEqual(True, value1 <= value2)
            Assert.AreEqual(True, value1 >= value2)
            Assert.AreEqual(False, value1 = value3)
            Assert.AreEqual(True, value1 <> value3)
            Assert.AreEqual(True, value1 < value3)
            Assert.AreEqual(False, value1 > value3)
            Assert.AreEqual(True, value1 <= value3)
            Assert.AreEqual(False, value1 >= value3)
            Assert.AreEqual(False, value3 = value2)
            Assert.AreEqual(True, value3 <> value2)
            Assert.AreEqual(False, value3 < value2)
            Assert.AreEqual(True, value3 > value2)
            Assert.AreEqual(False, value3 <= value2)
            Assert.AreEqual(True, value3 >= value2)
        End Sub

        <Test()> Public Sub MonthNames()
            Dim value1 As New CompuMaster.Calendar.ZeroableMonth(2010, 1)
            Dim value2 As New CompuMaster.Calendar.ZeroableMonth(2009, 12)
            Dim value3 As New CompuMaster.Calendar.ZeroableMonth(2009, 0)
            Assert.AreEqual("Jan/2010", value1.UniqueShortName)
            Assert.AreEqual("Dec/2009", value2.UniqueShortName)
            Assert.AreEqual("Jan", value1.MonthShortName(System.Globalization.CultureInfo.GetCultureInfo("en-US")))
            Assert.AreEqual("Dec", value2.MonthShortName(System.Globalization.CultureInfo.GetCultureInfo("en-US")))
            Assert.AreEqual("Jan", value1.MonthShortName(System.Globalization.CultureInfo.GetCultureInfo("de-DE")))
            Assert.AreEqual("Dez", value2.MonthShortName(System.Globalization.CultureInfo.GetCultureInfo("de-DE")))
            Assert.AreEqual("January", value1.MonthName(System.Globalization.CultureInfo.GetCultureInfo("en-US")))
            Assert.AreEqual("December", value2.MonthName(System.Globalization.CultureInfo.GetCultureInfo("en-US")))
            Assert.AreEqual("Januar", value1.MonthName(System.Globalization.CultureInfo.GetCultureInfo("de-DE")))
            Assert.AreEqual("Dezember", value2.MonthName(System.Globalization.CultureInfo.GetCultureInfo("de-DE")))
            Assert.AreEqual("Jan", value1.MonthShortName("en-US"))
            Assert.AreEqual("Dec", value2.MonthShortName("en-US"))
            Assert.AreEqual("January", value1.MonthName("en-US"))
            Assert.AreEqual("December", value2.MonthName("en-US"))
            Assert.AreEqual("???/2009", value3.UniqueShortName)
            Assert.AreEqual("???", value3.MonthShortName(System.Globalization.CultureInfo.GetCultureInfo("en-US")))
            Assert.AreEqual("???", value3.MonthName("en-US"))
            Assert.AreEqual("???", value3.MonthName(System.Globalization.CultureInfo.GetCultureInfo("de-DE")))
        End Sub

        <Test> Public Sub Add()
            Dim value1 As New CompuMaster.Calendar.ZeroableMonth(2010, 1)
            Dim value2 As New CompuMaster.Calendar.ZeroableMonth(2007, 12)
            Dim value3 As New CompuMaster.Calendar.ZeroableMonth(2017, 12)
            Dim value4 As New CompuMaster.Calendar.ZeroableMonth(2011, 1)
            Dim value5 As New CompuMaster.Calendar.ZeroableMonth(2010, 5)
            Dim value6 As New CompuMaster.Calendar.ZeroableMonth(2009, 5)
            Assert.AreEqual(value4, value1.AddYears(1))
            Assert.AreEqual(value5, value1.AddMonths(4))
            Assert.AreEqual(value1, value6.AddMonths(8))
            Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth(2009, 12), value1.AddMonths(-1))
            Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth(2009, 2), value1.AddMonths(-11))
            Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth(2009, 1), value1.AddMonths(-12))
            Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth(2008, 12), value1.AddMonths(-13))
            Assert.AreEqual(value6, value1.AddMonths(-8))
            Assert.AreEqual(value3, value1.AddYears(8).AddMonths(-1))
            Assert.AreEqual(value2, value1.AddMonths(-3 * 12 + 11))
            Assert.AreEqual(value2, value1.AddMonths(-2 * 12 + -1))
            Assert.AreEqual(value2, value1.Add(-3, 11))
            Assert.AreEqual(value2, value1.Add(-2, -1))
            Assert.AreEqual(value1, value2.AddMonths(3 * 12 - 11))
            Assert.AreEqual(value1, value2.AddMonths(2 * 12 - -1))
            Assert.AreEqual(value1, value2.Add(3, -11))
            Assert.AreEqual(value1, value2.Add(2, 1))
            Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth(2008, 1), (New CompuMaster.Calendar.ZeroableMonth(2007, 12)).Add(0, 1))
            Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth(2008, 1), (New CompuMaster.Calendar.ZeroableMonth(2008, 0)).Add(0, 1))
            Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth(2008, 0), (New CompuMaster.Calendar.ZeroableMonth(2008, 0)).Add(0, 0))
            Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth(2008, 12), (New CompuMaster.Calendar.ZeroableMonth(2008, 0)).Add(0, 12))
            Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth(2009, 1), (New CompuMaster.Calendar.ZeroableMonth(2008, 0)).Add(0, 13))
            Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth(2008, 1), (New CompuMaster.Calendar.ZeroableMonth(2008, 1)).Add(0, 0))
            Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth(2007, 12), (New CompuMaster.Calendar.ZeroableMonth(2008, 0)).Add(0, -1))
            Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth(2007, 12), (New CompuMaster.Calendar.ZeroableMonth(2008, 0)).Add(0, -1))
            Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth(2007, 12), (New CompuMaster.Calendar.ZeroableMonth(2008, 1)).Add(0, -1))
            Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth(2007, 1), (New CompuMaster.Calendar.ZeroableMonth(2008, 0)).Add(0, -12))
            Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth(2006, 12), (New CompuMaster.Calendar.ZeroableMonth(2008, 0)).Add(0, -13))
        End Sub

        <Test> Public Sub Min()
            Dim value0 As CompuMaster.Calendar.ZeroableMonth = Nothing
            Dim value0initialized As New CompuMaster.Calendar.ZeroableMonth
            Dim value1 As New CompuMaster.Calendar.ZeroableMonth(2010, 1)
            Dim value2 As New CompuMaster.Calendar.ZeroableMonth(2007, 12)
            Dim value3 As New CompuMaster.Calendar.ZeroableMonth(2017, 12)
            Dim value4 As New CompuMaster.Calendar.ZeroableMonth(2011, 1)
            Dim value4nulled As New CompuMaster.Calendar.ZeroableMonth(2011, 0)
            Dim value5 As New CompuMaster.Calendar.ZeroableMonth(2010, 5)
            Dim value6 As New CompuMaster.Calendar.ZeroableMonth(2009, 5)
            Dim value6nulled As New CompuMaster.Calendar.ZeroableMonth(2009, 0)
            Assert.AreEqual(value0, CompuMaster.Calendar.ZeroableMonth.Min(value0, value0))
            Assert.AreEqual(value0, CompuMaster.Calendar.ZeroableMonth.Min(value0, value1))
            Assert.AreEqual(value0, CompuMaster.Calendar.ZeroableMonth.Min(value1, value0))
            Assert.AreEqual(value0, CompuMaster.Calendar.ZeroableMonth.Min(value0initialized, value0))
            Assert.AreEqual(value0, CompuMaster.Calendar.ZeroableMonth.Min(value0, value0initialized))
            Assert.AreEqual(value1, CompuMaster.Calendar.ZeroableMonth.Min(value1, value1))
            Assert.AreEqual(value2, CompuMaster.Calendar.ZeroableMonth.Min(value1, value2))
            Assert.AreEqual(value1, CompuMaster.Calendar.ZeroableMonth.Min(value1, value3))
            Assert.AreEqual(value1, CompuMaster.Calendar.ZeroableMonth.Min(value1, value4))
            Assert.AreEqual(value1, CompuMaster.Calendar.ZeroableMonth.Min(value1, value5))
            Assert.AreEqual(value1, CompuMaster.Calendar.ZeroableMonth.Min(value5, value1))
            Assert.AreEqual(value6, CompuMaster.Calendar.ZeroableMonth.Min(value1, value6))
            Assert.AreEqual(value6nulled, CompuMaster.Calendar.ZeroableMonth.Min(value6nulled, value6))
            Assert.AreEqual(value6nulled, CompuMaster.Calendar.ZeroableMonth.Min(value6, value6nulled))
            Assert.AreEqual(value4nulled, CompuMaster.Calendar.ZeroableMonth.Min(value4nulled, value4))
            Assert.AreEqual(value4nulled, CompuMaster.Calendar.ZeroableMonth.Min(value4, value4nulled))
            Assert.AreEqual(value1, CompuMaster.Calendar.ZeroableMonth.Min(value4nulled, value1))
            Assert.AreEqual(value1, CompuMaster.Calendar.ZeroableMonth.Min(value1, value4nulled))
            Assert.AreEqual(value6nulled, CompuMaster.Calendar.ZeroableMonth.Min(value4nulled, value6nulled))
        End Sub

        <Test> Public Sub Max()
            Dim value0 As CompuMaster.Calendar.ZeroableMonth = Nothing
            Dim value0initialized As New CompuMaster.Calendar.ZeroableMonth
            Dim value1 As New CompuMaster.Calendar.ZeroableMonth(2010, 1)
            Dim value2 As New CompuMaster.Calendar.ZeroableMonth(2007, 12)
            Dim value3 As New CompuMaster.Calendar.ZeroableMonth(2017, 12)
            Dim value4 As New CompuMaster.Calendar.ZeroableMonth(2011, 1)
            Dim value4nulled As New CompuMaster.Calendar.ZeroableMonth(2011, 0)
            Dim value5 As New CompuMaster.Calendar.ZeroableMonth(2010, 5)
            Dim value6 As New CompuMaster.Calendar.ZeroableMonth(2009, 5)
            Dim value6nulled As New CompuMaster.Calendar.ZeroableMonth(2009, 0)
            Assert.AreEqual(value0, CompuMaster.Calendar.ZeroableMonth.Max(value0, value0))
            Assert.AreEqual(value1, CompuMaster.Calendar.ZeroableMonth.Max(value0, value1))
            Assert.AreEqual(value1, CompuMaster.Calendar.ZeroableMonth.Max(value1, value0))
            Assert.AreEqual(value0initialized, CompuMaster.Calendar.ZeroableMonth.Max(value0initialized, value0))
            Assert.AreEqual(value0initialized, CompuMaster.Calendar.ZeroableMonth.Max(value0, value0initialized))
            Assert.AreEqual(value1, CompuMaster.Calendar.ZeroableMonth.Max(value1, value1))
            Assert.AreEqual(value1, CompuMaster.Calendar.ZeroableMonth.Max(value1, value2))
            Assert.AreEqual(value3, CompuMaster.Calendar.ZeroableMonth.Max(value1, value3))
            Assert.AreEqual(value4, CompuMaster.Calendar.ZeroableMonth.Max(value1, value4))
            Assert.AreEqual(value5, CompuMaster.Calendar.ZeroableMonth.Max(value1, value5))
            Assert.AreEqual(value5, CompuMaster.Calendar.ZeroableMonth.Max(value5, value1))
            Assert.AreEqual(value1, CompuMaster.Calendar.ZeroableMonth.Max(value1, value6))
            Assert.AreEqual(value6, CompuMaster.Calendar.ZeroableMonth.Max(value6nulled, value6))
            Assert.AreEqual(value6, CompuMaster.Calendar.ZeroableMonth.Max(value6, value6nulled))
            Assert.AreEqual(value4, CompuMaster.Calendar.ZeroableMonth.Max(value4nulled, value4))
            Assert.AreEqual(value4, CompuMaster.Calendar.ZeroableMonth.Max(value4, value4nulled))
            Assert.AreEqual(value4nulled, CompuMaster.Calendar.ZeroableMonth.Max(value4nulled, value1))
            Assert.AreEqual(value4nulled, CompuMaster.Calendar.ZeroableMonth.Max(value1, value4nulled))
            Assert.AreEqual(value4nulled, CompuMaster.Calendar.ZeroableMonth.Max(value4nulled, value6nulled))
        End Sub

        <Test> Public Sub Conversions()
            Assert.AreEqual(Nothing, CType(CType(Nothing, CompuMaster.Calendar.Month), CompuMaster.Calendar.ZeroableMonth))
            Assert.AreEqual(Nothing, CType(CType(Nothing, CompuMaster.Calendar.ZeroableMonth), CompuMaster.Calendar.Month))
            Assert.AreEqual(Nothing, CType(CType(Nothing, CompuMaster.Calendar.Month), String))
            Assert.AreEqual(Nothing, CType(CType(Nothing, CompuMaster.Calendar.ZeroableMonth), String))
            Assert.AreEqual(Nothing, CType(CType(Nothing, String), CompuMaster.Calendar.ZeroableMonth))
            Assert.AreEqual(Nothing, CType(CType(Nothing, String), CompuMaster.Calendar.Month))
            Assert.AreEqual(Nothing, CType(String.Empty, CompuMaster.Calendar.ZeroableMonth))
            Assert.AreEqual(Nothing, CType(String.Empty, CompuMaster.Calendar.Month))
            Assert.AreEqual("2020-10", CType(New CompuMaster.Calendar.ZeroableMonth(2020, 10), String))
            Assert.AreEqual("2020-10", CType("2020-10", CompuMaster.Calendar.ZeroableMonth).ToString)
            Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth(2020, 10), CType(New CompuMaster.Calendar.Month(2020, 10), CompuMaster.Calendar.ZeroableMonth))
            Assert.Catch(Of InvalidCastException)(Sub()
                                                      Dim Dummy As CompuMaster.Calendar.ZeroableMonth
                                                      Dummy = CType(New CompuMaster.Calendar.ZeroableMonth(2020, 0), CompuMaster.Calendar.Month)
                                                  End Sub)
        End Sub

        <Test> Public Sub SmallerBiggerEquals()
            Assert.IsTrue(New CompuMaster.Calendar.ZeroableMonth(2010, 9) = New CompuMaster.Calendar.ZeroableMonth(2010, 9))
            Assert.IsTrue(New CompuMaster.Calendar.ZeroableMonth(2010, 1) <> New CompuMaster.Calendar.ZeroableMonth(2010, 9))
            Assert.IsTrue(New CompuMaster.Calendar.ZeroableMonth(2010, 1) < New CompuMaster.Calendar.ZeroableMonth(2010, 9))
            Assert.IsTrue(New CompuMaster.Calendar.ZeroableMonth(2010, 9) > New CompuMaster.Calendar.ZeroableMonth(2010, 1))
            Assert.IsFalse(New CompuMaster.Calendar.ZeroableMonth(2010, 9) <> New CompuMaster.Calendar.ZeroableMonth(2010, 9))
            Assert.IsFalse(New CompuMaster.Calendar.ZeroableMonth(2010, 1) = New CompuMaster.Calendar.ZeroableMonth(2010, 9))
            Assert.IsFalse(New CompuMaster.Calendar.ZeroableMonth(2010, 1) > New CompuMaster.Calendar.ZeroableMonth(2010, 9))
            Assert.IsFalse(New CompuMaster.Calendar.ZeroableMonth(2010, 9) < New CompuMaster.Calendar.ZeroableMonth(2010, 1))

            Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth(2010, 9), New CompuMaster.Calendar.ZeroableMonth(2010, 9))
            Assert.AreNotEqual(New CompuMaster.Calendar.ZeroableMonth(2010, 1), New CompuMaster.Calendar.ZeroableMonth(2010, 9))
            'Assert.Less(New CompuMaster.Calendar.ZeroableMonth(2010, 1), New CompuMaster.Calendar.ZeroableMonth(2010, 9))
            'Assert.Greater(New CompuMaster.Calendar.ZeroableMonth(2010, 9), New CompuMaster.Calendar.ZeroableMonth(2010, 1))

            Assert.IsTrue(New CompuMaster.Calendar.ZeroableMonth(2010, 0) = New CompuMaster.Calendar.ZeroableMonth(2010, 0))
            Assert.IsTrue(New CompuMaster.Calendar.ZeroableMonth(2010, 0) <> New CompuMaster.Calendar.ZeroableMonth(2010, 1))
            Assert.IsTrue(New CompuMaster.Calendar.ZeroableMonth(2010, 0) < New CompuMaster.Calendar.ZeroableMonth(2010, 1))
            Assert.IsTrue(New CompuMaster.Calendar.ZeroableMonth(2010, 0) > New CompuMaster.Calendar.ZeroableMonth(2009, 12))
            Assert.IsFalse(New CompuMaster.Calendar.ZeroableMonth(2010, 0) <> New CompuMaster.Calendar.ZeroableMonth(2010, 0))
            Assert.IsFalse(New CompuMaster.Calendar.ZeroableMonth(2010, 0) = New CompuMaster.Calendar.ZeroableMonth(2010, 1))
            Assert.IsFalse(New CompuMaster.Calendar.ZeroableMonth(2010, 0) > New CompuMaster.Calendar.ZeroableMonth(2010, 1))
            Assert.IsFalse(New CompuMaster.Calendar.ZeroableMonth(2010, 0) < New CompuMaster.Calendar.ZeroableMonth(2009, 12))
        End Sub

        <Test> Public Sub Sorting()
            Dim Values As New Generic.List(Of CompuMaster.Calendar.ZeroableMonth)
            Values.Add(New CompuMaster.Calendar.ZeroableMonth(2010, 11))
            Values.Add(New CompuMaster.Calendar.ZeroableMonth(1999, 10))
            Values.Add(New CompuMaster.Calendar.ZeroableMonth(2010, 0))
            Values.Add(New CompuMaster.Calendar.ZeroableMonth(2010, 10))
            Values.Add(New CompuMaster.Calendar.ZeroableMonth)
            Values.Add(New CompuMaster.Calendar.ZeroableMonth(2010, 9))

            Values.Sort()

            Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth, Values(0))
            Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth(1999, 10), Values(1))
            Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth(2010, 0), Values(2))
            Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth(2010, 9), Values(3))
            Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth(2010, 10), Values(4))
            Assert.AreEqual(New CompuMaster.Calendar.ZeroableMonth(2010, 11), Values(5))
        End Sub

    End Class

End Namespace
Imports NUnit.Framework

Namespace CompuMaster.Test.Calendar

    <TestFixture()> Public Class Month

        <Test()> Public Sub OperatorMinus()
            Assert.AreEqual(0, New CompuMaster.Calendar.Month(2012, 5) - New CompuMaster.Calendar.Month(2012, 5))
            Assert.AreEqual(1, New CompuMaster.Calendar.Month(2012, 5) - New CompuMaster.Calendar.Month(2012, 4))
            Assert.AreEqual(2, New CompuMaster.Calendar.Month(2012, 5) - New CompuMaster.Calendar.Month(2012, 3))
            Assert.AreEqual(2, New CompuMaster.Calendar.Month(2012, 1) - New CompuMaster.Calendar.Month(2011, 12))
            Assert.AreEqual(4, New CompuMaster.Calendar.Month(2012, 2) - New CompuMaster.Calendar.Month(2011, 11))
            Assert.AreEqual(16, New CompuMaster.Calendar.Month(2012, 2) - New CompuMaster.Calendar.Month(2010, 11))
            Assert.AreEqual(-1, New CompuMaster.Calendar.Month(2012, 4) - New CompuMaster.Calendar.Month(2012, 5))
            Assert.AreEqual(-2, New CompuMaster.Calendar.Month(2012, 3) - New CompuMaster.Calendar.Month(2012, 5))
            Assert.AreEqual(-2, New CompuMaster.Calendar.Month(2011, 12) - New CompuMaster.Calendar.Month(2012, 1))
            Assert.AreEqual(-4, New CompuMaster.Calendar.Month(2011, 11) - New CompuMaster.Calendar.Month(2012, 2))
            Assert.AreEqual(-16, New CompuMaster.Calendar.Month(2010, 11) - New CompuMaster.Calendar.Month(2012, 2))
        End Sub

        <Test()> Public Sub Compares()
            Dim value1 As New CompuMaster.Calendar.Month(2010, 1)
            Dim value2 As New CompuMaster.Calendar.Month(2009, 12)
            Dim value3 As New CompuMaster.Calendar.Month(2010, 1)
            Assert.AreEqual(True, value1 > value2)
            Assert.AreEqual(False, value1 < value2)
            Assert.AreEqual(True, value1 >= value2)
            Assert.AreEqual(False, value1 <= value2)
            Assert.AreEqual(True, value1 >= value3)
            Assert.AreEqual(True, value1 <= value3)
            Assert.AreEqual(True, value1 = value3)
            Assert.AreEqual(False, value1 <> value3)
            Assert.AreEqual(False, value1 > value3)
            Assert.AreEqual(False, value1 < value3)
        End Sub

        <Test()> Public Sub MonthNames()
            Dim value1 As New CompuMaster.Calendar.Month(2010, 1)
            Dim value2 As New CompuMaster.Calendar.Month(2009, 12)
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
        End Sub

    End Class

End Namespace
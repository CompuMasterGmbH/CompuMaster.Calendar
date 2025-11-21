Imports NUnit.Framework

Namespace CompuMaster.Test.Calendar

    <TestFixture()> Public Class MonthTest

        <Test> Public Sub Init()
            Dim TestMonthNo As Integer, TestYear As Integer

            Assert.AreEqual(1, (New CompuMaster.Calendar.Month).Year, "Init with zero alias empty value")
            Assert.AreEqual(1, (New CompuMaster.Calendar.Month).Month, "Init with zero alias empty value")
            Assert.AreEqual(New CompuMaster.Calendar.Month, New CompuMaster.Calendar.Month(1, 1), "Init with zero alias empty value")
            Assert.AreEqual(CompuMaster.Calendar.Month.MinValue, New CompuMaster.Calendar.Month(1, 1), "Init with zero alias min value")

            TestYear = 0
            TestMonthNo = 0
            Assert.Catch(Of ArgumentOutOfRangeException)(Function()
                                                             Return New CompuMaster.Calendar.Month(TestYear, TestMonthNo)
                                                         End Function)

            For MonthCounter As Integer = 1 To 12
                TestYear = 1
                Assert.AreEqual(TestYear, (New CompuMaster.Calendar.Month(TestYear, MonthCounter).Year))
                Assert.AreEqual(MonthCounter, (New CompuMaster.Calendar.Month(TestYear, MonthCounter).Month))

                TestYear = 9999
                Assert.AreEqual(TestYear, (New CompuMaster.Calendar.Month(TestYear, MonthCounter).Year))
                Assert.AreEqual(MonthCounter, (New CompuMaster.Calendar.Month(TestYear, MonthCounter).Month))
            Next

            TestYear = 1
            TestMonthNo = 13
            Assert.Catch(Of ArgumentOutOfRangeException)(Function()
                                                             Return New CompuMaster.Calendar.Month(TestYear, TestMonthNo)
                                                         End Function)

            TestYear = 1
            TestMonthNo = 0
            Assert.Catch(Of ArgumentOutOfRangeException)(Function()
                                                             Return New CompuMaster.Calendar.Month(TestYear, TestMonthNo)
                                                         End Function)
            TestYear = 0
            TestMonthNo = 1
            Assert.Catch(Of ArgumentOutOfRangeException)(Function()
                                                             Return New CompuMaster.Calendar.Month(TestYear, TestMonthNo)
                                                         End Function)

            TestYear = 10000
            TestMonthNo = 1
            Assert.Catch(Of ArgumentOutOfRangeException)(Function()
                                                             Return New CompuMaster.Calendar.Month(TestYear, TestMonthNo)
                                                         End Function)

        End Sub

        <Test> Public Sub Parse()
            Assert.AreEqual("2010-10", CompuMaster.Calendar.Month.Parse("2010-10").ToString)
            Assert.AreEqual("1900-01", CompuMaster.Calendar.Month.Parse("1900-01").ToString)
            Assert.AreEqual("2010-10", CompuMaster.Calendar.Month.Parse("Oct/2010", "MMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US")).ToString)
            Assert.AreEqual("1900-01", CompuMaster.Calendar.Month.Parse("Jan/1900", "MMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US")).ToString)
            Assert.AreEqual("9999-12", CompuMaster.Calendar.Month.Parse("Dec/9999", "MMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US")).ToString)
            Assert.Throws(Of ArgumentNullException)(Sub()
                                                        CompuMaster.Calendar.Month.Parse(Nothing, "MMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US"))
                                                    End Sub)
            Assert.Throws(Of ArgumentNullException)(Sub()
                                                        CompuMaster.Calendar.Month.Parse("", "MMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US"))
                                                    End Sub)
            Assert.Throws(Of ArgumentException)(Sub()
                                                    CompuMaster.Calendar.Month.Parse("invalid-value", "MMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US"))
                                                End Sub)
            Assert.AreEqual("2010-01", CompuMaster.Calendar.Month.Parse("01/2010", "MM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US")).ToString)
            Assert.AreEqual("2010-10", CompuMaster.Calendar.Month.Parse("10/2010", "MM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US")).ToString)
            Assert.AreEqual("2010-01", CompuMaster.Calendar.Month.Parse("1/2010", "M/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US")).ToString)
            Assert.AreEqual("2010-01", CompuMaster.Calendar.Month.Parse("01/2010", "M/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US")).ToString)
            Assert.AreEqual("2010-10", CompuMaster.Calendar.Month.Parse("10/2010", "MM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US")).ToString)
            Assert.AreEqual("2001-10", CompuMaster.Calendar.Month.Parse("10/01", "MM/yy", System.Globalization.CultureInfo.GetCultureInfo("en-US")).ToString)
            Assert.AreEqual("1999-10", CompuMaster.Calendar.Month.Parse("10/99", "MM/yy", System.Globalization.CultureInfo.GetCultureInfo("en-US")).ToString)
            Assert.AreEqual("2023-03", CompuMaster.Calendar.Month.Parse("März 2023", "MMMM yyyy", System.Globalization.CultureInfo.GetCultureInfo("de-DE")).ToString)

            Dim CustomMonths As String() = New String() {"Jan", "Feb", "Mär", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"}
            Assert.AreEqual("2020-03", CompuMaster.Calendar.Month.Parse("Mär 2020", "CCC yyyy", System.Globalization.CultureInfo.GetCultureInfo("de-DE"), CustomMonths).ToString, "#100")
            Assert.AreEqual("2020-12", CompuMaster.Calendar.Month.Parse("Dez 2020", "MMM yyyy", System.Globalization.CultureInfo.GetCultureInfo("de-DE"), CustomMonths).ToString, "#102")
            Assert.AreEqual("2020-12", CompuMaster.Calendar.Month.Parse("Dez 2020", "MMM yyyy", System.Globalization.CultureInfo.GetCultureInfo("de-DE")).ToString, "#104")
            Assert.AreEqual("2020-03", CompuMaster.Calendar.Month.Parse("Mär/2020", "CCC/yyyy", System.Globalization.CultureInfo.GetCultureInfo("de-DE"), CustomMonths).ToString, "#110")
        End Sub

        <Test> Public Sub TryParse()
            Dim Buffer As CompuMaster.Calendar.Month = Nothing
            Assert.AreEqual(True, CompuMaster.Calendar.Month.TryParse("Oct/2010", "MMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US"), Buffer), "#010")
            Assert.AreEqual("2010-10", Buffer.ToString, "#020")
            Assert.AreEqual(True, CompuMaster.Calendar.Month.TryParse("Jan/1900", "MMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US"), Buffer), "#011")
            Assert.AreEqual("1900-01", Buffer.ToString, "#021")
            Assert.AreEqual(True, CompuMaster.Calendar.Month.TryParse("Dec/9999", "MMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US"), Buffer), "#012")
            Assert.AreEqual("9999-12", Buffer.ToString, "#022")
            Assert.IsFalse(CompuMaster.Calendar.Month.TryParse("BeginValue", "yyyy-MM", System.Globalization.CultureInfo.InvariantCulture, Nothing))
            Dim ExpectedMarchShortNameOnCurrentPlatform As String = New DateTime(2010, 3, 1).ToString("MMM""/""yyyy", System.Globalization.CultureInfo.GetCultureInfo("de-DE")) 'Linux systems/Mono: "Mär", Windows systems: "Mrz"
            System.Console.WriteLine("SPECIAL CASE HANDLING: month March in German culture appears differently between windows and linux/mono systems")
            System.Console.WriteLine("SPECIAL CASE HANDLING: month March in German culture expected as """ & ExpectedMarchShortNameOnCurrentPlatform & """")
            Assert.AreEqual(True, CompuMaster.Calendar.Month.TryParse(ExpectedMarchShortNameOnCurrentPlatform, "MMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("de-DE"), Buffer), "#013")
            Assert.AreEqual("2010-03", Buffer.ToString, "#023")
            Assert.AreEqual(True, CompuMaster.Calendar.Month.TryParse("Okt/2010", "MMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("de-DE"), Buffer), "#014")
            Assert.AreEqual("2010-10", Buffer.ToString, "#024")
            Assert.AreEqual(True, CompuMaster.Calendar.Month.TryParse("Jan/1900", "MMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("de-DE"), Buffer), "#015")
            Assert.AreEqual("1900-01", Buffer.ToString, "#025")
            Assert.AreEqual(True, CompuMaster.Calendar.Month.TryParse("Dez/9999", "MMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("de-DE"), Buffer), "#016")
            Assert.AreEqual("9999-12", Buffer.ToString, "#026")
            Buffer = Nothing
            Assert.AreEqual(False, CompuMaster.Calendar.Month.TryParse(Nothing, "MMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US"), Buffer), "#030")
            Assert.AreEqual(Nothing, Buffer, "#050")
            Assert.AreEqual(False, CompuMaster.Calendar.Month.TryParse("", "MMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US"), Buffer), "#031")
            Assert.AreEqual(Nothing, Buffer, "#051")
            Assert.AreEqual(False, CompuMaster.Calendar.Month.TryParse("invalid-value", "MMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US"), Buffer), "#032")
            Assert.AreEqual(Nothing, Buffer, "#052")
            Assert.AreEqual(True, CompuMaster.Calendar.Month.TryParse("January/1900", "MMMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US"), Buffer), "#033")
            Assert.AreEqual("1900-01", Buffer.ToString, "#053")
            Assert.AreEqual(True, CompuMaster.Calendar.Month.TryParse("October/2010", "MMMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US"), Buffer), "#034")
            Assert.AreEqual("2010-10", Buffer.ToString, "#054")
            Assert.AreEqual(True, CompuMaster.Calendar.Month.TryParse("1900 January", "yyyy MMMM", System.Globalization.CultureInfo.GetCultureInfo("en-US"), Buffer), "#035")
            Assert.AreEqual("1900-01", Buffer.ToString, "#055")
            Assert.AreEqual(True, CompuMaster.Calendar.Month.TryParse("2010 October", "yyyy MMMM", System.Globalization.CultureInfo.GetCultureInfo("en-US"), Buffer), "#036")
            Assert.AreEqual("2010-10", Buffer.ToString, "#056")
            Assert.AreEqual(True, CompuMaster.Calendar.Month.TryParse("01/1900", "MM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US"), Buffer), "#037")
            Assert.AreEqual("1900-01", Buffer.ToString, "#057")
            Assert.AreEqual(True, CompuMaster.Calendar.Month.TryParse("10/2010", "MM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US"), Buffer), "#038")
            Assert.AreEqual("2010-10", Buffer.ToString, "#058")
            Assert.AreEqual(True, CompuMaster.Calendar.Month.TryParse("1/1900", "M/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US"), Buffer), "#039")
            Assert.AreEqual("1900-01", Buffer.ToString, "#059")
            Assert.AreEqual(True, CompuMaster.Calendar.Month.TryParse("10/2010", "M/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US"), Buffer), "#040")
            Assert.AreEqual("2010-10", Buffer.ToString, "#060")

            Dim CustomMonths As String() = New String() {"Jan", "Feb", "Mär", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"}
            Assert.AreEqual(True, CompuMaster.Calendar.Month.TryParse("Mär 2020", "CCC yyyy", System.Globalization.CultureInfo.GetCultureInfo("de-DE"), CustomMonths, Buffer), "#100")
            Assert.AreEqual("2020-03", Buffer.ToString, "#101")
            Assert.AreEqual(True, CompuMaster.Calendar.Month.TryParse("Dez 2020", "MMM yyyy", System.Globalization.CultureInfo.GetCultureInfo("de-DE"), CustomMonths, Buffer), "#102")
            Assert.AreEqual("2020-12", Buffer.ToString, "#103")
            Assert.AreEqual(True, CompuMaster.Calendar.Month.TryParse("Dez 2020", "MMM yyyy", System.Globalization.CultureInfo.GetCultureInfo("de-DE"), Buffer), "#104")
            Assert.AreEqual("2020-12", Buffer.ToString, "#105")
            Assert.AreEqual(True, CompuMaster.Calendar.Month.TryParse("Mär/2020", "CCC/yyyy", System.Globalization.CultureInfo.GetCultureInfo("de-DE"), CustomMonths, Buffer), "#110")
            Assert.AreEqual("2020-03", Buffer.ToString, "#111")
        End Sub

        <Test> Public Sub TryParsePerformance()
            Const TableXmlSchema As String =
                "<?xml version=""1.0"" standalone=""yes""?>" & ControlChars.CrLf &
                "<xs:schema id=""NewDataSet"" xmlns="""" xmlns:xs=""http://www.w3.org/2001/XMLSchema"" xmlns:msdata=""urn:schemas-microsoft-com:xml-msdata"">" & ControlChars.CrLf &
                "  <xs:element name=""NewDataSet"" msdata:IsDataSet=""true"" msdata:MainDataTable=""SuSa"" msdata:UseCurrentLocale=""true"">" & ControlChars.CrLf &
                "    <xs:complexType>" & ControlChars.CrLf &
                "      <xs:choice minOccurs=""0"" maxOccurs=""unbounded"">" & ControlChars.CrLf &
                "        <xs:element name=""SuSa"">" & ControlChars.CrLf &
                "          <xs:complexType>" & ControlChars.CrLf &
                "            <xs:sequence>" & ControlChars.CrLf &
                "              <xs:element name=""AccountNumber"" type=""xs:string"" />" & ControlChars.CrLf &
                "              <xs:element name=""AccountName"" type=""xs:string"" minOccurs=""0"" />" & ControlChars.CrLf &
                "              <xs:element name=""BeginValue"" type=""xs:double"" minOccurs=""0"" />" & ControlChars.CrLf &
                "              <xs:element name=""_x0032_024-01"" type=""xs:double"" minOccurs=""0"" />" & ControlChars.CrLf &
                "              <xs:element name=""_x0032_024-02"" type=""xs:double"" minOccurs=""0"" />" & ControlChars.CrLf &
                "              <xs:element name=""_x0032_024-03"" type=""xs:double"" minOccurs=""0"" />" & ControlChars.CrLf &
                "              <xs:element name=""_x0032_024-04"" type=""xs:double"" minOccurs=""0"" />" & ControlChars.CrLf &
                "              <xs:element name=""_x0032_024-05"" type=""xs:double"" minOccurs=""0"" />" & ControlChars.CrLf &
                "              <xs:element name=""_x0032_024-06"" type=""xs:double"" minOccurs=""0"" />" & ControlChars.CrLf &
                "              <xs:element name=""_x0032_024-07"" type=""xs:double"" minOccurs=""0"" />" & ControlChars.CrLf &
                "              <xs:element name=""_x0032_024-08"" type=""xs:double"" minOccurs=""0"" />" & ControlChars.CrLf &
                "              <xs:element name=""_x0032_024-09"" type=""xs:double"" minOccurs=""0"" />" & ControlChars.CrLf &
                "              <xs:element name=""_x0032_024-10"" type=""xs:double"" minOccurs=""0"" />" & ControlChars.CrLf &
                "              <xs:element name=""_x0032_024-11"" type=""xs:double"" minOccurs=""0"" />" & ControlChars.CrLf &
                "              <xs:element name=""_x0032_024-12"" type=""xs:double"" minOccurs=""0"" />" & ControlChars.CrLf &
                "              <xs:element name=""EndValue"" type=""xs:double"" minOccurs=""0"" />" & ControlChars.CrLf &
                "            </xs:sequence>" & ControlChars.CrLf &
                "          </xs:complexType>" & ControlChars.CrLf &
                "        </xs:element>" & ControlChars.CrLf &
                "      </xs:choice>" & ControlChars.CrLf &
                "    </xs:complexType>" & ControlChars.CrLf &
                "    <xs:unique name=""Constraint1"" msdata:PrimaryKey=""true"">" & ControlChars.CrLf &
                "      <xs:selector xpath="".//SuSa"" />" & ControlChars.CrLf &
                "      <xs:field xpath=""AccountNumber"" />" & ControlChars.CrLf &
                "    </xs:unique>" & ControlChars.CrLf &
                "  </xs:element>" & ControlChars.CrLf &
                "</xs:schema>" & ControlChars.CrLf

            Dim table As New System.Data.DataTable
            table.ReadXmlSchema(System.Xml.XmlReader.Create(New System.IO.StringReader(TableXmlSchema)))
            TryParsePerformance_LastFilledDataPeriod(table)
        End Sub

        Private Function TryParsePerformance_LastFilledDataPeriod(table As System.Data.DataTable) As CompuMaster.Calendar.Month
            Dim Result As CompuMaster.Calendar.Month = Nothing
            For LoopCounter As Integer = 0 To 200
                For MyCounter As Integer = 0 To table.Columns.Count - 1
                    If CompuMaster.Calendar.Month.TryParse(table.Columns(MyCounter).ColumnName, "yyyy-MM", System.Globalization.CultureInfo.InvariantCulture, Nothing) Then
                        Dim FoundDataMonth As New CompuMaster.Calendar.Month(table.Columns(MyCounter).ColumnName)
                        If LoopCounter = 0 Then
                            System.Console.WriteLine("Found data month: " & FoundDataMonth.ToString)
                            Assert.NotNull(FoundDataMonth)
                        End If
                    End If
                Next
            Next
            Return Result
        End Function

        <Test> Public Sub EncodeForRegEx()
            Dim Items As New System.Collections.Generic.List(Of String)
            Dim Expected As New System.Collections.Generic.List(Of String)
            Items.Add("???") : Expected.Add("\?\?\?")
            Items.Add("Mär") : Expected.Add("Mär")
            Items.Add("Mär.") : Expected.Add("Mär\.")
            Assert.AreEqual(Expected.ToArray, CompuMaster.Calendar.Month.EncodeForRegEx(Items.ToArray))
        End Sub

        ''' <summary>
        ''' March in German is formatted differently with MMM between windows and linux platforms (Mrz vs. Mär)
        ''' </summary>
                <Test> Public Sub MarchInGermanToStringAndReParse()
            Dim Buffer As CompuMaster.Calendar.Month = Nothing
            Dim TestNumber As Double

            TestNumber = 100
            Dim ExpectedMarchShortNameOnCurrentPlatform As String = New DateTime(2010, 3, 1).ToString("MMM""/""yyyy", System.Globalization.CultureInfo.GetCultureInfo("de-DE")) 'Linux systems/Mono: "Mär", Windows systems: "Mrz"
            System.Console.WriteLine("SPECIAL CASE HANDLING: month March in German culture appears differently between windows and linux/mono systems")
            System.Console.WriteLine("SPECIAL CASE HANDLING: month March in German culture expected as """ & ExpectedMarchShortNameOnCurrentPlatform & """")
            CompuMaster.Calendar.Month.Parse(ExpectedMarchShortNameOnCurrentPlatform, "MMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("de-DE"))
            Assert.AreEqual(True, CompuMaster.Calendar.Month.TryParse(ExpectedMarchShortNameOnCurrentPlatform, "MMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("de-DE"), Buffer), "#" & (TestNumber + 0.2).ToString)
            Assert.AreEqual("2010-03", Buffer.ToString, "#" & (TestNumber + 0.3).ToString)

            TestNumber = 200
            CompuMaster.Calendar.Month.Parse("Mar/2010", "UUU/yyyy", System.Globalization.CultureInfo.GetCultureInfo("de-DE"))
            Assert.AreEqual(True, CompuMaster.Calendar.Month.TryParse("Mar/2010", "UUU/yyyy", System.Globalization.CultureInfo.GetCultureInfo("de-DE"), Buffer), "#" & (TestNumber + 0.2).ToString)
            Assert.AreEqual("2010-03", Buffer.ToString, "#" & (TestNumber + 0.3).ToString)
            Assert.AreEqual("Mar.2010", Buffer.ToString("UUU/yyyy", System.Globalization.CultureInfo.GetCultureInfo("de-DE")), "#" & (TestNumber + 0.4).ToString)
            Assert.AreEqual("Mar/2010", Buffer.ToString("UUU\/yyyy", System.Globalization.CultureInfo.GetCultureInfo("de-DE")), "#" & (TestNumber + 0.4).ToString)
            Assert.AreEqual("Mar/2010", Buffer.ToString("UUU/yyyy", System.Globalization.CultureInfo.InvariantCulture), "#" & (TestNumber + 0.4).ToString)

            TestNumber = 210
            Select Case New CompuMaster.Calendar.Month(2020, 3).ToString("MMM", System.Globalization.CultureInfo.GetCultureInfo("de-DE"))
                Case "Mär" 'typically on platforms PlatformID.Unix, PlatformID.MacOSX + yet unclear: newer Windows versions or on .NET Core 
                    TestNumber = 211
                    Assert.AreEqual(True, CompuMaster.Calendar.Month.TryParse("Mär/2010", "MMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("de-DE"), Buffer), "#" & (TestNumber + 0.2).ToString)
                    Assert.AreEqual("2010-03", Buffer.ToString, "#" & (TestNumber + 0.3).ToString)
                    Assert.AreEqual("Mär.2010", Buffer.ToString("MMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("de-DE")), "#" & (TestNumber + 0.4).ToString)
                    Assert.AreEqual("Mär/2010", Buffer.ToString("MMM\/yyyy", System.Globalization.CultureInfo.GetCultureInfo("de-DE")), "#" & (TestNumber + 0.4).ToString)
                    Assert.AreEqual("Mar/2010", Buffer.ToString("UUU/yyyy", System.Globalization.CultureInfo.InvariantCulture), "#" & (TestNumber + 0.4).ToString)
                Case "Mrz" 'Windows platform
                    TestNumber = 212
                    Assert.AreEqual(True, CompuMaster.Calendar.Month.TryParse("Mrz/2010", "MMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("de-DE"), Buffer), "#" & (TestNumber + 0.2).ToString)
                    Assert.AreEqual("2010-03", Buffer.ToString, "#" & (TestNumber + 0.3).ToString)
                    Assert.AreEqual("Mrz.2010", Buffer.ToString("MMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("de-DE")), "#" & (TestNumber + 0.4).ToString)
                    Assert.AreEqual("Mrz/2010", Buffer.ToString("MMM\/yyyy", System.Globalization.CultureInfo.GetCultureInfo("de-DE")), "#" & (TestNumber + 0.4).ToString)
                    Assert.AreEqual("Mar/2010", Buffer.ToString("UUU/yyyy", System.Globalization.CultureInfo.InvariantCulture), "#" & (TestNumber + 0.4).ToString)
                Case Else
                    Throw New InvalidOperationException
            End Select

            TestNumber = 220
            Assert.Catch(Of ArgumentNullException)(Sub()
                                                       CompuMaster.Calendar.Month.Parse("Custom.März/2010",
                                                                                           "CCC/yyyy",
                                                                                           System.Globalization.CultureInfo.GetCultureInfo("de-DE"))
                                                   End Sub, "#" & (TestNumber + 0.1).ToString)
            Assert.Catch(Of ArgumentException)(Sub()
                                                   CompuMaster.Calendar.Month.Parse("Custom.März/2010",
                                                                                       "CCC/yyyy", System.Globalization.CultureInfo.GetCultureInfo("de-DE"),
                                                                                       New String() {"Custom.Jan", "Custom.Feb"})
                                               End Sub, "#" & (TestNumber + 0.2).ToString)

            TestNumber = 221
            Dim CustomMonths As String() = New String() {"Custom.Jan", "Custom.Feb", "Custom.März", "Custom.Apr", "Custom.Mai", "Custom.Jun", "Custom.Jul", "Custom.Aug", "Custom.Sept", "Custom.Oct", "Custom.Nov", "Custom.Dez"}
            Assert.AreEqual(True, CompuMaster.Calendar.Month.TryParse("Custom.März/2010", "CCC/yyyy", System.Globalization.CultureInfo.GetCultureInfo("de-DE"), CustomMonths, Buffer), "#" & (TestNumber + 0.2).ToString)
            Assert.AreEqual("2010-03", Buffer.ToString, "#" & (TestNumber + 0.3).ToString)
            Assert.AreEqual("Custom.März/2010", Buffer.ToString("CCC\/yyyy", System.Globalization.CultureInfo.GetCultureInfo("de-DE"), CustomMonths), "#" & (TestNumber + 0.4).ToString)
            Assert.AreEqual("Custom.März.2010", Buffer.ToString("CCC/yyyy", System.Globalization.CultureInfo.GetCultureInfo("de-DE"), CustomMonths), "#" & (TestNumber + 0.4).ToString)
            TestNumber = 222
            CompuMaster.Calendar.Month.Parse("13. Custom.März 2010 14:42:12", "dd. CCC yyyy HH:mm:ss", System.Globalization.CultureInfo.GetCultureInfo("de-DE"), CustomMonths)
            Assert.AreEqual(True, CompuMaster.Calendar.Month.TryParse("13. Custom.März 2010 14:42:12", "dd. CCC yyyy HH:mm:ss", System.Globalization.CultureInfo.GetCultureInfo("de-DE"), CustomMonths, Buffer), "#" & (TestNumber + 0.2).ToString)
            Assert.AreEqual("2010-03", Buffer.ToString, "#" & (TestNumber + 0.3).ToString)
            Assert.AreEqual("01. Custom.März 2010 00:00:00", Buffer.ToString("dd. CCC yyyy HH:mm:ss", System.Globalization.CultureInfo.GetCultureInfo("de-DE"), CustomMonths), "#" & (TestNumber + 0.4).ToString)

        End Sub

        <Test> Public Sub PreservingStringSplit()
            Dim Separators As String() = New String() {"CCC", "UUU"}
            Assert.AreEqual(New String() {}, CompuMaster.Calendar.Month.PreservingStringSplit(Nothing, Separators))
            Assert.AreEqual(New String() {}, CompuMaster.Calendar.Month.PreservingStringSplit("", Separators))
            Assert.AreEqual(New String() {"1. ", "CCC", " YYYY"}, CompuMaster.Calendar.Month.PreservingStringSplit("1. CCC YYYY", Separators))
            Assert.AreEqual(New String() {"1. ", "CCC"}, CompuMaster.Calendar.Month.PreservingStringSplit("1. CCC", Separators))
            Assert.AreEqual(New String() {"CCC", " YYYY"}, CompuMaster.Calendar.Month.PreservingStringSplit("CCC YYYY", Separators))
            Assert.AreEqual(New String() {"CCC"}, CompuMaster.Calendar.Month.PreservingStringSplit("CCC", Separators))
            Assert.AreEqual(New String() {"1. ", "CCC", " ", "UUU", " YYYY"}, CompuMaster.Calendar.Month.PreservingStringSplit("1. CCC UUU YYYY", Separators))
            Assert.AreEqual(New String() {"1. ", "CCC", "UUU", " YYYY"}, CompuMaster.Calendar.Month.PreservingStringSplit("1. CCCUUU YYYY", Separators))
        End Sub

        <Test> Public Sub Conversions()
            Assert.AreEqual("2020-10", (New CompuMaster.Calendar.Month(2020, 10).ToString))
            Assert.AreEqual(Nothing, CType(Nothing, CompuMaster.Calendar.Month)?.ToString)
            Assert.AreEqual(Nothing, CType(CType(Nothing, String), CompuMaster.Calendar.Month))
            Assert.AreEqual(Nothing, CType(String.Empty, CompuMaster.Calendar.Month))
            Assert.AreEqual("2020-10", CType("2020-10", CompuMaster.Calendar.Month).ToString)
        End Sub

        <Test> Public Sub ParseFromUniqueShortName()
            Assert.AreEqual("2010-10", CompuMaster.Calendar.Month.ParseFromUniqueShortName("Oct/2010").ToString)
            Assert.AreEqual("1900-01", CompuMaster.Calendar.Month.ParseFromUniqueShortName("Jan/1900").ToString)
            Assert.AreEqual("9999-12", CompuMaster.Calendar.Month.ParseFromUniqueShortName("Dec/9999").ToString)
        End Sub

        <Test> Public Sub UniqueShortName()
            Assert.AreEqual("Oct/2010", (New CompuMaster.Calendar.Month("2010-10")).UniqueShortName)
            Assert.AreEqual("Jan/1900", (New CompuMaster.Calendar.Month("1900-01")).UniqueShortName)
            Assert.AreEqual("Dec/9999", (New CompuMaster.Calendar.Month("9999-12")).UniqueShortName)
        End Sub

        <Test()> Public Sub OperatorMinus()
            Assert.AreEqual(0, New CompuMaster.Calendar.Month(2012, 5) - New CompuMaster.Calendar.Month(2012, 5))
            Assert.AreEqual(1, New CompuMaster.Calendar.Month(2012, 5) - New CompuMaster.Calendar.Month(2012, 4))
            Assert.AreEqual(2, New CompuMaster.Calendar.Month(2012, 5) - New CompuMaster.Calendar.Month(2012, 3))
            Assert.AreEqual(1, New CompuMaster.Calendar.Month(2012, 1) - New CompuMaster.Calendar.Month(2011, 12))
            Assert.AreEqual(3, New CompuMaster.Calendar.Month(2012, 2) - New CompuMaster.Calendar.Month(2011, 11))
            Assert.AreEqual(15, New CompuMaster.Calendar.Month(2012, 2) - New CompuMaster.Calendar.Month(2010, 11))
            Assert.AreEqual(21, New CompuMaster.Calendar.Month(2012, 11) - New CompuMaster.Calendar.Month(2011, 2))
            Assert.AreEqual(33, New CompuMaster.Calendar.Month(2012, 11) - New CompuMaster.Calendar.Month(2010, 2))
            Assert.AreEqual(-1, New CompuMaster.Calendar.Month(2012, 4) - New CompuMaster.Calendar.Month(2012, 5))
            Assert.AreEqual(-2, New CompuMaster.Calendar.Month(2012, 3) - New CompuMaster.Calendar.Month(2012, 5))
            Assert.AreEqual(-1, New CompuMaster.Calendar.Month(2011, 12) - New CompuMaster.Calendar.Month(2012, 1))
            Assert.AreEqual(-3, New CompuMaster.Calendar.Month(2011, 11) - New CompuMaster.Calendar.Month(2012, 2))
            Assert.AreEqual(-15, New CompuMaster.Calendar.Month(2010, 11) - New CompuMaster.Calendar.Month(2012, 2))
            Assert.AreEqual(12, New CompuMaster.Calendar.Month(2020, 1) - New CompuMaster.Calendar.Month(2019, 1))
            Assert.AreEqual(-12, New CompuMaster.Calendar.Month(2019, 1) - New CompuMaster.Calendar.Month(2020, 1))
            Assert.AreEqual(24, New CompuMaster.Calendar.Month(2021, 1) - New CompuMaster.Calendar.Month(2019, 1))
            Assert.AreEqual(-24, New CompuMaster.Calendar.Month(2018, 1) - New CompuMaster.Calendar.Month(2020, 1))
        End Sub

        <Test()> Public Sub Compares()
            Dim value1 As New CompuMaster.Calendar.Month(2010, 1)
            Dim value2 As New CompuMaster.Calendar.Month(2009, 12)
            Dim value3 As New CompuMaster.Calendar.Month(2010, 1)
            Assert.AreEqual(True, value1 > value2)
            Assert.AreEqual(False, value1 <value2)
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

        <Test> Public Sub Add()
            Dim value1 As New CompuMaster.Calendar.Month(2010, 1)
            Dim value2 As New CompuMaster.Calendar.Month(2007, 12)
            Dim value3 As New CompuMaster.Calendar.Month(2017, 12)
            Dim value4 As New CompuMaster.Calendar.Month(2011, 1)
            Dim value5 As New CompuMaster.Calendar.Month(2010, 5)
            Dim value6 As New CompuMaster.Calendar.Month(2009, 5)
            Assert.AreEqual(value4, value1.AddYears(1))
            Assert.AreEqual(value5, value1.AddMonths(4))
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
            Assert.AreEqual(New CompuMaster.Calendar.Month(2009, 12), value1.AddMonths(-1))
            Assert.AreEqual(New CompuMaster.Calendar.Month(2009, 2), value1.AddMonths(-11))
            Assert.AreEqual(New CompuMaster.Calendar.Month(2009, 1), value1.AddMonths(-12))
            Assert.AreEqual(New CompuMaster.Calendar.Month(2008, 12), value1.AddMonths(-13))
        End Sub

        <Test> Public Sub Min()
            Dim value0 As CompuMaster.Calendar.Month = Nothing
            Dim value0initialized As New CompuMaster.Calendar.Month
            Dim value1 As New CompuMaster.Calendar.Month(2010, 1)
            Dim value2 As New CompuMaster.Calendar.Month(2007, 12)
            Dim value3 As New CompuMaster.Calendar.Month(2017, 12)
            Dim value4 As New CompuMaster.Calendar.Month(2011, 1)
            Dim value5 As New CompuMaster.Calendar.Month(2010, 5)
            Dim value6 As New CompuMaster.Calendar.Month(2009, 5)
            Assert.AreEqual(value0, CompuMaster.Calendar.Month.min(value0, value0))
            Assert.AreEqual(value0, CompuMaster.Calendar.Month.min(value0, value1))
            Assert.AreEqual(value0, CompuMaster.Calendar.Month.min(value1, value0))
            Assert.AreEqual(value0, CompuMaster.Calendar.Month.Min(value0initialized, value0))
            Assert.AreEqual(value0, CompuMaster.Calendar.Month.Min(value0, value0initialized))
            Assert.AreEqual(value1, CompuMaster.Calendar.Month.Min(value1, value1))
            Assert.AreEqual(value2, CompuMaster.Calendar.Month.min(value1, value2))
            Assert.AreEqual(value1, CompuMaster.Calendar.Month.min(value1, value3))
            Assert.AreEqual(value1, CompuMaster.Calendar.Month.min(value1, value4))
            Assert.AreEqual(value1, CompuMaster.Calendar.Month.min(value1, value5))
            Assert.AreEqual(value1, CompuMaster.Calendar.Month.min(value5, value1))
            Assert.AreEqual(value6, CompuMaster.Calendar.Month.min(value1, value6))
        End Sub

        <Test> Public Sub Max()
            Dim value0 As CompuMaster.Calendar.Month = Nothing
            Dim value0initialized As New CompuMaster.Calendar.Month
            Dim value1 As New CompuMaster.Calendar.Month(2010, 1)
            Dim value2 As New CompuMaster.Calendar.Month(2007, 12)
            Dim value3 As New CompuMaster.Calendar.Month(2017, 12)
            Dim value4 As New CompuMaster.Calendar.Month(2011, 1)
            Dim value5 As New CompuMaster.Calendar.Month(2010, 5)
            Dim value6 As New CompuMaster.Calendar.Month(2009, 5)
            Assert.AreEqual(value0, CompuMaster.Calendar.Month.max(value0, value0))
            Assert.AreEqual(value1, CompuMaster.Calendar.Month.Max(value0, value1))
            Assert.AreEqual(value1, CompuMaster.Calendar.Month.Max(value1, value0))
            Assert.AreEqual(value0initialized, CompuMaster.Calendar.Month.Max(value0initialized, value0))
            Assert.AreEqual(value0initialized, CompuMaster.Calendar.Month.Max(value0, value0initialized))
            Assert.AreEqual(value1, CompuMaster.Calendar.Month.Max(value1, value1))
            Assert.AreEqual(value1, CompuMaster.Calendar.Month.Max(value1, value2))
            Assert.AreEqual(value3, CompuMaster.Calendar.Month.Max(value1, value3))
            Assert.AreEqual(value4, CompuMaster.Calendar.Month.Max(value1, value4))
            Assert.AreEqual(value5, CompuMaster.Calendar.Month.Max(value1, value5))
            Assert.AreEqual(value5, CompuMaster.Calendar.Month.Max(value5, value1))
            Assert.AreEqual(value1, CompuMaster.Calendar.Month.Max(value1, value6))
        End Sub

        <Test> Public Sub SmallerBiggerEquals()
            Assert.Multiple(Sub()
                                Assert.IsTrue(New CompuMaster.Calendar.Month(2010, 9) = New CompuMaster.Calendar.Month(2010, 9))
                                Assert.IsTrue(New CompuMaster.Calendar.Month(2010, 1) <> New CompuMaster.Calendar.Month(2010, 9))
                                Assert.IsTrue(New CompuMaster.Calendar.Month(2010, 1) < New CompuMaster.Calendar.Month(2010, 9))
                                Assert.IsTrue(New CompuMaster.Calendar.Month(2010, 9) > New CompuMaster.Calendar.Month(2010, 1))
                                Assert.IsFalse(New CompuMaster.Calendar.Month(2010, 9) <> New CompuMaster.Calendar.Month(2010, 9))
                                Assert.IsFalse(New CompuMaster.Calendar.Month(2010, 1) = New CompuMaster.Calendar.Month(2010, 9))
                                Assert.IsFalse(New CompuMaster.Calendar.Month(2010, 1) > New CompuMaster.Calendar.Month(2010, 9))
                                Assert.IsFalse(New CompuMaster.Calendar.Month(2010, 9) < New CompuMaster.Calendar.Month(2010, 1))

                                Assert.AreEqual(New CompuMaster.Calendar.Month(2010, 9), New CompuMaster.Calendar.Month(2010, 9))
                                Assert.AreNotEqual(New CompuMaster.Calendar.Month(2010, 1), New CompuMaster.Calendar.Month(2010, 9))
                                Assert.Less(New CompuMaster.Calendar.Month(2010, 1), New CompuMaster.Calendar.Month(2010, 9))
                                Assert.Greater(New CompuMaster.Calendar.Month(2010, 9), New CompuMaster.Calendar.Month(2010, 1))
                                Assert.That(New CompuMaster.Calendar.Month(2010, 1), NUnit.Framework.Is.LessThan(New CompuMaster.Calendar.Month(2010, 9)))
                                Assert.That(New CompuMaster.Calendar.Month(2010, 9), NUnit.Framework.Is.GreaterThan(New CompuMaster.Calendar.Month(2010, 1)))
                            End Sub)
        End Sub

        <Test> Public Sub Sorting()
            Dim Values As New Generic.List(Of CompuMaster.Calendar.Month)
            Values.Add(New CompuMaster.Calendar.Month(2010, 11))
            Values.Add(New CompuMaster.Calendar.Month(1999, 10))
            Values.Add(New CompuMaster.Calendar.Month(2010, 10))
            Values.Add(New CompuMaster.Calendar.Month)
            Values.Add(New CompuMaster.Calendar.Month(2010, 9))

            Values.Sort()

            Assert.AreEqual(New CompuMaster.Calendar.Month, Values(0))
            Assert.AreEqual(New CompuMaster.Calendar.Month(1999, 10), Values(1))
            Assert.AreEqual(New CompuMaster.Calendar.Month(2010, 9), Values(2))
            Assert.AreEqual(New CompuMaster.Calendar.Month(2010, 10), Values(3))
            Assert.AreEqual(New CompuMaster.Calendar.Month(2010, 11), Values(4))
        End Sub

    End Class

End Namespace
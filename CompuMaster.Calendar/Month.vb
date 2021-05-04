﻿Option Explicit On 
Option Strict On

Namespace CompuMaster.Calendar

    ''' <summary>
    ''' A representation for a month period for years with 12 months
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Month
        Implements IComparable

        Public Sub New()
            Me.Year = 1
            Me.Month = 1
        End Sub
        Public Sub New(ByVal value As DateTime)
            Me.Year = value.Year
            Me.Month = value.Month
        End Sub
        Public Sub New(ByVal year As Integer, ByVal month As Integer)
            Me.Year = year
            Me.Month = month
        End Sub
        Public Sub New(ByVal value As String)
            If value = Nothing OrElse value.Length <> 7 OrElse value.Substring(4, 1) <> "-" Then
                Throw New ArgumentException("value must be formatted as yyyy-MM")
            End If
            Me.Year = Integer.Parse(value.Substring(0, 4))
            Me.Month = Integer.Parse(value.Substring(5, 2))
        End Sub

        Public Shared ReadOnly MinValue As Month = New Month(1, 1)
        Public Shared ReadOnly MaxValue As Month = New Month(9999, 12)

        ''' <summary>
        ''' The current month
        ''' </summary>
        ''' <returns></returns>
        Public Shared Function Now() As Month
            Return New Month(DateTime.Now)
        End Function

        Public Shared Function Parse(value As Integer) As Month
            If value < 0 Then Throw New InvalidCastException("Value must be a positive number")
            Dim Year As Integer = CType(value / 100, Integer)
            If Year > 9999 Then Throw New InvalidCastException("Year must be <= 9999")
            Dim Month As Integer = value - Year * 100
            If Month > 12 Then Throw New InvalidCastException("Month must be <= 12")
            Return New Month(Year, Month)
        End Function

        Public Shared Function Parse(value As Long) As Month
            Return Parse(CType(value, Integer))
        End Function

        'Public Shared Function Parse(value As UInteger) As Month
        '    Return Parse(CType(value, Integer))
        'End Function
        '
        'Public Shared Function Parse(value As ULong) As Month
        '    Return Parse(CType(value, Integer))
        'End Function

        Public Shared Function Parse(value As String) As Month
            If value = Nothing Then
                Return Nothing
            ElseIf value.Length <> 7 OrElse value.Substring(4, 1) <> "-" Then
                Throw New ArgumentException("value must be formatted as yyyy-MM")
            Else
                Return New Month(Integer.Parse(value.Substring(0, 4)), Integer.Parse(value.Substring(5, 2)))
            End If
        End Function

        Private _Year As Integer
        ''' <summary>
        ''' The year of the month
        ''' </summary>
        Public Property Year() As Integer
            Get
                Return _Year
            End Get
            Set(ByVal value As Integer)
                If value <= 0 Then Throw New ArgumentOutOfRangeException(NameOf(value), "Year must be a positive number")
                _Year = value
            End Set
        End Property

        Private _Month As Integer
        ''' <summary>
        ''' The number of the month (Jan = 1, Dec = 12)
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        Public Property Month() As Integer
            Get
                Return _Month
            End Get
            Set(ByVal value As Integer)
                If value < 0 OrElse value > 12 Then Throw New ArgumentOutOfRangeException(NameOf(value), "Invalid value " & value.ToString)
                _Month = value
            End Set
        End Property

        Public Function Add(years As Integer, months As Integer) As CompuMaster.Calendar.Month
            Return Me.AddYears(years).AddMonths(months)
        End Function
        Public Function AddYears(value As Integer) As CompuMaster.Calendar.Month
            Return New CompuMaster.Calendar.Month(Me.Year + value, Me.Month)
        End Function
        Public Function AddMonths(value As Integer) As CompuMaster.Calendar.Month
            Dim NewMonthValue As Integer = Me.Month + value
            Dim NewYearValue As Integer = Me.Year
            If NewMonthValue > 12 Then
                Dim AddYears As Integer = NewMonthValue \ 12
                NewMonthValue -= AddYears * 12
                NewYearValue += AddYears
            End If
            If NewMonthValue < 1 Then
                Dim MonthsToGoBack As Integer = 1 - NewMonthValue
                Dim SubstractYears As Integer = MonthsToGoBack \ 12
                NewMonthValue += (SubstractYears + 1) * 12
                NewYearValue -= SubstractYears + 1
                If NewMonthValue = 13 Then
                    NewMonthValue = 1
                    NewYearValue += 1
                End If
            End If
            Return New CompuMaster.Calendar.Month(NewYearValue, NewMonthValue)
        End Function

        ''' <summary>
        ''' Text formatting with yyyy-MM
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Overrides Function ToString() As String
            Return Year.ToString("0000") & "-" & Month.ToString("00")
        End Function

        ''' <summary>
        ''' Format the month with a typical datetime format string using the begin date of the period
        ''' </summary>
        ''' <param name="format">YYYY for 4-digit year, YY for 2-digit year, MMMM for long month name, MMM for month name abbreviation, MM for always-2-digit month number, M for month number with 1 or 2 digits, UUU for unique short name of month (might depend on system platform), CCC for a custom set of expected names from January to December</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        Public Overloads Function ToString(ByVal format As String) As String
            Return Me.BeginOfPeriod.ToString(format)
        End Function
        ''' <summary>
        ''' Format the month with a typical datetime format using the given format provider and using the begin date of the period
        ''' </summary>
        ''' <param name="provider"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        Public Overloads Function ToString(ByVal provider As System.IFormatProvider) As String
            Return Me.BeginOfPeriod.ToString(provider)
        End Function

        ''' <summary>
        ''' Format the month with a typical datetime format string and the given format provider using the begin date of the period
        ''' </summary>
        ''' <param name="format">YYYY for 4-digit year, YY for 2-digit year, MMMM for long month name, MMM for month name abbreviation, MM for always-2-digit month number, M for month number with 1 or 2 digits, UUU for unique short name of month (might depend on system platform), CCC for a custom set of expected names from January to December</param>
        ''' <param name="provider"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        Public Overloads Function ToString(ByVal format As String, ByVal provider As System.IFormatProvider) As String
            Return Me.ToString(format, provider, CType(Nothing, String()))
        End Function

        ''' <summary>
        ''' Format the month with a typical datetime format string and the given format provider using the begin date of the period
        ''' </summary>
        ''' <param name="format">YYYY for 4-digit year, YY for 2-digit year, MMMM for long month name, MMM for month name abbreviation, MM for always-2-digit month number, M for month number with 1 or 2 digits, UUU for unique short name of month (might depend on system platform), CCC for a custom set of expected names from January to December</param>
        ''' <param name="provider"></param>
        ''' <param name="customMonths">An array of 12 strings reprensenting the expected month names, starting with January up to December</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        Public Overloads Function ToString(ByVal format As String, ByVal provider As System.IFormatProvider, customMonths As String()) As String
            If format.Contains("CCC") Then
                If customMonths Is Nothing Then Throw New ArgumentNullException(NameOf(customMonths))
                If customMonths.Length <> 12 Then Throw New ArgumentException("Array with 12 elements required", NameOf(customMonths))
            End If
            If format.Contains("CCC") OrElse format.Contains("UUU") Then
                Dim FormatSplitted As String() = PreservingStringSplit(format, New String() {"CCC", "UUU"})
                For MyCounter As Integer = 0 To FormatSplitted.Length - 1
                    If FormatSplitted(MyCounter) = "CCC" Then
                        FormatSplitted(MyCounter) = customMonths(Me.Month - 1)
                    ElseIf FormatSplitted(MyCounter) = "UUU" Then
                        FormatSplitted(MyCounter) = UniqueMonthShortName(Me.Month)
                    Else
                        FormatSplitted(MyCounter) = Me.BeginOfPeriod.ToString(FormatSplitted(MyCounter), provider)
                    End If
                Next
                Return String.Join("", FormatSplitted)
            Else
                Return Me.BeginOfPeriod.ToString(format, provider)
            End If
        End Function

        ''' <summary>
        ''' Split a string but keep the separators as items
        ''' </summary>
        ''' <param name="value"></param>
        ''' <param name="separators"></param>
        ''' <returns></returns>
        Friend Shared Function PreservingStringSplit(value As String, separators As String()) As String()
            If value Is Nothing OrElse separators Is Nothing OrElse separators.Length = 0 Then Return New String() {}
            Dim ValueToProcess As String = value
            Dim Result As New System.Collections.Generic.List(Of String)
            Do
                'Check for a separator at first position
                Dim SeparatorAlreadyFound As Boolean = False
                For SeparatorCounter As Integer = 0 To separators.Length - 1
                    If ValueToProcess.StartsWith(separators(SeparatorCounter)) Then
                        Result.Add(separators(SeparatorCounter))
                        ValueToProcess = ValueToProcess.Substring(separators(SeparatorCounter).Length)
                        SeparatorAlreadyFound = True
                    End If
                Next
                'Check for next separator occurance if all separators at index 0 have been cut repetitive
                If SeparatorAlreadyFound = False Then
                    Dim NextSeparator As Integer = ValueToProcess.Length
                    For SeparatorCounter As Integer = 0 To separators.Length - 1
                        Dim ValidNextSeparator As Integer = ValueToProcess.IndexOf(separators(SeparatorCounter))
                        If ValidNextSeparator <> -1 Then NextSeparator = System.Math.Min(NextSeparator, ValidNextSeparator)
                    Next
                    If NextSeparator <= 0 Then
                        Exit Do
                    Else
                        Result.Add(ValueToProcess.Substring(0, NextSeparator))
                        ValueToProcess = ValueToProcess.Substring(NextSeparator)
                    End If
                End If
            Loop
            Return Result.ToArray
        End Function

        ''' <summary>
        ''' Parse the month from a defined datetime format string and the given culture
        ''' </summary>
        ''' <param name="format">YYYY for 4-digit year, YY for 2-digit year, MMMM for long month name, MMM for month name abbreviation, MM for always-2-digit month number, M for month number with 1 or 2 digits, UUU for unique short name of month (might depend on system platform), CCC for a custom set of expected names from January to December</param>
        ''' <param name="culture"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        Public Shared Function Parse(value As String, format As String, culture As System.Globalization.CultureInfo) As Month
            Return Parse(value, format, culture, CType(Nothing, String()))
        End Function

        ''' <summary>
        ''' Parse the month from a defined datetime format string and the given culture
        ''' </summary>
        ''' <param name="format">YYYY for 4-digit year, YY for 2-digit year, MMMM for long month name, MMM for month name abbreviation, MM for always-2-digit month number, M for month number with 1 or 2 digits, UUU for unique short name of month (might depend on system platform), CCC for a custom set of expected names from January to December</param>
        ''' <param name="culture"></param>
        ''' <param name="customMonths">An array of 12 strings reprensenting the expected month names, starting with January up to December</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        Public Shared Function Parse(value As String, format As String, culture As System.Globalization.CultureInfo, customMonths As String()) As Month
            If value = Nothing Then
                Throw New ArgumentNullException(NameOf(value))
            End If
            Dim MonthShortNames As New System.Collections.Generic.List(Of String)
            Dim MonthLongNames As New System.Collections.Generic.List(Of String)
            For MyCounter As Integer = 1 To 12
                MonthShortNames.Add(New Month(2000, MyCounter).MonthShortName(culture))
                MonthLongNames.Add(New Month(2000, MyCounter).MonthName(culture))
            Next
            Dim Pattern As String = System.Text.RegularExpressions.Regex.Escape(format)
            If Pattern.Contains("dd") Then 'won't be used, but must be present to allow parsing of date strings with day information
                Pattern = Pattern.Replace("dd", "(?<d>\d\d)")
            ElseIf Pattern.Contains("d") Then
                Pattern = Pattern.Replace("d", "(?<d>\d?\d)")
            End If
            If Pattern.Contains("HH") Then 'won't be used, but must be present to allow parsing of date strings with day information
                Pattern = Pattern.Replace("HH", "(?<d>\d\d)")
            ElseIf Pattern.Contains("H") Then
                Pattern = Pattern.Replace("H", "(?<d>\d?\d)")
            End If
            If Pattern.Contains("mm") Then 'won't be used, but must be present to allow parsing of date strings with day information
                Pattern = Pattern.Replace("mm", "(?<d>\d\d)")
            ElseIf Pattern.Contains("m") Then
                Pattern = Pattern.Replace("m", "(?<d>\d?\d)")
            End If
            If Pattern.Contains("ss") Then 'won't be used, but must be present to allow parsing of date strings with day information
                Pattern = Pattern.Replace("ss", "(?<d>\d\d)")
            ElseIf Pattern.Contains("s") Then
                Pattern = Pattern.Replace("s", "(?<d>\d?\d)")
            End If
            If Pattern.Contains("ffff") Then 'won't be used, but must be present to allow parsing of date strings with day information
                Pattern = Pattern.Replace("ffff", "(?<d>\d\d\d\d)")
            ElseIf Pattern.Contains("fff") Then 'won't be used, but must be present to allow parsing of date strings with day information
                Pattern = Pattern.Replace("fff", "(?<d>\d\d\d)")
            ElseIf Pattern.Contains("ff") Then 'won't be used, but must be present to allow parsing of date strings with day information
                Pattern = Pattern.Replace("ff", "(?<d>\d\d)")
            ElseIf Pattern.Contains("f") Then
                Pattern = Pattern.Replace("f", "(?<d>\d?\d)")
            End If
            If Pattern.Contains("zzz") Then 'won't be used, but must be present to allow parsing of date strings with day information
                Pattern = Pattern.Replace("zzz", "(?<d>[\-|\+|\ ]\d\d\:\d\d)")
            End If
            Pattern = Pattern.Replace("yyyy", "(?<year4>\d\d\d\d)").Replace("yy", "(?<year2>\d\d)")
            If Pattern.Contains("CCC") Then
                If customMonths Is Nothing Then Throw New ArgumentNullException(NameOf(customMonths))
                If customMonths.Length <> 12 Then Throw New ArgumentException("Array with 12 elements required", NameOf(customMonths))
                MonthShortNames = New Generic.List(Of String)(customMonths)
                MonthLongNames = New Generic.List(Of String)(customMonths)
                Pattern = Pattern.Replace("CCC", "(?<m>" & String.Join("|", EncodeForRegEx(customMonths)) & ")")
            ElseIf Pattern.Contains("UUU") Then
                MonthShortNames = New Generic.List(Of String)(UniqueMonthShortNames)
                MonthLongNames = New Generic.List(Of String)(UniqueMonthShortNames)
                Pattern = Pattern.Replace("UUU", "(?<m>" & String.Join("|", EncodeForRegEx(UniqueMonthShortNames)) & ")")
            ElseIf Pattern.Contains("MMMM") Then
                Pattern = Pattern.Replace("MMMM", "(?<m>" & String.Join("|", EncodeForRegEx(MonthLongNames.ToArray)) & ")")
            ElseIf Pattern.Contains("MMM") Then
                Pattern = Pattern.Replace("MMM", "(?<m>" & String.Join("|", EncodeForRegEx(MonthShortNames.ToArray)) & ")")
            ElseIf Pattern.Contains("MM") Then
                Pattern = Pattern.Replace("MM", "(?<m>\d\d)")
            ElseIf Pattern.Contains("M") Then
                Pattern = Pattern.Replace("M", "(?<m>\d?\d)")
            End If
            Dim RegEx As New System.Text.RegularExpressions.Regex(Pattern, Text.RegularExpressions.RegexOptions.Compiled Or Text.RegularExpressions.RegexOptions.Singleline Or Text.RegularExpressions.RegexOptions.Multiline)
            If RegEx.IsMatch(value) = False Then
                Throw New ArgumentException("Invalid value", NameOf(value))
            End If
            Dim RegMatch As System.Text.RegularExpressions.Match = RegEx.Match(value)
            Dim GroupNames As New System.Collections.Generic.List(Of String)(RegEx.GetGroupNames())
            If RegMatch.Groups.Count >= 2 Then
                Dim FoundYear As Integer
                If GroupNames.Contains("year4") Then
                    FoundYear = Integer.Parse(RegMatch.Groups("year4").Value)
                ElseIf GroupNames.Contains("year2") Then
                    FoundYear = culture.Calendar.ToFourDigitYear(Integer.Parse(RegMatch.Groups("year2").Value))
                Else
                    Throw New ArgumentException("Invalid value", NameOf(value))
                End If
                Dim FoundMonthName As String = RegMatch.Groups("m").Value
                Dim FoundMonth As Integer = 0
                For MyCounter As Integer = 1 To 12
                    If MonthLongNames(MyCounter - 1) = FoundMonthName OrElse MonthShortNames(MyCounter - 1) = FoundMonthName OrElse MyCounter.ToString("00") = FoundMonthName OrElse MyCounter.ToString("0") = FoundMonthName Then
                        FoundMonth = MyCounter
                        Exit For
                    End If
                Next
                If FoundYear = 0 OrElse FoundMonth = 0 Then
                    Throw New ArgumentException("Invalid value", NameOf(value))
                Else
                    Return New Month(FoundYear, FoundMonth)
                End If
            Else
                Throw New ArgumentException("Invalid value", NameOf(value))
            End If
            Throw New NotImplementedException
        End Function

        ''' <summary>
        ''' Try to parse the month from datetime format YYYY-MM
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        Public Shared Function TryParse(value As String, ByRef result As Month) As Boolean
            Try
                result = Parse(value)
                Return True
#Disable Warning CA1031 ' Do not catch general exception types
            Catch
                Return False
#Enable Warning CA1031 ' Do not catch general exception types
            End Try
        End Function

        ''' <summary>
        ''' Try to parse the month from a defined datetime format string and the given culture
        ''' </summary>
        ''' <param name="format">YYYY for 4-digit year, YY for 2-digit year, MMMM for long month name, MMM for month name abbreviation, MM for always-2-digit month number, M for month number with 1 or 2 digits, UUU for unique short name of month (might depend on system platform), CCC for a custom set of expected names from January to December</param>
        ''' <param name="culture"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        Public Shared Function TryParse(value As String, format As String, culture As System.Globalization.CultureInfo, ByRef result As Month) As Boolean
            Try
                result = Parse(value, format, culture)
                Return True
#Disable Warning CA1031 ' Do not catch general exception types
            Catch
                Return False
#Enable Warning CA1031 ' Do not catch general exception types
            End Try
        End Function

        ''' <summary>
        ''' Try to parse the month from a defined datetime format string and the given culture
        ''' </summary>
        ''' <param name="format">YYYY for 4-digit year, YY for 2-digit year, MMMM for long month name, MMM for month name abbreviation, MM for always-2-digit month number, M for month number with 1 or 2 digits, UUU for unique short name of month (might depend on system platform), CCC for a custom set of expected names from January to December</param>
        ''' <param name="culture"></param>
        ''' <param name="customMonths">An array of 12 strings reprensenting the expected month names, starting with January up to December</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        Public Shared Function TryParse(value As String, format As String, culture As System.Globalization.CultureInfo, customMonths As String(), ByRef result As Month) As Boolean
            Try
                result = Parse(value, format, culture, customMonths)
                Return True
#Disable Warning CA1031 ' Do not catch general exception types
            Catch
                Return False
#Enable Warning CA1031 ' Do not catch general exception types
            End Try
        End Function

        ''' <summary>
        ''' Parse the month from a format string UUU/YYYY
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        Public Shared Function ParseFromUniqueShortName(value As String) As Month
            Dim Pattern As String = "^(?<m>" & String.Join("|", EncodeForRegEx(UniqueMonthShortNames)) & ")\/(?<y>\d\d\d\d)$"
            Dim RegEx As New System.Text.RegularExpressions.Regex(Pattern, Text.RegularExpressions.RegexOptions.Compiled Or Text.RegularExpressions.RegexOptions.Singleline Or Text.RegularExpressions.RegexOptions.Multiline)
            If RegEx.IsMatch(value) = False Then
                Throw New ArgumentException("Invalid value", NameOf(value))
            End If
            Dim RegMatch As System.Text.RegularExpressions.Match = RegEx.Match(value)
            Dim GroupNames As String() = RegEx.GetGroupNames()
            If RegMatch.Groups.Count >= 2 Then
                Dim FoundYear As Integer = Integer.Parse(RegMatch.Groups("y").Value)
                Dim FoundMonthName As String = RegMatch.Groups("m").Value
                Dim FoundMonth As Integer = 0
                For MyCounter As Integer = 1 To 12
                    If UniqueMonthShortNames(MyCounter - 1) = FoundMonthName Then
                        FoundMonth = MyCounter
                        Exit For
                    End If
                Next
                If FoundYear = 0 OrElse FoundMonth = 0 Then
                    Throw New ArgumentException("Invalid value", NameOf(value))
                Else
                    Return New Month(FoundYear, FoundMonth)
                End If
            Else
                Throw New ArgumentException("Invalid value", NameOf(value))
            End If
        End Function

        ''' <summary>
        ''' Encode all array elements for RegEx
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Friend Shared Function EncodeForRegEx(items As String()) As String()
            Dim Result As New System.Collections.Generic.List(Of String)
            For MyCounter As Integer = 0 To items.Length - 1
                Result.Add(System.Text.RegularExpressions.Regex.Escape(items(MyCounter)))
            Next
            Return Result.ToArray
        End Function

        ''' <summary>
        ''' All short names in format UUU where UUU equals MMM (English names)
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared Function UniqueMonthShortNames() As String()
            Dim MonthNames As New System.Collections.Generic.List(Of String)
            For MyCounter As Integer = 1 To 12
                MonthNames.Add(UniqueMonthShortName(MyCounter))
            Next
            Return MonthNames.ToArray
        End Function

        ''' <summary>
        ''' A short name in format UUU where UUU equals MMM (English names)
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared Function UniqueMonthShortName(monthNo As Integer) As String
            Dim Result As String
            Select Case monthNo
                Case 1
                    Result = "Jan"
                Case 2
                    Result = "Feb"
                Case 3
                    Result = "Mar"
                Case 4
                    Result = "Apr"
                Case 5
                    Result = "May"
                Case 6
                    Result = "Jun"
                Case 7
                    Result = "Jul"
                Case 8
                    Result = "Aug"
                Case 9
                    Result = "Sep"
                Case 10
                    Result = "Oct"
                Case 11
                    Result = "Nov"
                Case 12
                    Result = "Dec"
                Case Else
                    Throw New ArgumentOutOfRangeException(NameOf(monthNo))
            End Select
            Return Result
        End Function

        ''' <summary>
        ''' A short name in format UUU/YYYY where UUU equals MMM (English names)
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function UniqueShortName() As String
            Return UniqueMonthShortName(Me.Month) & "/" & Me.Year.ToString("0000")
        End Function

        ''' <summary>
        ''' A short name in format MMM/YYYY of specified culture
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function MonthShortName(cultureName As String) As String
            Return Me.MonthShortName(System.Globalization.CultureInfo.GetCultureInfo(cultureName))
        End Function

        ''' <summary>
        ''' A short name in format MMM/YYYY of specified culture
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function MonthShortName(culture As System.Globalization.CultureInfo) As String
            Return Me.BeginOfPeriod.ToString("MMM", culture.DateTimeFormat)
        End Function

        ''' <summary>
        ''' A full month name in format MMMM/YYYY of specified culture
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function MonthName(cultureName As String) As String
            Return Me.MonthName(System.Globalization.CultureInfo.GetCultureInfo(cultureName))
        End Function

        ''' <summary>
        ''' A full month name in format MMMM/YYYY of specified culture
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function MonthName(culture As System.Globalization.CultureInfo) As String
            Return Me.BeginOfPeriod.ToString("MMMM", culture.DateTimeFormat)
        End Function

#If Not NET_1_1 Then
        ''' <summary>
        ''' Equals operator for Month classes
        ''' </summary>
        ''' <param name="value1"></param>
        ''' <param name="value2"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Operator =(ByVal value1 As Month, ByVal value2 As Month) As Boolean
            If value1 Is Nothing And value2 Is Nothing Then
                Return True
            ElseIf value1 Is Nothing And value2 IsNot Nothing Then
                Return False
            Else
                Return value1.Equals(value2)
            End If
        End Operator

        ''' <summary>
        ''' Not-Equals operator for Month classes
        ''' </summary>
        ''' <param name="value1"></param>
        ''' <param name="value2"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Operator <>(ByVal value1 As Month, ByVal value2 As Month) As Boolean
            If value1 Is Nothing And value2 Is Nothing Then
                Return False
            ElseIf value1 Is Nothing And value2 IsNot Nothing Then
                Return True
            Else
                Return Not value1.Equals(value2)
            End If
        End Operator

        ''' <summary>
        ''' Smaller than operator for Month classes
        ''' </summary>
        ''' <param name="value1"></param>
        ''' <param name="value2"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Operator <(ByVal value1 As Month, ByVal value2 As Month) As Boolean
            If value1 Is Nothing And value2 Is Nothing Then
                Return False
            ElseIf value1 Is Nothing And value2 IsNot Nothing Then
                Return True
            Else
                Return value1.CompareTo(value2) < 0
            End If
        End Operator

        ''' <summary>
        ''' Greater than operator for Month classes
        ''' </summary>
        ''' <param name="value1"></param>
        ''' <param name="value2"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Operator >(ByVal value1 As Month, ByVal value2 As Month) As Boolean
            If value1 Is Nothing And value2 Is Nothing Then
                Return False
            ElseIf value1 Is Nothing And value2 IsNot Nothing Then
                Return False
            Else
                Return value1.CompareTo(value2) > 0
            End If
        End Operator

        ''' <summary>
        ''' Smaller than operator for Month classes
        ''' </summary>
        ''' <param name="value1"></param>
        ''' <param name="value2"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Operator <=(ByVal value1 As Month, ByVal value2 As Month) As Boolean
            If value1 Is Nothing And value2 Is Nothing Then
                Return True
            ElseIf value1 Is Nothing And value2 IsNot Nothing Then
                Return True
            Else
                Return value1.CompareTo(value2) <= 0
            End If
        End Operator

        ''' <summary>
        ''' Greater than operator for Month classes
        ''' </summary>
        ''' <param name="value1"></param>
        ''' <param name="value2"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Operator >=(ByVal value1 As Month, ByVal value2 As Month) As Boolean
            If value1 Is Nothing And value2 Is Nothing Then
                Return True
            ElseIf value1 Is Nothing And value2 IsNot Nothing Then
                Return False
            Else
                Return value1.CompareTo(value2) >= 0
            End If
        End Operator

        ''' <summary>
        ''' The number of total months between two values
        ''' </summary>
        ''' <param name="value1"></param>
        ''' <param name="value2"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Operator -(value1 As Month, value2 As Month) As Integer
            If value1 Is Nothing Then Throw New ArgumentNullException(NameOf(value1))
            If value2 Is Nothing Then Throw New ArgumentNullException(NameOf(value2))
            Dim SwappedValues As Boolean, StartMonth As Month, EndMonth As Month
            If value1 = value2 Then
                Return 0
            ElseIf value2 > value1 Then
                EndMonth = value2
                StartMonth = value1
                SwappedValues = True
            Else
                EndMonth = value1
                StartMonth = value2
                SwappedValues = False
            End If
            Dim Result As Integer
            If EndMonth.Year = StartMonth.Year Then
                Result = EndMonth.Month - StartMonth.Month
            Else 'If EndMonth.Year - StartMonth.Year>=1
                Result = EndMonth.Month + (12 - StartMonth.Month) + (EndMonth.Year - StartMonth.Year - 1) * 12
            End If
            If SwappedValues Then Result *= -1
            Return Result
        End Operator
#End If

        ''' <summary>
        ''' Create an instance of the following period
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function [NextPeriod]() As Month
            Return New Month(BeginOfPeriod.AddMonths(1))
        End Function

        ''' <summary>
        ''' Create an instance of the previous period
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function [PreviousPeriod]() As Month
            Return New Month(BeginOfPeriod.AddMonths(-1))
        End Function

        ''' <summary>
        ''' The begin of the month
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function BeginOfPeriod() As DateTime
            Return New DateTime(Year, Month, 1)
        End Function

        ''' <summary>
        ''' The end of the month
        ''' </summary>
        ''' <param name="precision"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function EndOfPeriod(ByVal precision As CompuMaster.Calendar.DateInformation.Accuracy) As DateTime
            Return CompuMaster.Calendar.DateInformation.EndOfMonth(BeginOfPeriod, precision)
        End Function

        ''' <summary>
        ''' The begin of the year
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function BeginOfYear() As CompuMaster.Calendar.Month
            Return New CompuMaster.Calendar.Month(Year, 1)
        End Function

        ''' <summary>
        ''' The end of the year
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function EndOfYear() As CompuMaster.Calendar.Month
            Return New CompuMaster.Calendar.Month(Year, 12)
        End Function

        ''' <summary>
        ''' Compares a value to the current instance value
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function CompareTo(ByVal obj As Object) As Integer Implements System.IComparable.CompareTo
            Return Me.CompareTo(CType(obj, Month))
        End Function

        ''' <summary>
        ''' Compares a value to the current instance value
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns>-1 if value is smaller, +1 if value is greater, 0 if value equals</returns>
        ''' <remarks></remarks>
        Public Function CompareTo(ByVal value As Month) As Integer
            If value Is Nothing Then
                Return 1
            ElseIf Me.BeginOfPeriod < value.BeginOfPeriod Then
                Return -1
            ElseIf Me.BeginOfPeriod > value.BeginOfPeriod Then
                Return 1
            Else
                Return 0
            End If
        End Function

        ''' <summary>
        ''' Equals method for period classes
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Function Equals(ByVal value As Month) As Boolean
            If value IsNot Nothing Then
                If Me.Year = value.Year AndAlso Me.Month = value.Month Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        End Function

        ''' <summary>
        ''' Equals method for period classes
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        Public Overrides Function Equals(obj As Object) As Boolean
            If ReferenceEquals(Me, obj) Then
                Return True
            End If

            If obj Is Nothing Then
                Return False
            End If

            If GetType(Month).IsInstanceOfType(obj) = False Then
                Return False
            End If

            Return Me.Year = CType(obj, Month).Year AndAlso Me.Month = CType(obj, Month).Month
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return Me.Year * 100 + Me.Month
        End Function

        Public Shared Function Min(value1 As Month, value2 As Month) As Month
            If value1 < value2 Then
                Return value1
            Else
                Return value2
            End If
        End Function

        Public Shared Function Max(value1 As Month, value2 As Month) As Month
            If value1 > value2 Then
                Return value1
            Else
                Return value2
            End If
        End Function

        Public Shared Narrowing Operator CType(value As CompuMaster.Calendar.ZeroableMonth) As CompuMaster.Calendar.Month
            If value Is Nothing Then
                Return Nothing
            ElseIf value.Month >= 1 AndAlso value.Month <= 12 Then
                Return New CompuMaster.Calendar.Month(value.Year, value.Month)
            ElseIf value.Month = 0 Then
                Throw New InvalidCastException("Invalid month number: 0")
            Else
                Throw New InvalidCastException("Month value must be within 1 to 12")
            End If
        End Operator

        Public Shared Widening Operator CType(value As CompuMaster.Calendar.Month) As String
            If value Is Nothing Then
                Return Nothing
            Else
                Return value.ToString
            End If
        End Operator

        Public Shared Narrowing Operator CType(value As String) As Month
            Return Parse(value)
        End Operator

        Public Shared Narrowing Operator CType(value As Int64) As Month
            Return Parse(CType(value, Integer))
        End Operator

        'Public Shared Narrowing Operator CType(value As UInt64) As Month
        '    Return Parse(CType(value, Integer))
        'End Operator

        Public Shared Widening Operator CType(value As CompuMaster.Calendar.Month) As Integer
            If value Is Nothing Then
                Return Nothing
            Else
                Return value.Year * 100 + value.Month
            End If
        End Operator

        'Public Shared Widening Operator CType(value As CompuMaster.Calendar.Month) As UInteger
        '    If value Is Nothing Then
        '        Return Nothing
        '    Else
        '        Return CType(value.Year * 100 + value.Month, UInteger)
        '    End If
        'End Operator

    End Class

End Namespace
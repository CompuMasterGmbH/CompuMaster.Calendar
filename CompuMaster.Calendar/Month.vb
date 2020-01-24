Option Explicit On 
Option Strict On

Namespace CompuMaster.Calendar

    ''' <summary>
    ''' A representation for a month period
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
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' The year of the month
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[wezel]	12.01.2011	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property Year() As Integer
            Get
                Return _Year
            End Get
            Set(ByVal value As Integer)
                If value = 0 Then Throw New ArgumentOutOfRangeException("value")
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
                If value < 0 OrElse value > 12 Then Throw New ArgumentOutOfRangeException("value")
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
        ''' <param name="format"></param>
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
        ''' <param name="format">YYYY for 4-digit year, YY for 2-digit year, MMMM for long month name, MMM for month name abbreviation, MM for always-2-digit month number, M for month number with 1 or 2 digits</param>
        ''' <param name="provider"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        Public Overloads Function ToString(ByVal format As String, ByVal provider As System.IFormatProvider) As String
            Return Me.BeginOfPeriod.ToString(format, provider)
        End Function

        Public Shared Function Parse(value As String, format As String, culture As System.Globalization.CultureInfo) As Month
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
            If Pattern.Contains("MMMM") Then
                Pattern = Pattern.Replace("MMMM", "(?<m>" & Strings.Join(MonthLongNames.ToArray, "|") & ")")
            ElseIf Pattern.Contains("MMM") Then
                Pattern = Pattern.Replace("MMM", "(?<m>" & Strings.Join(MonthShortNames.ToArray, "|") & ")")
            ElseIf Pattern.Contains("MM") Then
                Pattern = Pattern.Replace("MM", "(?<m>\d\d)")
            ElseIf Pattern.Contains("M") Then
                Pattern = Pattern.Replace("M", "(?<m>\d?\d)")
            End If
            Pattern = Pattern.Replace("yyyy", "(?<year4>\d\d\d\d)").Replace("yy", "(?<year2>\d\d)")
            Dim RegEx As New System.Text.RegularExpressions.Regex(Pattern, Text.RegularExpressions.RegexOptions.Compiled Or Text.RegularExpressions.RegexOptions.Singleline Or Text.RegularExpressions.RegexOptions.Multiline)
            If RegEx.IsMatch(value) = False Then
                Throw New ArgumentException("Invalid value", "value")
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
                    Throw New ArgumentException("Invalid value", "value")
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
                    Throw New ArgumentException("Invalid value", "value")
                Else
                    Return New Month(FoundYear, FoundMonth)
                End If
            Else
                Throw New ArgumentException("Invalid value", "value")
            End If
            Throw New NotImplementedException
        End Function

        Public Shared Function TryParse(value As String, format As String, culture As System.Globalization.CultureInfo, ByRef result As Month) As Boolean
            Try
                result = Parse(value, format, culture)
                Return True
            Catch
                Return False
            End Try
        End Function

        Public Shared Function ParseFromUniqueShortName(value As String) As Month
            Dim MonthNames As New System.Collections.Generic.List(Of String)
            For MyCounter As Integer = 1 To 12
                MonthNames.Add(UniqueMonthShortName(MyCounter))
            Next
            Dim Pattern As String = "(?<m>" & Strings.Join(MonthNames.ToArray, "|") & ")\/(?<y>\d\d\d\d)"
            Dim RegEx As New System.Text.RegularExpressions.Regex(Pattern, Text.RegularExpressions.RegexOptions.Compiled Or Text.RegularExpressions.RegexOptions.Singleline Or Text.RegularExpressions.RegexOptions.Multiline)
            If RegEx.IsMatch(value) = False Then
                Throw New ArgumentException("Invalid value", "value")
            End If
            Dim RegMatch As System.Text.RegularExpressions.Match = RegEx.Match(value)
            Dim GroupNames As String() = RegEx.GetGroupNames()
            If RegMatch.Groups.Count >= 2 Then
                Dim FoundYear As Integer = Integer.Parse(RegMatch.Groups("y").Value)
                Dim FoundMonthName As String = RegMatch.Groups("m").Value
                Dim FoundMonth As Integer = 0
                For MyCounter As Integer = 1 To 12
                    If MonthNames(MyCounter - 1) = FoundMonthName Then
                        FoundMonth = MyCounter
                        Exit For
                    End If
                Next
                If FoundYear = 0 OrElse FoundMonth = 0 Then
                    Throw New ArgumentException("Invalid value", "value")
                Else
                    Return New Month(FoundYear, FoundMonth)
                End If
            Else
                Throw New ArgumentException("Invalid value", "value")
            End If
        End Function

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
                    Throw New ArgumentOutOfRangeException("monthNo")
            End Select
            Return Result
        End Function

        ''' <summary>
        ''' A short name in format MMM/yyyy (English names)
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function UniqueShortName() As String
            Return UniqueMonthShortName(Me.Month) & "/" & Me.Year.ToString("0000")
        End Function

        Public Function MonthShortName(cultureName As String) As String
            Return Me.MonthShortName(System.Globalization.CultureInfo.GetCultureInfo(cultureName))
        End Function

        Public Function MonthShortName(culture As System.Globalization.CultureInfo) As String
            Return Me.BeginOfPeriod.ToString("MMM", culture.DateTimeFormat)
        End Function

        Public Function MonthName(cultureName As String) As String
            Return Me.MonthName(System.Globalization.CultureInfo.GetCultureInfo(cultureName))
        End Function

        Public Function MonthName(culture As System.Globalization.CultureInfo) As String
            Return Me.BeginOfPeriod.ToString("MMMM", culture.DateTimeFormat)
        End Function

        ''' <summary>
        ''' Equals method for period classes
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Function Equals(ByVal value As Month) As Boolean
            If Not value Is Nothing Then
                If Me.Year = value.Year AndAlso Me.Month = value.Month Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
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
                Return value1.CompareTo(value2) < 1
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
                Return value1.CompareTo(value2) > 1
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
                Return value1.CompareTo(value2) <= 1
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
                Return value1.CompareTo(value2) >= 1
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
            If value1 Is Nothing Then Throw New ArgumentNullException("value1")
            If value2 Is Nothing Then Throw New ArgumentNullException("value2")
            Dim Result As Integer = 0
            Dim SwappedValues As Boolean, StartMonth As Month, EndMonth As Month
            If value2 > value1 Then
                EndMonth = value2
                StartMonth = value1
                SwappedValues = True
            Else
                EndMonth = value1
                StartMonth = value2
                SwappedValues = False
            End If
            If EndMonth.Year = StartMonth.Year Then
                Result = EndMonth.Month - StartMonth.Month
            Else 'If EndMonth.Year - StartMonth.Year>=1
                Result = EndMonth.Month + (12 - StartMonth.Month + 1) + (EndMonth.Year - StartMonth.Year - 1) * 12
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
        ''' <returns>0 if value is greater, 2 if value is smaller, 1 if value equals</returns>
        ''' <remarks></remarks>
        Public Function CompareTo(ByVal value As Month) As Integer
            If value Is Nothing Then
                Return 2
            ElseIf Me.BeginOfPeriod < value.BeginOfPeriod Then
                Return 0
            ElseIf Me.BeginOfPeriod > value.BeginOfPeriod Then
                Return 2
            Else
                Return 1
            End If
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            If ReferenceEquals(Me, obj) Then
                Return True
            End If

            If ReferenceEquals(obj, Nothing) Then
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

    End Class

End Namespace
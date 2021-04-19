Option Explicit On
Option Strict On

Namespace CompuMaster.Calendar

    ''' <summary>
    '''     A data structure for a complete month number
    ''' </summary>
    Public Class MonthNumber
        Implements IComparable

        Private _Month As Integer
        Public Property Month As Integer
            Get
                Return Me._Month
            End Get
            Set(value As Integer)
                If value > 12 OrElse value < 1 Then Throw New ArgumentOutOfRangeException("Allowed values are 1 to 12")
                Me._Month = value
            End Set
        End Property

        Public Sub New()
        End Sub

        Public Sub New(month As Integer)
            Me.Month = month
        End Sub

        ''' <summary>
        ''' Add months to a MonthNumber, result will always be with a month no. between 1 and 12
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        Public Function AddMonths(value As Integer) As CompuMaster.Calendar.MonthNumber
            If value = 0 Then
                Return Me
            Else
                Dim NewMonthValue As Integer = Me.Month + value
                If NewMonthValue < 1 OrElse NewMonthValue > 12 Then
                    Throw New NotSupportedException("Calculated month value """ & NewMonthValue & """ is invalid (out of borders)")
                End If
                Return New MonthNumber(NewMonthValue)
            End If
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            If obj Is Nothing OrElse GetType(MonthNumber).IsInstanceOfType(obj) = False Then Return False
            Return Me.Month = CType(obj, MonthNumber).Month
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return Month
        End Function

        Public Shared Operator =(left As MonthNumber, right As MonthNumber) As Boolean
            If left Is Nothing AndAlso right Is Nothing Then
                Return True
            ElseIf left Is Nothing Xor right Is Nothing Then
                Return False
            Else
                Return left.Equals(right)
            End If
        End Operator

        Public Shared Operator <>(left As MonthNumber, right As MonthNumber) As Boolean
            Return Not left = right
        End Operator

        ''' <summary>
        ''' Smaller than operator for MonthNumber classes
        ''' </summary>
        ''' <param name="value1"></param>
        ''' <param name="value2"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Operator <(ByVal value1 As MonthNumber, ByVal value2 As MonthNumber) As Boolean
            If value1 Is Nothing And value2 Is Nothing Then
                Return False
            ElseIf value1 Is Nothing And value2 IsNot Nothing Then
                Return True
            Else
                Return value1.CompareTo(value2) < 0
            End If
        End Operator

        ''' <summary>
        ''' Greater than operator for MonthNumber classes
        ''' </summary>
        ''' <param name="value1"></param>
        ''' <param name="value2"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Operator >(ByVal value1 As MonthNumber, ByVal value2 As MonthNumber) As Boolean
            If value1 Is Nothing And value2 Is Nothing Then
                Return False
            ElseIf value1 Is Nothing And value2 IsNot Nothing Then
                Return False
            Else
                Return value1.CompareTo(value2) > 0
            End If
        End Operator

        ''' <summary>
        ''' Smaller than operator for MonthNumber classes
        ''' </summary>
        ''' <param name="value1"></param>
        ''' <param name="value2"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Operator <=(ByVal value1 As MonthNumber, ByVal value2 As MonthNumber) As Boolean
            If value1 Is Nothing And value2 Is Nothing Then
                Return True
            ElseIf value1 Is Nothing And value2 IsNot Nothing Then
                Return True
            Else
                Return value1.CompareTo(value2) <= 0
            End If
        End Operator

        ''' <summary>
        ''' Greater than operator for MonthNumber classes
        ''' </summary>
        ''' <param name="value1"></param>
        ''' <param name="value2"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Operator >=(ByVal value1 As MonthNumber, ByVal value2 As MonthNumber) As Boolean
            If value1 Is Nothing And value2 Is Nothing Then
                Return True
            ElseIf value1 Is Nothing And value2 IsNot Nothing Then
                Return False
            Else
                Return value1.CompareTo(value2) >= 0
            End If
        End Operator

        ''' <summary>
        ''' Text formatting with MM
        ''' </summary>
        ''' <returns></returns>
        Public Overrides Function ToString() As String
            Return Me.Month.ToString("00")
        End Function

        ''' <summary>
        ''' Format the month with a typical datetime format string using the begin date of the period
        ''' </summary>
        ''' <param name="format"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        Public Overloads Function ToString(ByVal format As String) As String
            Return Me.ToString(format, System.Globalization.CultureInfo.InvariantCulture)
        End Function

        ''' <summary>
        ''' Format the month with a typical datetime format string and the given format provider using the begin date of the period
        ''' </summary>
        ''' <param name="format">YYYY for 4-digit year, YY for 2-digit year, MMMM for long month name, MMM for month name abbreviation, MM for always-2-digit month number, M for month number with 1 or 2 digits</param>
        ''' <param name="culture"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        Public Overloads Function ToString(ByVal format As String, ByVal culture As System.Globalization.CultureInfo) As String
            Select Case format
                Case "M"
                    Return Me.Month.ToString("0")
                Case "MM"
                    Return Me.Month.ToString("00")
                Case "MMM"
                    If culture Is System.Globalization.CultureInfo.InvariantCulture Then
                        Return UniqueMonthShortName(Me.Month)
                    Else
                        Return Me.MonthShortName(culture)
                    End If
                Case "MMMM"
                    If culture Is System.Globalization.CultureInfo.InvariantCulture Then
                        Throw New NotSupportedException("Culture-independent full month names don't exist")
                    Else
                        Return Me.MonthName(culture)
                    End If
                Case Else
                    Throw New ArgumentException("Invalid format expression, allowed values are M, MM, MMM or MMMM", NameOf(format))
            End Select
        End Function

        Public Shared Function Parse(value As String, format As String, culture As System.Globalization.CultureInfo) As MonthNumber
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
                Pattern = Pattern.Replace("MMMM", "(?<m>" & String.Join("|", MonthLongNames.ToArray) & ")")
            ElseIf Pattern.Contains("MMM") Then
                Pattern = Pattern.Replace("MMM", "(?<m>" & String.Join("|", MonthShortNames.ToArray) & ")")
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
                Dim FoundMonthName As String = RegMatch.Groups("m").Value
                Dim FoundMonth As Integer = 0
                For MyCounter As Integer = 1 To 12
                    If MonthLongNames(MyCounter - 1) = FoundMonthName OrElse MonthShortNames(MyCounter - 1) = FoundMonthName OrElse MyCounter.ToString("00") = FoundMonthName OrElse MyCounter.ToString("0") = FoundMonthName Then
                        FoundMonth = MyCounter
                        Exit For
                    End If
                Next
                If FoundMonth = 0 Then
                    Throw New ArgumentException("Invalid value", NameOf(value))
                Else
                    Return New MonthNumber(FoundMonth)
                End If
            Else
                Throw New ArgumentException("Invalid value", NameOf(value))
            End If
            Throw New NotImplementedException
        End Function

        Public Shared Function TryParse(value As String, format As String, culture As System.Globalization.CultureInfo, ByRef result As MonthNumber) As Boolean
            Try
                result = Parse(value, format, culture)
                Return True
#Disable Warning CA1031 ' Do not catch general exception types
            Catch
                Return False
#Enable Warning CA1031 ' Do not catch general exception types
            End Try
        End Function

        Public Shared Function ParseFromUniqueShortName(value As String) As MonthNumber
            Dim MonthNames As New System.Collections.Generic.List(Of String)
            For MyCounter As Integer = 1 To 12
                MonthNames.Add(UniqueMonthShortName(MyCounter))
            Next
            Dim Pattern As String = "^(?<m>" & String.Join("|", MonthNames.ToArray) & ")$"
            Dim RegEx As New System.Text.RegularExpressions.Regex(Pattern, Text.RegularExpressions.RegexOptions.Compiled Or Text.RegularExpressions.RegexOptions.Singleline Or Text.RegularExpressions.RegexOptions.Multiline)
            If RegEx.IsMatch(value) = False Then
                Throw New ArgumentException("Invalid value", NameOf(value))
            End If
            Dim RegMatch As System.Text.RegularExpressions.Match = RegEx.Match(value)
            Dim GroupNames As String() = RegEx.GetGroupNames()
            If RegMatch.Groups.Count >= 2 Then
                Dim FoundMonthName As String = RegMatch.Groups("m").Value
                Dim FoundMonth As Integer = 0
                For MyCounter As Integer = 1 To 12
                    If MonthNames(MyCounter - 1) = FoundMonthName Then
                        FoundMonth = MyCounter
                        Exit For
                    End If
                Next
                If FoundMonth = 0 Then
                    Throw New ArgumentException("Invalid value", NameOf(value))
                Else
                    Return New MonthNumber(FoundMonth)
                End If
            Else
                Throw New ArgumentException("Invalid value", NameOf(value))
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
                    Throw New ArgumentOutOfRangeException(NameOf(monthNo))
            End Select
            Return Result
        End Function

        ''' <summary>
        ''' A short name in format MMM/yyyy (English names)
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function UniqueShortName() As String
            Return UniqueMonthShortName(Me.Month)
        End Function

        Public Function MonthShortName(cultureName As String) As String
            Return Me.MonthShortName(System.Globalization.CultureInfo.GetCultureInfo(cultureName))
        End Function

        Public Function MonthShortName(culture As System.Globalization.CultureInfo) As String
            Return New Month(2000, Me.Month).ToString("MMM", culture.DateTimeFormat)
        End Function

        Public Function MonthName(cultureName As String) As String
            Return Me.MonthName(System.Globalization.CultureInfo.GetCultureInfo(cultureName))
        End Function

        Public Function MonthName(culture As System.Globalization.CultureInfo) As String
            Return New Month(2000, Me.Month).ToString("MMMM", culture.DateTimeFormat)
        End Function

        Public Shared Widening Operator CType(value As CompuMaster.Calendar.MonthNumber) As String
            If value Is Nothing Then
                Return Nothing
            Else
                Return value.ToString
            End If
        End Operator

        Public Shared Narrowing Operator CType(value As String) As MonthNumber
            Return Parse(value)
        End Operator

        Public Shared Function Parse(value As String) As MonthNumber
            If value = Nothing Then
                Return Nothing
            ElseIf value.Length <> 2 Then
                Throw New ArgumentException("value must be formatted as MM")
            Else
                Return New MonthNumber(Integer.Parse(value))
            End If
        End Function

        Public Shared Function TryParse(value As String, ByRef result As MonthNumber) As Boolean
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
        ''' Create an instance of the following period
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function NextMonth() As MonthNumber
            If Me.Month = 12 Then
                Throw New NotSupportedException("NextMonth not available at end of year")
            Else
                Return Me.AddMonths(1)
            End If
        End Function

        ''' <summary>
        ''' Create an instance of the previous period
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function PreviousMonth() As MonthNumber
            If Me.Month = 1 Then
                Throw New NotSupportedException("PreviousMonth not available at begin of year")
            Else
                Return Me.AddMonths(-1)
            End If
        End Function

        ''' <summary>
        ''' Compares a value to the current instance value
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        Public Function CompareTo(obj As Object) As Integer Implements IComparable.CompareTo
            Return Me.CompareTo(CType(obj, MonthNumber))
        End Function

        ''' <summary>
        ''' Compares a value to the current instance value
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns>-1 if value is smaller, +1 if value is greater, 0 if value equals</returns>
        Public Function CompareTo(ByVal value As MonthNumber) As Integer
            If value Is Nothing Then
                Return 1
            ElseIf Me.Month = value.Month Then
                Return 0
            ElseIf Me.Month < value.Month Then
                Return -1
            Else
                Return 1
            End If
        End Function

    End Class

End Namespace
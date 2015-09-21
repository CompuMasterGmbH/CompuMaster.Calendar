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
                If value < 1 OrElse value > 12 Then Throw New ArgumentOutOfRangeException("value")
                _Month = value
            End Set
        End Property

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
        ''' <param name="format"></param>
        ''' <param name="provider"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        Public Overloads Function ToString(ByVal format As String, ByVal provider As System.IFormatProvider) As String
            Return Me.BeginOfPeriod.ToString(format, provider)
        End Function

        ''' <summary>
        ''' A short name in format MMM/yyyy
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function UniqueShortName() As String
            Dim Result As String = ""
            Select Case Me.Month
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
            End Select
            Result &= "/" & Me.Year.ToString("0000")
            Return Result
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
            Return value1.Equals(value2)
        End Operator

        ''' <summary>
        ''' Not-Equals operator for Month classes
        ''' </summary>
        ''' <param name="value1"></param>
        ''' <param name="value2"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Operator <>(ByVal value1 As Month, ByVal value2 As Month) As Boolean
            Return Not value1.Equals(value2)
        End Operator

        ''' <summary>
        ''' Smaller than operator for Month classes
        ''' </summary>
        ''' <param name="value1"></param>
        ''' <param name="value2"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Operator <(ByVal value1 As Month, ByVal value2 As Month) As Boolean
            Return value1.CompareTo(value2) < 1
        End Operator

        ''' <summary>
        ''' Greater than operator for Month classes
        ''' </summary>
        ''' <param name="value1"></param>
        ''' <param name="value2"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Operator >(ByVal value1 As Month, ByVal value2 As Month) As Boolean
            Return value1.CompareTo(value2) > 1
        End Operator

        ''' <summary>
        ''' Smaller than operator for Month classes
        ''' </summary>
        ''' <param name="value1"></param>
        ''' <param name="value2"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Operator <=(ByVal value1 As Month, ByVal value2 As Month) As Boolean
            Return value1.CompareTo(value2) <= 1
        End Operator

        ''' <summary>
        ''' Greater than operator for Month classes
        ''' </summary>
        ''' <param name="value1"></param>
        ''' <param name="value2"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Operator >=(ByVal value1 As Month, ByVal value2 As Month) As Boolean
            Return value1.CompareTo(value2) >= 1
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
            If Me.BeginOfPeriod < value.BeginOfPeriod Then
                Return 0
            ElseIf Me.BeginOfPeriod > value.BeginOfPeriod Then
                Return 2
            Else
                Return 1
            End If
        End Function

    End Class

End Namespace
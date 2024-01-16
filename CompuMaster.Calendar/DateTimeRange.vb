Option Explicit On
Option Strict On

Namespace CompuMaster.Calendar

    ''' <summary>
    ''' A date range with optional date limits
    ''' </summary>
    Public Class DateTimeRange
        Implements IEqualityComparer, IEqualityComparer(Of DateTimeRange), IComparable

        ''' <summary>
        ''' Create a new date range without any date limits
        ''' </summary>
        Public Sub New()
        End Sub

        ''' <summary>
        ''' Create a new date range with date limits
        ''' </summary>
        ''' <param name="from"></param>
        ''' <param name="till"></param>
        Public Sub New(from As DateTime?, till As DateTime?)
            Me.From = from
            Me.Till = till
        End Sub

        ''' <summary>
        ''' Parse a string representation of date range, e.g. for invariant culture "2005-01-01 00:00:00 - 2005-12-31 23:59:59" or "2005-01-01 00:00:00 - *" or "* - 2005-12-31 23:59:59" or "" (=unlimited range)
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        Public Shared Function Parse(value As String) As DateTimeRange
            Return Parse(value, Nothing)
        End Function

        ''' <summary>
        ''' Parse a string representation of date range, e.g. for invariant culture "2005-01-01 00:00:00 - 2005-12-31 23:59:59" or "2005-01-01 00:00:00 - *" or "* - 2005-12-31 23:59:59" or "" (=unlimited range)
        ''' </summary>
        ''' <param name="value"></param>
        ''' <param name="culture"></param>
        ''' <returns></returns>
        Public Shared Function Parse(value As String, culture As System.Globalization.CultureInfo) As DateTimeRange
            If value = Nothing OrElse value = "* - *" Then Return New DateTimeRange
#Disable Warning CA1861 ' Konstantenmatrizen als Argumente vermeiden
            Dim Parts = value.Split(New String() {" - "}, StringSplitOptions.None)
#Enable Warning CA1861 ' Konstantenmatrizen als Argumente vermeiden
            If Parts.Length <> 2 Then Throw New FormatException("Invalid date format")
            Dim FromDate As DateTime?
            If Parts(0).Trim <> Nothing AndAlso Parts(0).Trim <> "*" Then
                If culture Is Nothing Then
                    FromDate = DateTime.Parse(Parts(0).Trim)
                Else
                    FromDate = DateTime.Parse(Parts(0).Trim, culture)
                End If
            Else
                FromDate = Nothing
            End If
            Dim TillDate As DateTime?
            If Parts(1).Trim <> Nothing AndAlso Parts(1).Trim <> "*" Then
                If culture Is Nothing Then
                    TillDate = DateTime.Parse(Parts(1).Trim)
                Else
                    TillDate = DateTime.Parse(Parts(1).Trim, culture)
                End If
            Else
                TillDate = Nothing
            End If
            Return New DateTimeRange(FromDate, TillDate)
        End Function

        ''' <summary>
        ''' A from date value or null/Nothing/DateTime.MinValue for an open start date
        ''' </summary>
        ''' <returns></returns>
        Public Property From As DateTime?

        ''' <summary>
        ''' A till date value or null/Nothing/DateTime.MaxValue for an open end date
        ''' </summary>
        ''' <returns></returns>
        Public Property Till As DateTime?

        ''' <summary>
        ''' A string representation of date range, e.g. for invariant culture "2005-01-01 00:00:00 - 2005-12-31 23:59:59" or "2005-01-01 00:00:00 - *" or "* - 2005-12-31 23:59:59" or "" (=unlimited range)
        ''' </summary>
        ''' <remarks>Milliseconds or timezone information are ignored</remarks>
        ''' <returns></returns>
        Public Overloads Overrides Function ToString() As String
            Return Me.ToString(Nothing)
        End Function

        ''' <summary>
        ''' A string representation of date range, e.g. for invariant culture "2005-01-01 00:00:00 - 2005-12-31 23:59:59" or "2005-01-01 00:00:00 - *" or "* - 2005-12-31 23:59:59" or "" (=unlimited range)
        ''' </summary>
        ''' <param name="culture"></param>
        ''' <remarks>Milliseconds or timezone information are ignored</remarks>
        ''' <returns></returns>
        Public Overloads Function ToString(culture As System.Globalization.CultureInfo) As String
            If Me.From.HasValue = False AndAlso Me.Till.HasValue = False Then
                Return ""
            Else
                Dim Result As String = ""
                If Me.From.HasValue Then
                    If culture Is Nothing Then
                        Result &= Me.From.Value.ToString
                    ElseIf culture.TwoLetterISOLanguageName = "iv" Then
                        Result &= Me.From.Value.ToString("yyyy-MM-dd HH:mm:ss") 'override invariant culture with ISO8601 format
                    Else
                        Result &= Me.From.Value.ToString(culture)
                    End If
                Else
                    Result &= "*"
                End If
                Result &= " - "
                If Me.Till.HasValue Then
                    If culture Is Nothing Then
                        Result &= Me.Till.Value.ToString
                    ElseIf culture.TwoLetterISOLanguageName = "iv" Then
                        Result &= Me.Till.Value.ToString("yyyy-MM-dd HH:mm:ss") 'override invariant culture with ISO8601 format
                    Else
                        Result &= Me.Till.Value.ToString(culture)
                    End If
                Else
                    Result &= "*"
                End If
                Return Result
            End If
        End Function

        Public Shared Operator =(ByVal a As DateTimeRange, ByVal b As DateTimeRange) As Boolean
            If a Is Nothing AndAlso b Is Nothing Then
                Return True
            ElseIf a Is Nothing AndAlso b IsNot Nothing Then
                Return False
            ElseIf a IsNot Nothing AndAlso b Is Nothing Then
                Return False
            Else
                Return a.From.GetValueOrDefault = b.From.GetValueOrDefault AndAlso a.Till.GetValueOrDefault(DateTime.MaxValue) = b.Till.GetValueOrDefault(DateTime.MaxValue)
            End If
        End Operator

        Public Shared Operator <>(ByVal a As DateTimeRange, ByVal b As DateTimeRange) As Boolean
            Return Not a = b
        End Operator

        Public Function CompareTo(obj As Object) As Integer Implements IComparable.CompareTo
            Dim Other As DateTimeRange = CType(obj, DateTimeRange)
            If Me.From.GetValueOrDefault < Other.From.GetValueOrDefault Then
                Return -1
            ElseIf Me.From.GetValueOrDefault > Other.From.GetValueOrDefault Then
                Return 1
            ElseIf Me.Till.GetValueOrDefault(DateTime.MaxValue) < Other.Till.GetValueOrDefault(DateTime.MaxValue) Then
                Return -1
            ElseIf Me.Till.GetValueOrDefault(DateTime.MaxValue) > Other.Till.GetValueOrDefault(DateTime.MaxValue) Then
                Return 1
            Else
                Return 0
            End If
        End Function

        ''' <summary>
        ''' No values means 'fully unlimited range', same with values from DateTime.MinValue till DateTime.MaxValue
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property IsUnlimitedRange As Boolean
            Get
                Return Me.IsOpenToThePast AndAlso Me.IsOpenToTheFuture
            End Get
        End Property

        ''' <summary>
        ''' No value or DateTime.MinValue means 'open to the past'
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property IsOpenToThePast As Boolean
            Get
                Return Me.From.HasValue = False OrElse Me.From.Value = DateTime.MinValue
            End Get
        End Property

        ''' <summary>
        ''' No value or DateTime.MaxValue means 'open to the future'
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property IsOpenToTheFuture As Boolean
            Get
                Return Me.Till.HasValue = False OrElse Me.Till.Value = DateTime.MaxValue
            End Get
        End Property

        Public Overrides Function GetHashCode() As Integer
            If Me.IsUnlimitedRange Then
                Return 0
            Else
                Return Me.ToString(System.Globalization.CultureInfo.InvariantCulture).GetHashCode
            End If
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            Return IEqualityComparer_OfDateTimeRange_Equals(Me, CType(obj, DateTimeRange))
        End Function

        Private Function IEqualityComparer_Equals(x As Object, y As Object) As Boolean Implements IEqualityComparer.Equals
            Return CType(x, DateTimeRange) = CType(y, DateTimeRange)
        End Function

        Private Function IEqualityComparer_GetHashCode(obj As Object) As Integer Implements IEqualityComparer.GetHashCode
            Return Me.IEqualityComparer_OfDateTimeRange_GetHashCode(CType(obj, DateTimeRange))
        End Function

        Private Function IEqualityComparer_OfDateTimeRange_Equals(x As DateTimeRange, y As DateTimeRange) As Boolean Implements IEqualityComparer(Of DateTimeRange).Equals
            Return x = y
        End Function

        Private Function IEqualityComparer_OfDateTimeRange_GetHashCode(obj As DateTimeRange) As Integer Implements IEqualityComparer(Of DateTimeRange).GetHashCode
            If obj Is Nothing OrElse obj.IsUnlimitedRange Then
                Return 0
            Else
                Return obj.GetHashCode
            End If
        End Function
    End Class

End Namespace
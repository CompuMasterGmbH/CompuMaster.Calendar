Option Explicit On
Option Strict On

Namespace CompuMaster.Calendar

    ''' <summary>
    ''' A range of months whose inclusive first and last boundaries are independently optional.
    ''' </summary>
    Public Class OptionallyBoundedMonthRange
        Implements IEquatable(Of OptionallyBoundedMonthRange)

        Private ReadOnly _FirstMonth As Month
        Private ReadOnly _LastMonth As Month

        ''' <summary>
        ''' Creates a range. Nothing represents an unbounded side.
        ''' </summary>
        Public Sub New(firstMonth As Month, lastMonth As Month)
            If firstMonth IsNot Nothing AndAlso lastMonth IsNot Nothing AndAlso firstMonth > lastMonth Then
                Throw New ArgumentException("First month must be before or equal to last month")
            End If

            If firstMonth IsNot Nothing Then _FirstMonth = firstMonth.Clone()
            If lastMonth IsNot Nothing Then _LastMonth = lastMonth.Clone()
        End Sub

        ''' <summary>The inclusive first month, or Nothing when the start is unbounded.</summary>
        Public ReadOnly Property FirstMonth As Month
            Get
                Return _FirstMonth
            End Get
        End Property

        ''' <summary>The inclusive last month, or Nothing when the end is unbounded.</summary>
        Public ReadOnly Property LastMonth As Month
            Get
                Return _LastMonth
            End Get
        End Property

        Public ReadOnly Property HasFirstMonth As Boolean
            Get
                Return _FirstMonth IsNot Nothing
            End Get
        End Property

        Public ReadOnly Property HasLastMonth As Boolean
            Get
                Return _LastMonth IsNot Nothing
            End Get
        End Property

        ''' <summary>Returns a representation such as "* - 2026-12" or "2027-01 - *".</summary>
        Public Overrides Function ToString() As String
            Dim FirstText As String = If(_FirstMonth Is Nothing, "*", _FirstMonth.ToString())
            Dim LastText As String = If(_LastMonth Is Nothing, "*", _LastMonth.ToString())
            Return FirstText & " - " & LastText
        End Function

        Public Shared Function Parse(value As String) As OptionallyBoundedMonthRange
            Dim FirstMonth As Month = Nothing
            Dim LastMonth As Month = Nothing
            ParseBoundaries(value, FirstMonth, LastMonth)
            Return New OptionallyBoundedMonthRange(FirstMonth, LastMonth)
        End Function

        Public Shared Function TryParse(value As String, ByRef result As OptionallyBoundedMonthRange) As Boolean
            Try
                result = Parse(value)
                Return True
            Catch ex As Exception When TypeOf ex Is ArgumentException OrElse TypeOf ex Is FormatException OrElse TypeOf ex Is OverflowException
                result = Nothing
                Return False
            End Try
        End Function

        Protected Shared Sub ParseBoundaries(value As String, ByRef firstMonth As Month, ByRef lastMonth As Month)
            If String.IsNullOrEmpty(value) Then Throw New FormatException("Value must contain two month boundaries separated by "" - """)

            Const Separator As String = " - "
            Dim SeparatorIndex As Integer = value.IndexOf(Separator, StringComparison.Ordinal)
            If SeparatorIndex < 0 OrElse value.IndexOf(Separator, SeparatorIndex + Separator.Length, StringComparison.Ordinal) >= 0 Then
                Throw New FormatException("Value must contain exactly two month boundaries separated by "" - """)
            End If

            Dim FirstText As String = value.Substring(0, SeparatorIndex)
            Dim LastText As String = value.Substring(SeparatorIndex + Separator.Length)
            If FirstText.Length = 0 OrElse LastText.Length = 0 Then Throw New FormatException("A boundary must be a month or *")

            firstMonth = ParseBoundary(FirstText)
            lastMonth = ParseBoundary(LastText)
        End Sub

        Private Shared Function ParseBoundary(value As String) As Month
            If value = "*" Then Return Nothing
            Try
                Return Month.Parse(value)
            Catch ex As Exception When TypeOf ex Is ArgumentException OrElse TypeOf ex Is FormatException OrElse TypeOf ex Is OverflowException
                Throw New FormatException("A boundary must be formatted as yyyy-MM or *", ex)
            End Try
        End Function

        Public Function Contains(value As Month) As Boolean
            If value Is Nothing Then Return False
            Return (_FirstMonth Is Nothing OrElse value >= _FirstMonth) AndAlso
                   (_LastMonth Is Nothing OrElse value <= _LastMonth)
        End Function

        Public Function Contains(value As OptionallyBoundedMonthRange) As Boolean
            If value Is Nothing Then Throw New ArgumentNullException(NameOf(value))
            Dim ContainsStart As Boolean = _FirstMonth Is Nothing OrElse
                (value.FirstMonth IsNot Nothing AndAlso value.FirstMonth >= _FirstMonth)
            Dim ContainsEnd As Boolean = _LastMonth Is Nothing OrElse
                (value.LastMonth IsNot Nothing AndAlso value.LastMonth <= _LastMonth)
            Return ContainsStart AndAlso ContainsEnd
        End Function

        Public Function Contains(value As MonthRange) As Boolean
            If value Is Nothing Then Throw New ArgumentNullException(NameOf(value))
            If value.IsEmpty Then Return False
            Return Contains(New OptionallyBoundedMonthRange(value.FirstMonth, value.LastMonth))
        End Function

        Public Function Overlaps(value As OptionallyBoundedMonthRange) As Boolean
            If value Is Nothing Then Throw New ArgumentNullException(NameOf(value))
            If _LastMonth IsNot Nothing AndAlso value.FirstMonth IsNot Nothing AndAlso _LastMonth < value.FirstMonth Then Return False
            If value.LastMonth IsNot Nothing AndAlso _FirstMonth IsNot Nothing AndAlso value.LastMonth < _FirstMonth Then Return False
            Return True
        End Function

        Public Function Overlaps(value As MonthRange) As Boolean
            If value Is Nothing Then Throw New ArgumentNullException(NameOf(value))
            If value.IsEmpty Then Return False
            Return Overlaps(New OptionallyBoundedMonthRange(value.FirstMonth, value.LastMonth))
        End Function

        Public Function TryToMonthRange(ByRef result As MonthRange) As Boolean
            If _FirstMonth Is Nothing OrElse _LastMonth Is Nothing Then
                result = Nothing
                Return False
            End If
            result = New MonthRange(_FirstMonth, _LastMonth)
            Return True
        End Function

        Public Function ToMonthRange() As MonthRange
            Dim Result As MonthRange = Nothing
            If Not TryToMonthRange(Result) Then Throw New InvalidOperationException("Both boundaries are required for conversion to MonthRange")
            Return Result
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            Return Equals(TryCast(obj, OptionallyBoundedMonthRange))
        End Function

        Public Overloads Function Equals(other As OptionallyBoundedMonthRange) As Boolean Implements IEquatable(Of OptionallyBoundedMonthRange).Equals
            Return other IsNot Nothing AndAlso _FirstMonth = other.FirstMonth AndAlso _LastMonth = other.LastMonth
        End Function

        Public Overrides Function GetHashCode() As Integer
            Dim Result As Integer = 17
            Result = Result * 31 + If(_FirstMonth Is Nothing, 0, _FirstMonth.GetHashCode())
            Result = Result * 31 + If(_LastMonth Is Nothing, 0, _LastMonth.GetHashCode())
            Return Result
        End Function

        Public Shared Operator =(value1 As OptionallyBoundedMonthRange, value2 As OptionallyBoundedMonthRange) As Boolean
            If value1 Is Nothing Then Return value2 Is Nothing
            Return value1.Equals(value2)
        End Operator

        Public Shared Operator <>(value1 As OptionallyBoundedMonthRange, value2 As OptionallyBoundedMonthRange) As Boolean
            Return Not value1 = value2
        End Operator

        Public Shared Widening Operator CType(value As MonthRange) As OptionallyBoundedMonthRange
            If value Is Nothing Then Return Nothing
            If value.IsEmpty Then Throw New InvalidCastException("MonthRange.Empty cannot be converted because the target type has no empty state")
            Return New OptionallyBoundedMonthRange(value.FirstMonth, value.LastMonth)
        End Operator

        Public Shared Narrowing Operator CType(value As OptionallyBoundedMonthRange) As MonthRange
            If value Is Nothing Then Return Nothing
            Return value.ToMonthRange()
        End Operator

    End Class

End Namespace

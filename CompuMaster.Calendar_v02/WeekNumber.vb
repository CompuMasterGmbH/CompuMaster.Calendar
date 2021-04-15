Option Explicit On
Option Strict On

Namespace CompuMaster.Calendar

    ''' <summary>
    '''     A data structure for a complete week number inclusive the year information which is related to a pure week number
    ''' </summary>
    Public Class WeekNumber
        Implements IComparable

        Private _Week As Integer
        Public Property Week As Integer
            Get
                Return Me._Week
            End Get
            Set(value As Integer)
                If value >= 100 OrElse value < 0 Then Throw New ArgumentOutOfRangeException("Allowed values are 0 to 99")
                Me._Week = value
            End Set
        End Property

        Public Property Year As Integer

        Public Sub New()
        End Sub

        Public Sub New(year As Integer, week As Integer)
            Me.Year = year
            Me.Week = week
        End Sub

        ''' <summary>
        ''' Add years and weeks to a WeekNumber, result will always be with a week no. between 0 and 99
        ''' </summary>
        ''' <param name="years"></param>
        ''' <param name="weeks"></param>
        ''' <returns></returns>
        Public Function Add(years As Integer, weeks As Integer) As CompuMaster.Calendar.WeekNumber
            Return Me.AddYears(years).AddWeeks(weeks)
        End Function

        ''' <summary>
        ''' Add years and weeks to a WeekNumber, result will always be with a week no. between 0 and 99
        ''' </summary>
        ''' <param name="span"></param>
        ''' <returns></returns>
        Public Function Add(span As WeekNumberSpan) As CompuMaster.Calendar.WeekNumber
            Return Me.AddYears(span.Years).AddWeeks(span.Weeks)
        End Function

        ''' <summary>
        ''' Add years to a WeekNumber, result will always keep the week unchanged 
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        Public Function AddYears(value As Integer) As CompuMaster.Calendar.WeekNumber
            Return New CompuMaster.Calendar.WeekNumber(Me.Year + value, Me.Week)
        End Function

        ''' <summary>
        ''' Add weeks to a WeekNumber, result will always be with a week no. between 0 and 99
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        Public Function AddWeeks(value As Integer) As CompuMaster.Calendar.WeekNumber
            If value = 0 Then
                Return New CompuMaster.Calendar.WeekNumber(Me.Year, Me.Week)
            Else
                Dim NewWeekValue As Integer
                If value < 0 And Me.Week = 0 Then
                    NewWeekValue = Me.Week + 1 + value
                Else
                    NewWeekValue = Me.Week + value
                End If
                If NewWeekValue < 0 OrElse NewWeekValue > 99 Then
                    Throw New NotSupportedException("")
                End If
                Return New WeekNumber(Me.Year, NewWeekValue)
            End If
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            If obj Is Nothing OrElse GetType(WeekNumber).IsInstanceOfType(obj) = False Then Return False
            Return Me.Week = CType(obj, WeekNumber).Week AndAlso Me.Year = CType(obj, WeekNumber).Year
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return Year * 100 + Week
        End Function

        Public Shared Operator =(left As WeekNumber, right As WeekNumber) As Boolean
            Return left.Equals(right)
        End Operator

        Public Shared Operator <>(left As WeekNumber, right As WeekNumber) As Boolean
            Return Not left = right
        End Operator

        ''' <summary>
        ''' Smaller than operator for WeekNumber classes
        ''' </summary>
        ''' <param name="value1"></param>
        ''' <param name="value2"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Operator <(ByVal value1 As WeekNumber, ByVal value2 As WeekNumber) As Boolean
            If value1 Is Nothing And value2 Is Nothing Then
                Return False
            ElseIf value1 Is Nothing And value2 IsNot Nothing Then
                Return True
            Else
                Return value1.CompareTo(value2) < 0
            End If
        End Operator

        ''' <summary>
        ''' Greater than operator for WeekNumber classes
        ''' </summary>
        ''' <param name="value1"></param>
        ''' <param name="value2"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Operator >(ByVal value1 As WeekNumber, ByVal value2 As WeekNumber) As Boolean
            If value1 Is Nothing And value2 Is Nothing Then
                Return False
            ElseIf value1 Is Nothing And value2 IsNot Nothing Then
                Return False
            Else
                Return value1.CompareTo(value2) > 0
            End If
        End Operator

        ''' <summary>
        ''' Smaller than operator for WeekNumber classes
        ''' </summary>
        ''' <param name="value1"></param>
        ''' <param name="value2"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Operator <=(ByVal value1 As WeekNumber, ByVal value2 As WeekNumber) As Boolean
            If value1 Is Nothing And value2 Is Nothing Then
                Return True
            ElseIf value1 Is Nothing And value2 IsNot Nothing Then
                Return True
            Else
                Return value1.CompareTo(value2) <= 0
            End If
        End Operator

        ''' <summary>
        ''' Greater than operator for WeekNumber classes
        ''' </summary>
        ''' <param name="value1"></param>
        ''' <param name="value2"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Operator >=(ByVal value1 As WeekNumber, ByVal value2 As WeekNumber) As Boolean
            If value1 Is Nothing And value2 Is Nothing Then
                Return True
            ElseIf value1 Is Nothing And value2 IsNot Nothing Then
                Return False
            Else
                Return value1.CompareTo(value2) >= 0
            End If
        End Operator

        ''' <summary>
        ''' The time span of years and weeks between two values
        ''' </summary>
        ''' <param name="value1"></param>
        ''' <param name="value2"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Operator -(value1 As WeekNumber, value2 As WeekNumber) As WeekNumberSpan
            If value1 Is Nothing Then Throw New ArgumentNullException(NameOf(value1))
            If value2 Is Nothing Then Throw New ArgumentNullException(NameOf(value2))
            If value1 = value2 Then
                Return New WeekNumberSpan
            Else
                Return New WeekNumberSpan With
                    {
                    .Years = value1.Year - value2.Year,
                    .Weeks = value1.Week - value2.Week
                    }
            End If
        End Operator

        Public Overrides Function ToString() As String
            Return Me.Year.ToString("0000") & "/WK" & Me.Week.ToString("00")
        End Function

        Public Shared Widening Operator CType(value As CompuMaster.Calendar.WeekNumber) As String
            If value Is Nothing Then
                Return Nothing
            Else
                Return value.ToString
            End If
        End Operator

        Public Shared Narrowing Operator CType(value As String) As WeekNumber
            Return Parse(value)
        End Operator

        Public Shared Function Parse(value As String) As WeekNumber
            If value = Nothing Then
                Return Nothing
            ElseIf value.Length <> 9 OrElse value.Substring(4, 3) <> "/WK" Then
                Throw New ArgumentException("value must be formatted as yyyy""/WK""ww")
            Else
                Return New WeekNumber(Integer.Parse(value.Substring(0, 4)), Integer.Parse(value.Substring(7, 2)))
            End If
        End Function

        Public Shared Function TryParse(value As String, ByRef result As WeekNumber) As Boolean
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
        Public Function NextWeek() As WeekNumber
            If Me.Week = 0 Then
                Return New WeekNumber(Me.Year, 1)
            ElseIf Me.Week = 99 Then
                Throw New NotSupportedException("NextWeek not available in this year")
            Else
                Return Me.AddWeeks(1)
            End If
        End Function

        ''' <summary>
        ''' Create an instance of the previous period
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function PreviousWeek() As WeekNumber
            If Me.Week = 0 Then
                Throw New NotSupportedException("PreviousWeek not available in this year")
            ElseIf Me.Week = 1 Then
                Throw New NotSupportedException("PreviousWeek not available in this year")
            Else
                Return Me.AddWeeks(-1)
            End If
        End Function

        ''' <summary>
        ''' Create an instance of same year, but week part is zeroed
        ''' </summary>
        ''' <returns></returns>
        Public Function ZeroWeek() As WeekNumber
            Return New WeekNumber(Me.Year, 0)
        End Function

        ''' <summary>
        ''' Create an instance of same year, but first week
        ''' </summary>
        ''' <returns></returns>
        Public Function FirstWeek() As WeekNumber
            Return New WeekNumber(Me.Year, 1)
        End Function

        ''' <summary>
        ''' Compares a value to the current instance value
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        Public Function CompareTo(obj As Object) As Integer Implements IComparable.CompareTo
            Return Me.CompareTo(CType(obj, WeekNumber))
        End Function

        ''' <summary>
        ''' Compares a value to the current instance value
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns>-1 if value is smaller, +1 if value is greater, 0 if value equals</returns>
        Public Function CompareTo(ByVal value As WeekNumber) As Integer
            If value Is Nothing Then
                Return 1
            ElseIf Me.Year = value.Year AndAlso Me.Week = value.Week Then
                Return 0
            ElseIf Me.Week = 0 Then
                If Me.Year < value.Year Then
                    Return -1
                ElseIf Me.Year > value.Year Then
                    Return 1
                Else 'Me.Year = value.Year, but 0 < value.Week
                    Return -1
                End If
            ElseIf value.Week = 0 Then
                If Me.Year < value.Year Then
                    Return -1
                ElseIf Me.Year > value.Year Then
                    Return 1
                Else 'Me.Year = value.Year, but Me.Week > 0
                    Return 1
                End If
            ElseIf Me.Year < value.Year Then
                Return -1
            ElseIf Me.Year > value.Year Then
                Return 1
            Else 'Me.Year = value.Year, but Me.Week > 0
                If Me.Week < value.Week Then
                    Return -1
                ElseIf Me.Week > value.Week Then
                    Return 1
                Else
                    Return 0
                End If
            End If
        End Function

    End Class

End Namespace
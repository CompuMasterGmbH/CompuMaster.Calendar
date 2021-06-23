Option Explicit On
Option Strict On

Namespace CompuMaster.Calendar

    ''' <summary>
    ''' A range of months
    ''' </summary>
    Public Class MonthRange
        Implements IComparable

        Public Sub New(firstPeriod As CompuMaster.Calendar.Month, lastPeriod As CompuMaster.Calendar.Month)
            If firstPeriod Is Nothing Then Throw New ArgumentNullException(NameOf(firstPeriod))
            If lastPeriod Is Nothing Then Throw New ArgumentNullException(NameOf(lastPeriod))
            If firstPeriod > lastPeriod Then Throw New ArgumentException("First period must be before last period")
            Me._FirstMonth = firstPeriod
            Me._LastMonth = lastPeriod
        End Sub

        Private _FirstMonth As CompuMaster.Calendar.Month
        Public ReadOnly Property FirstMonth As CompuMaster.Calendar.Month
            Get
                Return _FirstMonth
            End Get
        End Property

        Private _LastMonth As CompuMaster.Calendar.Month
        Public ReadOnly Property LastMonth As CompuMaster.Calendar.Month
            Get
                Return _LastMonth
            End Get
        End Property

        Public Overrides Function ToString() As String
            Return Me.FirstMonth.ToString & " - " & Me.LastMonth.ToString
        End Function

        ''' <summary>
        ''' Compares a value to the current instance value
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns>-1 if value is smaller, +1 if value is greater, 0 if value equals</returns>
        Public Function CompareTo(obj As Object) As Integer Implements IComparable.CompareTo
            Return Me.CompareTo(CType(obj, MonthRange))
        End Function

        ''' <summary>
        ''' Equals method for MonthRange classes
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

            If GetType(MonthRange).IsInstanceOfType(obj) = False Then
                Return False
            End If

            Return Me.FirstMonth = CType(obj, MonthRange).FirstMonth AndAlso Me.LastMonth = CType(obj, MonthRange).LastMonth
        End Function

        ''' <summary>
        ''' Equals method for MonthRange classes
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Function Equals(ByVal value As MonthRange) As Boolean
            If value IsNot Nothing Then
                If Me.FirstMonth = value.FirstMonth AndAlso Me.LastMonth = value.LastMonth Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        End Function

        ''' <summary>
        ''' Compares a value to the current instance value
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns>-1 if value is smaller, +1 if value is greater, 0 if value equals</returns>
        Public Function CompareTo(value As MonthRange) As Integer
            If value Is Nothing Then
                Return 1
            ElseIf Me.FirstMonth < value.FirstMonth Then
                Return -1
            ElseIf Me.FirstMonth > value.FirstMonth Then
                Return 1
            ElseIf Me.LastMonth < value.LastMonth Then
                Return -1
            ElseIf Me.LastMonth > value.LastMonth Then
                Return 1
            Else
                Return 0
            End If
        End Function

        ''' <summary>
        ''' Equals operator for Month classes
        ''' </summary>
        ''' <param name="value1"></param>
        ''' <param name="value2"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Operator =(ByVal value1 As MonthRange, ByVal value2 As MonthRange) As Boolean
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
        Public Shared Operator <>(ByVal value1 As MonthRange, ByVal value2 As MonthRange) As Boolean
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
        Public Shared Operator <(ByVal value1 As MonthRange, ByVal value2 As MonthRange) As Boolean
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
        Public Shared Operator >(ByVal value1 As MonthRange, ByVal value2 As MonthRange) As Boolean
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
        Public Shared Operator <=(ByVal value1 As MonthRange, ByVal value2 As MonthRange) As Boolean
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
        Public Shared Operator >=(ByVal value1 As MonthRange, ByVal value2 As MonthRange) As Boolean
            If value1 Is Nothing And value2 Is Nothing Then
                Return True
            ElseIf value1 Is Nothing And value2 IsNot Nothing Then
                Return False
            Else
                Return value1.CompareTo(value2) >= 0
            End If
        End Operator

        ''' <summary>
        ''' The number of months covered by the range
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property MonthCount As Integer
            Get
                Return Me._LastMonth - Me._FirstMonth + 1
            End Get
        End Property

        Private _Months As Month()
#Disable Warning CA1819 ' Properties should not return arrays
        ''' <summary>
        ''' All months covered by the range
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property Months As Month()
#Enable Warning CA1819 ' Properties should not return arrays
            Get
                If _Months Is Nothing Then
                    Dim AllMonths As New List(Of Month)
                    Dim CurrentMonth As Month = Me._FirstMonth
                    Do While CurrentMonth <= Me._LastMonth
                        AllMonths.Add(CurrentMonth)
                        CurrentMonth = CurrentMonth.NextPeriod
                    Loop
                    _Months = AllMonths.ToArray
                End If
                Return _Months
            End Get
        End Property

        ''' <summary>
        ''' Is a specified Month a member of this range
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        Public Function Contains(value As Month) As Boolean
            Return value >= Me.FirstMonth AndAlso value <= Me.LastMonth
        End Function

        ''' <summary>
        ''' Is a specified MonthRange a member of this range
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        Public Function Contains(value As MonthRange) As Boolean
            If value Is Nothing Then Throw New ArgumentNullException(NameOf(value))
            Return value.FirstMonth >= Me.FirstMonth AndAlso value.LastMonth <= Me.LastMonth
        End Function

    End Class

End Namespace
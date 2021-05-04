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
            Me._FirstMonth = firstPeriod
            Me._LastMonth = lastPeriod
            CheckForExchangeMonthValues()
        End Sub

        ''' <summary>
        ''' Exchange FirstPeriod and LastPeriod values if LastPeriod is before FirstPeriod
        ''' </summary>
        Private Sub CheckForExchangeMonthValues()
            If _FirstMonth > _LastMonth Then
                Dim BufferedFirstPeriod = _FirstMonth
                Me._FirstMonth = _LastMonth
                Me._LastMonth = BufferedFirstPeriod
            End If
        End Sub

        Private _FirstMonth As CompuMaster.Calendar.Month
        Public Property FirstMonth As CompuMaster.Calendar.Month
            Get
                Return _FirstMonth
            End Get
            Set(value As CompuMaster.Calendar.Month)
                _FirstMonth = value
                CheckForExchangeMonthValues()
            End Set
        End Property

        Private _LastMonth As CompuMaster.Calendar.Month
        Public Property LastMonth As CompuMaster.Calendar.Month
            Get
                Return _LastMonth
            End Get
            Set(value As CompuMaster.Calendar.Month)
                _LastMonth = value
                CheckForExchangeMonthValues()
            End Set
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

    End Class

End Namespace
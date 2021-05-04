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

    End Class

End Namespace
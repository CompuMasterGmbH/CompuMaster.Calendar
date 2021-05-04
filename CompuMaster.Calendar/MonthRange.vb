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
            Me._FirstPeriod = firstPeriod
            Me._LastPeriod = lastPeriod
            CheckForExchangeMonthValues()
        End Sub

        ''' <summary>
        ''' Exchange FirstPeriod and LastPeriod values if LastPeriod is before FirstPeriod
        ''' </summary>
        Private Sub CheckForExchangeMonthValues()
            If _FirstPeriod > _LastPeriod Then
                Dim BufferedFirstPeriod = _FirstPeriod
                Me._FirstPeriod = _LastPeriod
                Me._LastPeriod = BufferedFirstPeriod
            End If
        End Sub

        Private _FirstPeriod As CompuMaster.Calendar.Month
        Public Property FirstPeriod As CompuMaster.Calendar.Month
            Get
                Return _FirstPeriod
            End Get
            Set(value As CompuMaster.Calendar.Month)
                _FirstPeriod = value
                CheckForExchangeMonthValues()
            End Set
        End Property

        Private _LastPeriod As CompuMaster.Calendar.Month
        Public Property LastPeriod As CompuMaster.Calendar.Month
            Get
                Return _LastPeriod
            End Get
            Set(value As CompuMaster.Calendar.Month)
                _LastPeriod = value
                CheckForExchangeMonthValues()
            End Set
        End Property

        Public Overrides Function ToString() As String
            Return Me.FirstPeriod.ToString & " - " & Me.LastPeriod.ToString
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
            ElseIf Me.FirstPeriod < value.FirstPeriod Then
                Return -1
            ElseIf Me.FirstPeriod > value.FirstPeriod Then
                Return 1
            ElseIf Me.LastPeriod < value.LastPeriod Then
                Return -1
            ElseIf Me.LastPeriod > value.LastPeriod Then
                Return 1
            Else
                Return 0
            End If
        End Function

    End Class

End Namespace
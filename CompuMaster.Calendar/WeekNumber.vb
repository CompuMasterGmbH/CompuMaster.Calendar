Option Explicit On
Option Strict On

Namespace CompuMaster.Calendar

    ''' <summary>
    '''     A data structure for a complete week number inclusive the year information which is related to a pure week number
    ''' </summary>
    Public Class WeekNumber

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

        Public Overrides Function ToString() As String
            Return Me.Year & "/WK" & Me.Week
        End Function

        Public Shared Widening Operator CType(value As CompuMaster.Calendar.WeekNumber) As String
            If value Is Nothing Then
                Return Nothing
            Else
                Return value.ToString
            End If
        End Operator

    End Class

End Namespace
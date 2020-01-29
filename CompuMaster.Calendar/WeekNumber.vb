Option Explicit On
Option Strict On

Namespace CompuMaster.Calendar

    ''' <summary>
    '''     A data structure for a complete week number inclusive the year information which is related to a pure week number
    ''' </summary>
    Public Structure WeekNumber
        Dim Week As Integer
        Dim Year As Integer

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
    End Structure

End Namespace
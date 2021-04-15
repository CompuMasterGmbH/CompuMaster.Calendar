Option Explicit On
Option Strict On

Namespace CompuMaster.Calendar

    ''' <summary>
    ''' A structure providing delta information between 2 WeekNumber objects
    ''' </summary>
    Public Class WeekNumberSpan

        Public Sub New()
        End Sub

        Public Sub New(years As Integer, weeks As Integer)
            Me.Years = years
            Me.Weeks = weeks
        End Sub

        Public Property Years As Integer
        Public Property Weeks As Integer

        Public Overrides Function Equals(obj As Object) As Boolean
            If obj IsNot Nothing AndAlso obj.GetType IsNot Me.GetType Then
                Throw New ArgumentException(NameOf(obj))
            End If
            Dim tObj As WeekNumberSpan = CType(obj, WeekNumberSpan)
            If tObj.Years = Me.Years AndAlso tObj.Weeks = Me.Weeks Then
                Return True
            Else
                Return False
            End If
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return (Years * 100 + Weeks).GetHashCode
        End Function

        Public Shared Operator =(left As WeekNumberSpan, right As WeekNumberSpan) As Boolean
            Return left.Equals(right)
        End Operator

        Public Shared Operator <>(left As WeekNumberSpan, right As WeekNumberSpan) As Boolean
            Return Not left = right
        End Operator

    End Class

End Namespace
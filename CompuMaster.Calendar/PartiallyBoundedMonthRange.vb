Option Explicit On
Option Strict On

Namespace CompuMaster.Calendar

    ''' <summary>
    ''' A range of months with at least one inclusive boundary.
    ''' </summary>
    Public Class PartiallyBoundedMonthRange
        Inherits OptionallyBoundedMonthRange

        Public Sub New(firstMonth As Month, lastMonth As Month)
            MyBase.New(RequireBoundary(firstMonth, lastMonth), lastMonth)
        End Sub

        Private Shared Function RequireBoundary(firstMonth As Month, lastMonth As Month) As Month
            If firstMonth Is Nothing AndAlso lastMonth Is Nothing Then
                Throw New ArgumentException("At least one boundary must be specified")
            End If
            Return firstMonth
        End Function

        Public Shadows Shared Function Parse(value As String) As PartiallyBoundedMonthRange
            Dim FirstMonth As Month = Nothing
            Dim LastMonth As Month = Nothing
            ParseBoundaries(value, FirstMonth, LastMonth)
            Return New PartiallyBoundedMonthRange(FirstMonth, LastMonth)
        End Function

        Public Shadows Shared Function TryParse(value As String, ByRef result As PartiallyBoundedMonthRange) As Boolean
            Try
                result = Parse(value)
                Return True
            Catch ex As Exception When TypeOf ex Is ArgumentException OrElse TypeOf ex Is FormatException OrElse TypeOf ex Is OverflowException
                result = Nothing
                Return False
            End Try
        End Function

        Public Shadows Shared Widening Operator CType(value As MonthRange) As PartiallyBoundedMonthRange
            If value Is Nothing Then Return Nothing
            If value.IsEmpty Then Throw New InvalidCastException("MonthRange.Empty cannot be converted because the target type has no empty state")
            Return New PartiallyBoundedMonthRange(value.FirstMonth, value.LastMonth)
        End Operator

        Public Shadows Shared Narrowing Operator CType(value As PartiallyBoundedMonthRange) As MonthRange
            If value Is Nothing Then Return Nothing
            Return value.ToMonthRange()
        End Operator

    End Class

End Namespace

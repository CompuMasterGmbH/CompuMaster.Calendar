﻿Option Explicit On
Option Strict On

Namespace CompuMaster.Calendar

    ''' <summary>
    ''' A range of months
    ''' </summary>
    Public Class MonthRange
        Implements IComparable

        ''' <summary>
        ''' Create a new empty instance of MonthRange
        ''' </summary>
        Public Sub New()
            Me._IsEmpty = True
        End Sub

        ''' <summary>
        ''' Create a new instance of MonthRange with clones of first and last month (to behave more like value type instead of reference type)
        ''' </summary>
        ''' <param name="firstMonth"></param>
        ''' <param name="lastMonth"></param>
        Public Sub New(firstMonth As CompuMaster.Calendar.Month, lastMonth As CompuMaster.Calendar.Month)
            If firstMonth Is Nothing Then Throw New ArgumentNullException(NameOf(firstMonth))
            If lastMonth Is Nothing Then Throw New ArgumentNullException(NameOf(lastMonth))
            Me._FirstMonth = firstMonth.Clone
            Me._LastMonth = lastMonth.Clone
            If Me._FirstMonth > Me._LastMonth Then Throw New ArgumentException("First month must be before last month")
        End Sub

        ''' <summary>
        ''' Create a new instance of MonthRange with clones of first and last month (to behave more like value type instead of reference type)
        ''' </summary>
        ''' <param name="startYear"></param>
        ''' <param name="startMonth"></param>
        ''' <param name="endYear"></param>
        ''' <param name="endMonth"></param>
        Public Sub New(startYear As Integer, startMonth As Integer, endYear As Integer, endMonth As Integer)
            Me._FirstMonth = New Month(startYear, startMonth)
            Me._LastMonth = New Month(endYear, endMonth)
            If Me._FirstMonth > Me._LastMonth Then Throw New ArgumentException("Start month must be before last month")
        End Sub

        ''' <summary>
        ''' An empty MonthRange, containing no months
        ''' </summary>
        ''' <returns></returns>
        Public Shared ReadOnly Property Empty As MonthRange
            Get
                Return New MonthRange
            End Get
        End Property

        Private _IsEmpty As Boolean
        ''' <summary>
        ''' An empty range, containing no months
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property IsEmpty As Boolean
            Get
                Return _IsEmpty
            End Get
        End Property

        Private _FirstMonth As CompuMaster.Calendar.Month
        ''' <summary>
        ''' The first month of the range
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property FirstMonth As CompuMaster.Calendar.Month
            Get
                Return _FirstMonth
            End Get
        End Property

        Private _LastMonth As CompuMaster.Calendar.Month
        ''' <summary>
        ''' The last month of the range
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property LastMonth As CompuMaster.Calendar.Month
            Get
                Return _LastMonth
            End Get
        End Property

        ''' <summary>
        ''' A well-formed text representation of the range: "yyyy-MM - yyyy-MM"
        ''' </summary>
        ''' <returns></returns>
        Public Overrides Function ToString() As String
            If Me._IsEmpty Then
                Return ""
            Else
                Return Me.FirstMonth.ToString & " - " & Me.LastMonth.ToString
            End If
        End Function

        ''' <summary>
        ''' A range matching a whole year (January till December) is represented simply by the year number, all other ranges are represented by "yyyy-MM - yyyy-MM" (valid range) or an empty string (empty range)
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property SimplifiedName As String
            Get
                If Me._IsEmpty Then
                    Return ""
                ElseIf _FirstMonth.Year = _LastMonth.Year AndAlso _FirstMonth.Month = 1 AndAlso _LastMonth.Month = 12 Then
                    Return _FirstMonth.Year.ToString
                Else
                    Return Me.ToString
                End If
            End Get
        End Property

        ''' <summary>
        ''' Parse a text with format "yyyy-MM - yyyy-MM" (valid range) or an empty string (empty range)
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        Public Shared Function Parse(value As String) As MonthRange
            If value = Nothing Then
                Return MonthRange.Empty
            ElseIf value.Length <> 17 OrElse value.substring(7, 3) <> " - " Then
                'Throw New FormatException("Value must be formatted as ""yyyy-MM - yyyy-MM"" to parse successfully")
                Throw New FormatException("Value must be formatted as ""yyyy-MM - yyyy-MM"" to parse successfully, but found """ & value & """")
            Else
                Dim FirstMonth As String = value.Substring(0, 7)
                Dim LastMonth As String = value.Substring(10, 7)
                Return New MonthRange(Month.Parse(FirstMonth), Month.Parse(LastMonth))
            End If
        End Function

        ''' <summary>
        ''' Parse a text with format "yyyy-MM - yyyy-MM" (valid range) or an empty string (empty range)
        ''' </summary>
        ''' <param name="s"></param>
        ''' <param name="result"></param>
        ''' <returns></returns>
        Public Shared Function TryParse(s As String, ByRef result As MonthRange) As Boolean
            Try
                result = Parse(s)
                Return True
            Catch
                Return False
            End Try
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

            Return Me.FirstMonth = CType(obj, MonthRange).FirstMonth AndAlso Me.LastMonth = CType(obj, MonthRange).LastMonth AndAlso Me.IsEmpty = CType(obj, MonthRange).IsEmpty
        End Function

        ''' <summary>
        ''' Equals method for MonthRange classes
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Function Equals(ByVal value As MonthRange) As Boolean
            If value IsNot Nothing Then
                If Me.FirstMonth = value.FirstMonth AndAlso Me.LastMonth = value.LastMonth AndAlso Me.IsEmpty = value.IsEmpty Then
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
            ElseIf Me.IsEmpty And Not value.IsEmpty Then
                Return -1
            ElseIf Not Me.IsEmpty And value.IsEmpty Then
                Return 1
            ElseIf Me.IsEmpty And value.IsEmpty Then
                Return 0
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
                If Me.IsEmpty Then
                    Return 0
                Else
                    Return Me._LastMonth - Me._FirstMonth + 1
                End If
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
                    If Me.IsEmpty Then
                        _Months = New Month() {}
                    Else
                        Dim AllMonths As New List(Of Month)
                        Dim CurrentMonth As Month = Me._FirstMonth
                        Do While CurrentMonth <= Me._LastMonth
                            AllMonths.Add(CurrentMonth)
                            CurrentMonth = CurrentMonth.NextMonth
                        Loop
                        _Months = AllMonths.ToArray
                    End If
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
            If Me.IsEmpty Then
                Return False
            Else
                Return value >= Me.FirstMonth AndAlso value <= Me.LastMonth
            End If
        End Function

        ''' <summary>
        ''' Is a specified MonthRange a member of this range
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        Public Function Contains(value As MonthRange) As Boolean
            If value Is Nothing Then Throw New ArgumentNullException(NameOf(value))
            If Me.IsEmpty Then
                Return False
            Else
                Return value.FirstMonth >= Me.FirstMonth AndAlso value.LastMonth <= Me.LastMonth
            End If
        End Function

        ''' <summary>
        ''' Is there an overlapping of at least 1 month in both ranges
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        Public Function Overlaps(value As MonthRange) As Boolean
            If value Is Nothing Then Throw New ArgumentNullException(NameOf(value))
            If Me.IsEmpty Or value.IsEmpty Then
                Return False
            Else
                For Each month As Month In value.Months
                    If Me.Contains(month) Then Return True
                Next
                Return False
            End If
        End Function

        ''' <summary>
        ''' Move the whole MonthRange (with begin and end) into future (or in case of negative values: into past)
        ''' </summary>
        ''' <param name="years">Number of years to add</param>
        ''' <param name="months">Number of months to add</param>
        ''' <returns></returns>
        Public Function Add(years As Integer, months As Integer) As MonthRange
            If Me.IsEmpty Then Throw New InvalidOperationException("Can't add to MonthRange.Empty")
            Return New MonthRange(Me.FirstMonth.Add(years, months), Me.LastMonth.Add(years, months))
        End Function

        ''' <summary>
        ''' Move the whole MonthRange (with begin and end) into future (or in case of negative values: into past)
        ''' </summary>
        ''' <param name="value">Number of years to add</param>
        ''' <returns></returns>
        Public Function AddYears(value As Integer) As MonthRange
            If Me.IsEmpty Then Throw New InvalidOperationException("Can't add to MonthRange.Empty")
            Return New MonthRange(Me.FirstMonth.AddYears(value), Me.LastMonth.AddYears(value))
        End Function

        ''' <summary>
        ''' Move the whole MonthRange (with begin and end) into future (or in case of negative values: into past)
        ''' </summary>
        ''' <param name="value">Number of months to add</param>
        ''' <returns></returns>
        Public Function AddMonths(value As Integer) As MonthRange
            If Me.IsEmpty Then Throw New InvalidOperationException("Can't add to MonthRange.Empty")
            Return New MonthRange(Me.FirstMonth.AddMonths(value), Me.LastMonth.AddMonths(value))
        End Function

    End Class

End Namespace
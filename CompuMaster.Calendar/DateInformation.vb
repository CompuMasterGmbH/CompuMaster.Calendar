Option Explicit On 
Option Strict On

Namespace CompuMaster.Calendar

    Public NotInheritable Class DateInformation

#Region "Week functions"

        Public Shared Function LastDateOfWeekday(ByVal weekDay As System.DayOfWeek) As DateTime
            For MyCounter As Integer = 0 To 6
                If Now().AddDays(MyCounter * -1).DayOfWeek = weekDay Then
                    Return Now().AddDays(MyCounter * -1).Date()
                End If
            Next
            Throw New Exception("Unexpected position in operation workflow")
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Lookup the calendar week of the year based on the current culture set
        ''' </summary>
        ''' <param name="dateValue">A date value whose week number is required</param>
        ''' <returns>A week number</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[adminwezel]	31.10.2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Function WeekOfYear(ByVal dateValue As Date) As WeekNumber
            Return WeekOfYear(dateValue, System.Globalization.CultureInfo.CurrentCulture)
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Lookup the calendar week of the year based on the current culture set
        ''' </summary>
        ''' <param name="dateValue">A date value whose week number is required</param>
        ''' <param name="cultureInfo">The culture information which shall be the base of all date calculations</param>
        ''' <returns>A week number</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[adminwezel]	31.10.2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Function WeekOfYear(ByVal dateValue As Date, ByVal cultureInfo As System.Globalization.CultureInfo) As WeekNumber
            Dim Result As New WeekNumber
            Result.Week = cultureInfo.DateTimeFormat.Calendar.GetWeekOfYear(dateValue, System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.CalendarWeekRule, System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.FirstDayOfWeek)
            If dateValue.DayOfYear < 15 And Result.Week > 5 Then
                'Week is part of the old year numbering
                Result.Year = dateValue.Year - 1
            Else
                'Week is part of the current year
                Result.Year = dateValue.Year
            End If
            Return Result
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Lookup the first date of a week
        ''' </summary>
        ''' <param name="year">The year</param>
        ''' <param name="week">The week</param>
        ''' <returns>The date of the first day of the week</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[adminwezel]	31.10.2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Function FirstDateOfWeek(ByVal year As Integer, ByVal week As Integer) As Date
            Return FirstDateOfWeek(year, week, System.Globalization.CultureInfo.CurrentCulture)
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Lookup the first date of a week
        ''' </summary>
        ''' <param name="year">The year</param>
        ''' <param name="week">The week</param>
        ''' <param name="cultureInfo">The culture information which shall be the base of all date calculations</param>
        ''' <returns>The date of the first day of the week</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[adminwezel]	31.10.2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Function FirstDateOfWeek(ByVal year As Integer, ByVal week As Integer, ByVal cultureInfo As System.Globalization.CultureInfo) As Date

            'Parameter validation
            If year < DateTime.MinValue.Year Or year > DateTime.MaxValue.Year Then
                Throw New ArgumentException("Invalid year value: " & year, NameOf(year))
            End If
            If week < 1 Or week > 53 Then
                Throw New ArgumentException("Invalid week number: " & week, NameOf(week))
            End If

            Dim Result As Date

            'Walk through all days which might be in the area if the requested week
            Dim DaysPerWeek As Integer = cultureInfo.DateTimeFormat.DayNames.Length
            For MyDayOfYearCounter As Integer = week * DaysPerWeek - 2 * DaysPerWeek To week * DaysPerWeek + 2 * DaysPerWeek
                Dim DayOfYear As Date = New Date(year, 1, 1, cultureInfo.DateTimeFormat.Calendar).AddDays(MyDayOfYearCounter - 1)
                Dim WeekOfDay As CompuMaster.Calendar.WeekNumber = WeekOfYear(DayOfYear)
                If WeekOfDay.Week = week And WeekOfDay.Year = year Then
                    'This is a match, save the value
                    Result = DayOfYear
                    Exit For
                End If
            Next

            If Result = Nothing Then
                Throw New Exception("Date resolution failed")
            End If

            Return Result

        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Lookup the last date of a week
        ''' </summary>
        ''' <param name="year">The year</param>
        ''' <param name="week">The week</param>
        ''' <returns>The date of the last day of the week on 0 o'clock (not 24 o'clock!)</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[adminwezel]	31.10.2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Function LastDateOfWeek(ByVal year As Integer, ByVal week As Integer) As Date
            Return LastDateOfWeek(year, week, System.Globalization.CultureInfo.CurrentCulture)
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Lookup the last date of a week
        ''' </summary>
        ''' <param name="year">The year</param>
        ''' <param name="week">The week</param>
        ''' <param name="cultureInfo">The culture information which shall be the base of all date calculations</param>
        ''' <returns>The date of the last day of the week on 0 o'clock (not 24 o'clock!)</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[adminwezel]	31.10.2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Function LastDateOfWeek(ByVal year As Integer, ByVal week As Integer, ByVal cultureInfo As System.Globalization.CultureInfo) As Date

            'Parameter validation
            If year < DateTime.MinValue.Year Or year > DateTime.MaxValue.Year Then
                Throw New ArgumentException("Invalid year value: " & year, NameOf(year))
            End If
            If week < 1 Or week > 53 Then
                Throw New ArgumentException("Invalid week number: " & week, NameOf(week))
            End If

            Dim Result As Date

            'Walk through all days which might be in the area if the requested week
            Dim DaysPerWeek As Integer = cultureInfo.DateTimeFormat.DayNames.Length
            For MyDayOfYearCounter As Integer = week * DaysPerWeek - 2 * DaysPerWeek To week * DaysPerWeek + 2 * DaysPerWeek
                Dim DayOfYear As Date = New Date(year, 1, 1).AddDays(MyDayOfYearCounter - 1)
                Dim WeekOfDay As CompuMaster.Calendar.WeekNumber = WeekOfYear(DayOfYear)
                If WeekOfDay.Week = week AndAlso WeekOfDay.Year = year Then
                    'This is a match, save the value
                    Result = DayOfYear
                ElseIf (WeekOfDay.Week > week AndAlso WeekOfDay.Year = year) OrElse WeekOfDay.Year > year Then
                    'Days are above the requested week, we can leave the for-loop now - except we're looking at the beginning of a year (because the first days might be week 52 or 53, but related to the old year
                    Exit For
                End If
            Next

            If Result = Nothing Then
                Throw New Exception("Date resolution failed")
            End If

            Return Result

        End Function

#End Region

#Region "Mathematicals"

        Public Shared Function Min(ByVal value1 As System.DateTime, ByVal value2 As System.DateTime) As System.DateTime
            If value1 < value2 Then
                Return value1
            Else
                Return value2
            End If
        End Function
        Public Shared Function Min(ByVal value1 As TimeSpan, ByVal value2 As TimeSpan) As TimeSpan
            If TimeSpan.Compare(value1, value2) < 0 Then
                Return value1
            Else
                Return value2
            End If
        End Function
        Public Shared Function Max(ByVal value1 As System.DateTime, ByVal value2 As System.DateTime) As System.DateTime
            If value1 > value2 Then
                Return value1
            Else
                Return value2
            End If
        End Function
        Public Shared Function Max(ByVal value1 As TimeSpan, ByVal value2 As TimeSpan) As TimeSpan
            If TimeSpan.Compare(value1, value2) > 0 Then
                Return value1
            Else
                Return value2
            End If
        End Function

#End Region

#Region "Period begin and end date/time"

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' The precision of a datetime representation which may have got an influence on rounded values
        ''' </summary>
        ''' <remarks>
        ''' Some database or application systems don't fully support the high precision of the .NET framework's datetime value. In this case, those systems may round the datetime value to store/handle the value. But you may want to prevent roundings to the next day.
        ''' <example>You want to filter for records till the end of a day, but don't want records from the next day with time 00:00:00 because of any roundings.</example>
        ''' </remarks>
        ''' <history>
        ''' 	[wezel]	13.12.2010	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Enum Accuracy As Byte
            MilliSecond = 0
            Second = 1
            MsSqlServer = 2
        End Enum

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' The first moment of a year
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[wezel]	13.12.2010	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Function BeginOfYear(ByVal value As Date) As Date
            BeginOfYear = DateSerial(value.Year, 1, 1)
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' The last moment of a year
        ''' </summary>
        ''' <param name="value"></param>
        ''' <param name="precision"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[wezel]	13.12.2010	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Function EndOfYear(ByVal value As Date, ByVal precision As Accuracy) As Date
            If precision = Accuracy.MilliSecond Then
                Return BeginOfYear(value).AddYears(1).AddMilliseconds(-1)
            ElseIf precision = Accuracy.MsSqlServer Then
                Return BeginOfYear(value).AddYears(1).AddMilliseconds(-3)
            ElseIf precision = Accuracy.Second Then
                Return BeginOfYear(value).AddYears(1).AddSeconds(-1)
            Else
                Throw New ArgumentOutOfRangeException(NameOf(precision))
            End If
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' The first moment of a month
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[wezel]	13.12.2010	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Function BeginOfMonth(ByVal value As Date) As Date
            BeginOfMonth = DateSerial(Year(value), Microsoft.VisualBasic.Month(value), 1)
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' The last moment of a month
        ''' </summary>
        ''' <param name="value"></param>
        ''' <param name="precision"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[wezel]	13.12.2010	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Function EndOfMonth(ByVal value As Date, ByVal precision As Accuracy) As Date
            If precision = Accuracy.MilliSecond Then
                Return BeginOfMonth(value).AddMonths(1).AddMilliseconds(-1)
            ElseIf precision = Accuracy.MsSqlServer Then
                Return BeginOfMonth(value).AddMonths(1).AddMilliseconds(-3)
            ElseIf precision = Accuracy.Second Then
                Return BeginOfMonth(value).AddMonths(1).AddSeconds(-1)
            Else
                Throw New ArgumentOutOfRangeException(NameOf(precision))
            End If
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' The first moment of a day
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[wezel]	13.12.2010	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Function BeginOfDay(ByVal value As Date) As Date
            BeginOfDay = DateSerial(Year(value), Microsoft.VisualBasic.Month(value), Day(value))
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' The last moment of a day
        ''' </summary>
        ''' <param name="value"></param>
        ''' <param name="precision"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[wezel]	13.12.2010	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Function EndOfDay(ByVal value As Date, ByVal precision As Accuracy) As Date
            If precision = Accuracy.MilliSecond Then
                Return BeginOfDay(value).AddDays(1).AddMilliseconds(-1)
            ElseIf precision = Accuracy.MsSqlServer Then
                Return BeginOfDay(value).AddDays(1).AddMilliseconds(-3)
            ElseIf precision = Accuracy.Second Then
                Return BeginOfDay(value).AddDays(1).AddSeconds(-1)
            Else
                Throw New ArgumentOutOfRangeException(NameOf(precision))
            End If
        End Function
#End Region

    End Class

End Namespace
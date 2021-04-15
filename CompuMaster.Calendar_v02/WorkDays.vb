Option Explicit On 
Option Strict On

Namespace CompuMaster.Calendar

    ''' <summary>
    ''' Provides several methods for working with work days
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[wezel]	13.12.2010	Created
    ''' </history>
    Public NotInheritable Class WorkDays

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Count the number of weekdays Monday to Friday, also known as workdays
        ''' </summary>
        ''' <param name="fromDate">The begin of the period</param>
        ''' <param name="toDate">The last inclusive day of the period</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[wezel]	10.02.2009	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Function CountTypicalWorkDays(ByVal fromDate As Date, ByVal toDate As Date) As Integer
            Return CountWeekDays(fromDate, toDate, False, True, True, True, True, True, False)
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Count the number of weekdays Saturday to Sunday, also known as free days
        ''' </summary>
        ''' <param name="fromDate">The begin of the period</param>
        ''' <param name="toDate">The last inclusive day of the period</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[wezel]	10.02.2009	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Function CountTypicalFreeDays(ByVal fromDate As Date, ByVal toDate As Date) As Integer
            Return CountWeekDays(fromDate, toDate, True, False, False, False, False, False, True)
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Count the number of weekdays
        ''' </summary>
        ''' <param name="fromDate">The begin of the period</param>
        ''' <param name="toDate">The last inclusive day of the period</param>
        ''' <param name="countSundays">True if Sundays shall be counted</param>
        ''' <param name="countMondays">True if Monday shall be counted</param>
        ''' <param name="countTuesdays">True if Tuesdays shall be counted</param>
        ''' <param name="countWednesdays">True if Wednesdays shall be counted</param>
        ''' <param name="countThursdays">True if Thursdays shall be counted</param>
        ''' <param name="countFridays">True if Fridays shall be counted</param>
        ''' <param name="countSaturdays">True if Saturdays shall be counted</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[wezel]	10.02.2009	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Function CountWeekDays(ByVal fromDate As Date, ByVal toDate As Date, ByVal countSundays As Boolean, ByVal countMondays As Boolean, ByVal countTuesdays As Boolean, ByVal countWednesdays As Boolean, ByVal countThursdays As Boolean, ByVal countFridays As Boolean, ByVal countSaturdays As Boolean) As Integer
            Dim Result As Integer = 0
            For MyCounter As Integer = 0 To toDate.Date.Subtract(fromDate.Date).Days
                Select Case fromDate.AddDays(MyCounter).DayOfWeek
                    Case DayOfWeek.Sunday
                        If countSundays Then Result += 1
                    Case DayOfWeek.Monday
                        If countMondays Then Result += 1
                    Case DayOfWeek.Tuesday
                        If countTuesdays Then Result += 1
                    Case DayOfWeek.Wednesday
                        If countWednesdays Then Result += 1
                    Case DayOfWeek.Thursday
                        If countThursdays Then Result += 1
                    Case DayOfWeek.Friday
                        If countFridays Then Result += 1
                    Case DayOfWeek.Saturday
                        If countSaturdays Then Result += 1
                End Select
            Next
            Return Result
        End Function

    End Class

End Namespace
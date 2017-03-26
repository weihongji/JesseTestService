Public Class FtpChecker
	Private StartTime As TimeSpan
	Private EndTime As TimeSpan
	Private Interval As Integer
	Private LastMatchTime As DateTime

	Public Sub New()
		Clear()
	End Sub

	Private Sub Clear()
		StartTime = Nothing
		EndTime = Nothing
		Interval = 0
		LastMatchTime = Nothing
	End Sub

	Public Function Initialize(ByVal start As TimeSpan, ByVal [end] As TimeSpan, ByVal interval As Integer) As String
		Dim result As String = InitializeWork(start, [end], interval)
		If Not String.IsNullOrEmpty(result) Then
			Clear()
		End If
		Return result
	End Function

	Public Function InitializeWork(ByVal start As TimeSpan, ByVal [end] As TimeSpan, ByVal interval As Integer) As String
		If start > [end] Then
			'23:50 - 0:10
			Return "Failed to initialize checker. Checker start time shouldn't be greater than end time. Considering sliptting into two checkers."
			'Default checker: 23:50 - 23:59
			'Another checker:  0:00 -  0: 10
		End If

		If interval <= 0 Then
			Return "Failed to initialize checker. Invalid interval value: " + interval.ToString()
		End If

		Dim changed As Boolean = False

		If Not Me.StartTime.Equals(start) Then
			Me.StartTime = start
			changed = True
		End If
		If Not Me.EndTime.Equals([end]) Then
			Me.EndTime = [end]
			changed = True
		End If
		If Not Me.Interval.Equals(interval) Then
			Me.Interval = interval
			changed = True
		End If

		If changed Then
			LastMatchTime = Nothing
		End If

		Return ""
	End Function

	Public Function IsMatched() As Boolean
		Return IsMatched(Now)
	End Function

	Public Function IsMatched(ByVal dt As DateTime) As Boolean
		If Interval = 0 Then
			Return False
		End If
		If dt.Year <= 2000 OrElse Today.Year <= 2000 Then
			Return False
		End If
		If dt <= Me.LastMatchTime Then
			Return False
		End If

		Dim time As TimeSpan = New TimeSpan(dt.Hour, dt.Minute, dt.Second)
		If Me.StartTime <= time AndAlso time <= Me.EndTime Then	'Time matched
			If Me.LastMatchTime.Year <= 2000 Then 'First match after initialized
				Me.LastMatchTime = dt
				Return True
			Else 'Ever matched before. Need to check interval between neighbor matches.
				Dim span As Integer = Math.Round((dt - Me.LastMatchTime).TotalMinutes)
				If span >= Me.Interval Then
					Me.LastMatchTime = dt
					Return True
				End If
			End If
		End If

		Return False
	End Function
End Class

Imports System.Runtime.InteropServices
Imports System.Threading
Imports Tamir.SharpSsh

Public Enum ServiceState
	SERVICE_STOPPED = 1
	SERVICE_START_PENDING = 2
	SERVICE_STOP_PENDING = 3
	SERVICE_RUNNING = 4
	SERVICE_CONTINUE_PENDING = 5
	SERVICE_PAUSE_PENDING = 6
	SERVICE_PAUSED = 7
End Enum

<StructLayout(LayoutKind.Sequential)> Public Structure ServiceStatus
	Public dwServiceType As Long
	Public dwCurrentState As ServiceState
	Public dwControlsAccepted As Long
	Public dwWin32ExitCode As Long
	Public dwServiceSpecificExitCode As Long
	Public dwCheckPoint As Long
	Public dwWaitHint As Long
End Structure

Public Class JesseTestService
	Declare Auto Function SetServiceStatus Lib "advapi32.dll" (ByVal handle As IntPtr, ByRef serviceStatus As ServiceStatus) As Boolean

	Dim m_writer As TextFile
	Dim m_checker As FtpChecker
	Dim m_ftp_max_per_day As Integer
	Dim m_ftp_last_day As DateTime
	Dim m_ftp_last_day_count As Integer

	Public Sub New()
		' Update the service state to Start Pending.  
		Dim serviceStatus As ServiceStatus = New ServiceStatus()
		serviceStatus.dwCurrentState = ServiceState.SERVICE_START_PENDING
		serviceStatus.dwWaitHint = 100000
		SetServiceStatus(Me.ServiceHandle, serviceStatus)

		' This call is required by the Windows Form Designer.
		InitializeComponent()

		'Update the service state to Running
		serviceStatus.dwCurrentState = ServiceState.SERVICE_RUNNING
		SetServiceStatus(Me.ServiceHandle, serviceStatus)
	End Sub

	Protected Overrides Sub OnStart(ByVal args() As String)
		' Add code here to start your service. This method should set things
		' in motion so your service can do its work.
		'Thread.Sleep(1000 * 20)


		Dim ini As IniFile = GetIniFile()
		m_writer = New TextFile(System.AppDomain.CurrentDomain.BaseDirectory + "test.log")
		m_checker = New FtpChecker()

		m_ftp_max_per_day = Integer.Parse(ini.ReadString("service", "max_ftp_per_day", 50))

		Dim timer As System.Timers.Timer = New System.Timers.Timer()
		timer.Interval = Double.Parse(ini.ReadString("service", "interval", 1)) * 60000	' 60 seconds
		AddHandler timer.Elapsed, AddressOf Me.OnTimer
		timer.Start()

		m_writer.WriteLn("")
		m_writer.WriteLn("Service started ******************************")
		m_writer.WriteLn("")
	End Sub

	Private Function GetIniFile() As IniFile
		Dim ini As IniFile = New IniFile()
		ini.FileName = System.AppDomain.CurrentDomain.BaseDirectory + "config.ini"
		Return ini
	End Function

	Protected Overrides Sub OnStop()
		m_writer.WriteLn("Service closed #################################")
		m_writer.WriteLn("")
		m_writer.CloseFile()
	End Sub

	Private Sub OnTimer(ByVal sender As Object, ByVal e As Timers.ElapsedEventArgs)
		Dim t As Date = Now
		If t.Minute = 0 Then
			m_writer.WriteLn("=================================")
		End If

		SetChecker()

		If m_checker.IsMatched(t) Then
			ConnectFTP()
		End If
	End Sub

	Private Sub SetChecker()
		Dim ini As IniFile = GetIniFile()

		Dim start As TimeSpan = TimeSpan.Parse(ini.ReadString("checker", "start", "4:00"))
		Dim [end] As TimeSpan = TimeSpan.Parse(ini.ReadString("checker", "end", "4:15"))
		Dim interval As Integer = Integer.Parse(ini.ReadString("checker", "interval", 1))

		Dim result As String = m_checker.Initialize(start, [end], interval)
		If Not String.IsNullOrEmpty(result) Then
			m_writer.WriteLn(result)
		End If
	End Sub

	Private Sub ConnectFTP()
		If m_ftp_last_day < Today Then
			m_ftp_last_day_count = 0
			m_ftp_last_day = Today
		End If
		m_ftp_last_day_count += 1
		If m_ftp_last_day_count > m_ftp_max_per_day Then
			Return
		End If

		Dim ini As IniFile = GetIniFile()
		Dim ftp_address As String = ini.ReadString("ftp", "host")	'"sftp://sftp.ymca-kc.org"
		Dim ftp_port As Integer = ini.ReadString("ftp", "port")	'"64022"
		Dim ftp_user_name As String = ini.ReadString("ftp", "user")	'"activenet"
		Dim ftp_user_password As String = ini.ReadString("ftp", "password")	'"EnigmaX.sf1"

		' remove protocal, we dont need it
		If ftp_address.ToLower.StartsWith("sftp://") Then
			ftp_address = ftp_address.Substring(7)
		End If

		' check sub folder if any
		Dim subFolder As String = ""
		Dim index As Integer = ftp_address.IndexOf("/")
		If index >= 0 Then
			subFolder = ftp_address.Substring(index)
			ftp_address = ftp_address.Remove(index, subFolder.Length)
			subFolder = subFolder.Substring(1)
			If subFolder.EndsWith("/") Then
				subFolder = subFolder.Remove(subFolder.Length - 1, 1)
			End If
		End If

		Dim ftp As Sftp = New Sftp(ftp_address, ftp_user_name, ftp_user_password)

		' we need a port
		If ftp_port <= 0 Then
			' try to get it from address
			index = ftp_address.IndexOf(":")
			If index >= 0 AndAlso index < ftp_address.Length - 1 Then
				ftp_port = Integer.Parse(ftp_address.Substring(index + 1))
				ftp_address = ftp_address.Substring(index)
			End If
			' still not able to get it, set 21 as default
			If ftp_port <= 0 Then
				ftp_port = 21
			End If
		End If
		Try
			m_writer.WriteLn(String.Format("Connecting FTP ({0}, {1}, {2}, {3}) ...", ftp_address, ftp_port, ftp_user_name, ftp_user_password))
			ftp.Connect(ftp_port)
			m_writer.WriteLn("FTP Connected.")
			ftp.Close()
			m_writer.WriteLn("FTP Closed.")
		Catch ex As Exception
			m_writer.WriteLn(String.Format("Failed: {1}{0}{2}{0}{3}", vbCrLf, ex.Message(), ex.ToString(), ex.StackTrace))
		End Try
		m_writer.WriteLn("---------------------------------")
	End Sub
End Class

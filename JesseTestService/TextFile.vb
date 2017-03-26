Option Strict On
Option Explicit On

Imports System.IO
Imports System.Threading

Public Class TextFile

    Public Enum enmTextFileMode
        enmTfmReadOnly = 0
        enmTfmWrite = 1
        enmTfmAppend = 2
    End Enum

    Dim mstrFileName As String ' File number
    Dim mintFileNum As Integer ' File number for log file; zero means closed
    Dim menmTextFileMode As enmTextFileMode ' Current mode

    '****************************************************************************
    ' Initialize and terminate events -- That's just about the whole thing
    '****************************************************************************

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal vstrFile$)
        Me.FileName = vstrFile
    End Sub

    Public Sub New(ByVal vstrFile$, ByVal vblnOpenAsShared As Boolean)
        Me.New(vstrFile)
    End Sub

    Protected Overrides Sub Finalize()
        ' Close file just in case never explicitly closed
        If mintFileNum > 0 Then
            CloseFile()
        End If
        MyBase.Finalize()
    End Sub

    '****************************************************************************
    ' Properties
    '****************************************************************************

    Public Property FileName() As String
        Get
            FileName = mstrFileName
        End Get
        Set(ByVal Value As String)

            ' If another file is open, close it first
            Value = Trim(Value)
            If mintFileNum > 0 And UCase(mstrFileName) <> UCase(Value) Then
                CloseFile()
            End If

            mstrFileName = Value
        End Set
    End Property

    '****************************************************************************
    ' Methods
    '****************************************************************************

    ' Attempt to delete file if it exists
    Public Function DeleteFile() As Boolean

        ' If file doesn't exist, that's cool
        If Dir(mstrFileName) = "" Then
            DeleteFile = True

            ' Otherwise, try to delete
        Else
            Try
                Kill(mstrFileName)
                DeleteFile = True
            Catch ex As Exception
                DeleteFile = False
            End Try
        End If
    End Function

    Public Function RenameFile(ByVal vstrNewName As String) As Boolean

        ' Attempt to rename
        Try
            Rename(mstrFileName, vstrNewName)
            mstrFileName = vstrNewName
            RenameFile = True
        Catch ex As Exception
            RenameFile = False
        End Try

    End Function

    ' Whether file exists
    Public Function Exists() As Boolean
        Try
            Exists = Len(Dir(FileName)) > 0
        Catch
            'do nothing
        End Try
    End Function

    Public Function OpenFile(ByVal vblnForAppend As Boolean) As Boolean
        OpenFile = OpenFileMode(CType(IIf(vblnForAppend, enmTextFileMode.enmTfmAppend, enmTextFileMode.enmTfmReadOnly), enmTextFileMode))
    End Function

    Public Function OpenFileMode(ByVal venmTextFileMode As enmTextFileMode) As Boolean

        ' If already open, success if mode is same, else failure
        If mintFileNum > 0 Then
            OpenFileMode = (venmTextFileMode = menmTextFileMode)
            Exit Function
        End If

        ' Get file handle
        Try
            mintFileNum = FreeFile()
        Catch
            'do nothing
        End Try
        If mintFileNum <= 0 Then
            If Not OpenFileMode Then mintFileNum = 0
            Exit Function
        End If

        ' Attempt to open; return success
        menmTextFileMode = venmTextFileMode
        Select Case menmTextFileMode
            Case enmTextFileMode.enmTfmReadOnly
                Try
                    FileOpen(mintFileNum, FileName, OpenMode.Input, OpenAccess.Default, OpenShare.Shared)
                    OpenFileMode = True
                Catch
                    OpenFileMode = False
                End Try
            Case enmTextFileMode.enmTfmWrite
                Try
                    FileOpen(mintFileNum, FileName, OpenMode.Output, OpenAccess.Default, OpenShare.Shared)
                    OpenFileMode = True
                Catch
                    OpenFileMode = False
                End Try

            Case enmTextFileMode.enmTfmAppend
                Try
                    FileOpen(mintFileNum, FileName, OpenMode.Append, OpenAccess.Default, OpenShare.Shared)
                    OpenFileMode = True
                Catch
                    OpenFileMode = False
                End Try

        End Select



    End Function

    ' Write a certain number of line feeds
    Public Sub WriteLn(ByVal iLines As Integer)
        While iLines > 0
            WriteLn()
            iLines -= 1
        End While
    End Sub

    Public Function WriteLn(ByVal oStream As Stream) As Boolean
        Return WriteLn(TextFromStream(oStream))
    End Function

    ' Write a line to the file
    Public Function WriteLn(Optional ByVal vstrText As String = "") As Boolean
        Return WriteText(vstrText & vbCrLf)
    End Function

    ' Append to the file from a stream
    Public Function WriteText(ByVal oStream As Stream) As Boolean
        Return WriteText(TextFromStream(oStream))
    End Function

    Private Function TextFromStream(ByVal oStream As Stream) As String
        Dim reader As New StreamReader(oStream)
        oStream.Position = 0
        Return reader.ReadToEnd()
    End Function

    ' Append the text to the file
    Private Function WriteText(ByVal vstrText As String) As Boolean
        Dim bEnterOK As Boolean
        bEnterOK = Monitor.TryEnter(Me, 5000)
        If bEnterOK Then
            Try
                ' Make sure file is open; try to open otherwise
                If mintFileNum <= 0 Then
                    If Not Me.OpenFile(True) Then
                        Exit Function
                    End If
                End If
                ' Attempt to write text
				Try
					vstrText = String.Format("{0} - {1}", Now.ToString("yyyy-MM-dd HH:mm:ss"), vstrText)
					Print(mintFileNum, vstrText)
					WriteText = True
				Catch
					WriteText = False
				Finally
					Me.CloseFile()
				End Try
            Finally
                Monitor.Exit(Me)
            End Try
        End If
    End Function

    ' Read the full file
    Public Function ReadFile(ByRef vstrText As String) As Boolean
        Dim pstrChar As String
        Dim pstrTemp As String = ""

        ' Make sure file is open
        If mintFileNum <= 0 Then Exit Function
        ' Try to read file character at a time
        vstrText = ""

        While True
            Try
                pstrChar = InputString(mintFileNum, 1)
                pstrTemp = pstrTemp & pstrChar
                If Len(pstrTemp) = 1024 Then
                    vstrText = vstrText & pstrTemp
                    pstrTemp = ""
                End If

            Catch e As System.IO.EndOfStreamException
                vstrText = vstrText & pstrTemp
                ReadFile = True
                Exit Function
            Catch
                vstrText = vstrText & pstrTemp
                ReadFile = False
                Exit Function
            End Try

        End While
    End Function

    ' Read a line of the file
    Public Function ReadLn(ByRef vstrText As String) As Boolean
        ' Make sure file is open
        If mintFileNum <= 0 Then Exit Function
        ' Try to read line
        Try
            vstrText = LineInput(mintFileNum)
            ReadLn = True
        Catch
            ReadLn = False
        End Try
    End Function

    ' Try to close the file
    Public Function CloseFile() As Boolean
        ' If already closed, just return true
        If mintFileNum <= 0 Then
            CloseFile = True
            Exit Function
        End If

        ' Attempt to close
        Try
            FileClose(mintFileNum)
            CloseFile = True
            mintFileNum = 0
        Catch
            CloseFile = False
        End Try
    End Function
End Class

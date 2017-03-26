Option Strict On
Option Explicit On

Imports VB6 = Microsoft.VisualBasic.Compatibility.VB6

Public Class IniFile

    Const ModuleName As String = "IniFile"

    Dim mstrFileName As String ' Name of .INI file
    Dim mcolEntries As Collection ' Collection of cached values

    ' ===========================================================================
    ' Initialize and terminate events
    ' ===========================================================================

    Public Sub New()
        MyBase.New()
        mcolEntries = New Collection
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    ' ===========================================================================
    ' Properties
    ' ===========================================================================


    Public Property FileName() As String
        Get
            FileName = mstrFileName
        End Get
        Set(ByVal Value As String)
            ' Clear collection if filename is changed
            Value = UCase(Trim(Value))
            If mstrFileName <> Value Then
                mcolEntries = New Collection
            End If
            mstrFileName = Value
        End Set
    End Property

    ' Property to read .INI file, using cached values if already present
    Public Function ReadString(ByVal vstrSection As String, ByVal vstrKeyword As String, Optional ByVal vstrDefault As String = "") As String
        Dim pstrKey As String = ""
        Dim pstrValue As String = ""

        ' Create collection key
        pstrKey = Trim(vstrSection) & "." & Trim(vstrKeyword)

        ' First look in collection
        Try
            pstrValue = CStr(mcolEntries.Item(pstrKey))
        Catch
            Try
                ' If not in collection, try to read
                pstrValue = sReadIni(mstrFileName, vstrSection, vstrKeyword, vstrDefault)
            Catch
                'do nothing
            End Try
            Try
                ' Add to collection
                mcolEntries.Add(pstrValue, pstrKey)
            Catch
                'do nothing
            End Try
        Finally
            ReadString = pstrValue
        End Try


    End Function

    ' Write .INI file, and update cache
    Public Function WriteString(ByVal vstrSection As String, ByVal vstrKeyword As String, _
        ByVal RHS As String) As Boolean
        Dim pstrKey As String
        If WriteIni(mstrFileName, vstrSection, vstrKeyword, RHS) Then
            pstrKey = Trim(vstrSection) & "." & Trim(vstrKeyword)

            Try
                'Let's remove this from cache if it exists
                mcolEntries.Remove(pstrKey)
            Catch ex As Exception
                'Looks like that entry did not exist in the collection so do nothing
            End Try

            'Now add it
            mcolEntries.Add(RHS, pstrKey)
            WriteString = True
        Else
            'Write ini failed so the value never got written
            WriteString = False
        End If

    End Function

    ' Return an entry, assuming its a logical
    Public Function ReadBoolean(ByVal vstrSection As String, ByVal vstrKeyword As String, Optional ByVal vblnDefault As Boolean = False) As Boolean
        Try
            ReadBoolean = CBool(ReadString(vstrSection, vstrKeyword, CStr(IIf(vblnDefault, "1", "0"))))
        Catch
            'do nothing
        End Try
    End Function

    ' Return an entry, assuming its a logical
    Public Function ReadLong(ByVal vstrSection As String, ByVal vstrKeyword As String, Optional ByVal vlngDefault As Integer = 0) As Integer
        Try
            ReadLong = CInt(ReadString(vstrSection, vstrKeyword, CStr(vlngDefault)))
        Catch
            'do nothing
        End Try
	End Function

#Region "Ini basic functions"
	' Read a boolean value from an .INI file; note that canonical format is =0 for false,
	' =1 for true, although any non-zero number will be converted to true
	Public Function bReadIni(ByVal sIniFile As String, ByVal sSection As String, ByVal sKeyword As String, Optional ByVal bDefault As Boolean = False) As Boolean
		Dim sReturn As String
		sReturn = sReadIni(sIniFile, sSection, sKeyword, "")
		If sReturn = "" Then
			bReadIni = bDefault
		Else
			' this makes non-numeric data count as false
			Try
				bReadIni = CBool(sReturn)
			Catch
				bReadIni = False
			End Try
		End If
	End Function

	Public Function sReadIni(ByVal sIniFile As String, ByVal sSection As String, ByVal sKeyword As String, Optional ByVal sDefault As String = "") As String
		Dim sReturn As New VB6.FixedLengthString(1000)
		Dim lCount As Integer
		lCount = GetPrivateProfileString(sSection, sKeyword, Trim(sDefault), sReturn.Value, 1000, sIniFile)
		If lCount = 0 Then
			sReadIni = sDefault
		Else
			sReadIni = Left(sReturn.Value, lCount)
		End If
	End Function

	' Write an INI file entry
	Public Function WriteIni(ByVal sIniFile As String, ByVal sSection As String, ByVal sKeyword As String, _
	  ByVal vValue As String) As Boolean

		Dim pstrValue As String
		pstrValue = CStr(vValue)

		WriteIni = WritePrivateProfileString(sSection, sKeyword, vValue, sIniFile)

	End Function
#End Region
End Class
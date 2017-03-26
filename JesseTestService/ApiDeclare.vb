Option Strict On
Option Explicit On 

Public Module ApiDeclare


    ''-------------------------------
    ''   API Type Definitions
    ''-------------------------------

    'Public Structure SYSTEM_INFO
    '	Dim dwOemID As Integer
    '	Dim dwPageSize As Integer
    '	Dim lpMinimumApplicationAddress As Integer
    '	Dim lpMaximumApplicationAddress As Integer
    '	Dim dwActiveProcessorMask As Integer
    '	Dim dwNumberOfProcessors As Integer
    '	Dim dwProcessorType As Integer
    '	Dim dwAllocationGranularity As Integer
    '	Dim wProcessorLevel As Short
    '	Dim wProcessorRevision As Short
    'End Structure

    'Public Structure OSVERSIONINFO ' 148 bytes
    '	Dim dwOSVersionInfoSize As Integer
    '	Dim dwMajorVersion As Integer
    '	Dim dwMinorVersion As Integer
    '	Dim dwBuildNumber As Integer
    '	Dim dwPlatformId As Integer
    '	<VBFixedString(128),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr,SizeConst:=128)> Public szCSDVersion As String
    'End Structure

    'Structure FILETIME
    '	Dim dwLowDateTime As Integer
    '	Dim dwHighDateTime As Integer
    'End Structure

    'Public Const MAX_DEFAULTCHAR As Short = 2
    'Public Const MAX_LEADBYTES As Short = 12

    'Structure CPINFO
    '	Dim MaxCharSize As Integer '  max length (Byte) of a char
    '	<VBFixedArray(MAX_DEFAULTCHAR)> Dim DefaultChar() As Byte '  default character
    '	<VBFixedArray(MAX_LEADBYTES)> Dim LeadByte() As Byte '  lead byte ranges

    '	'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1026"'
    '	Public Sub Initialize()
    '		ReDim DefaultChar(MAX_DEFAULTCHAR)
    '		ReDim LeadByte(MAX_LEADBYTES)
    '	End Sub
    'End Structure

    ''user defined type required by Shell_NotifyIcon API call
    'Public Structure NOTIFYICONDATA
    '	Dim cbSize As Integer
    '	Dim hwnd As Integer
    '	Dim uId As Integer
    '	Dim uFlags As Integer
    '	Dim uCallBackMessage As Integer
    '	Dim hIcon As Integer
    '	<VBFixedString(64),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr,SizeConst:=64)> Public szTip As String
    'End Structure

    'Public Const KL_NAMELENGTH As Short = 9

    'Public Const MAX_PATH As Short = 260

    ''-------------------------------
    ''   API Declares -- Alphabetically organized
    ''-------------------------------

    Declare Function BringWindowToTop Lib "user32" Alias "BringWindowToTop" (ByVal hwnd As Integer) As Integer

    Declare Function Beep Lib "kernel32" (ByVal dwFreq As Integer, ByVal dwDuration As Integer) As Integer

    ''TODO: Determine reason for different declaration
    ''Declare Sub CopyMemoryDave Lib "kernel32" Alias "RtlMoveMemory" _
    '''    (dest As Any, Source As Any, ByVal numBytes As Long)

    Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Integer) As Integer

    ''UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1016"'
    'Declare Sub CopyMemory Lib "kernel32"  Alias "RtlMoveMemory"(ByRef hpvDest As Any, ByVal hpvSource As Integer, ByVal cbCopy As Integer)

    'Declare Function CreateCaret Lib "user32" (ByVal hwnd As Integer, ByVal hBitmap As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer) As Integer

    'Declare Function DeleteFile Lib "kernel32"  Alias "DeleteFileA"(ByVal lpFileName As String) As Integer

    'Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Integer) As Integer

    'Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Integer, ByVal cb As Integer, ByRef cbNeeded As Integer) As Integer

    'Declare Sub ExitProcess Lib "kernel32" (ByVal vintExitCode As Short)

    Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal vlngClassName As Integer, ByVal vstrCaption As String) As Integer

    'Declare Function GetACP Lib "kernel32" () As Integer

    'Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Integer) As Short

    'Declare Function GetCaretBlinkTime Lib "user32" () As Integer

    'Declare Function GetComputerName Lib "kernel32"  Alias "GetComputerNameA"(ByVal lpBuffer As String, ByRef nSize As Integer) As Integer

    ''UPGRADE_WARNING: Structure CPINFO may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1050"'
    'Declare Function GetCPInfo Lib "kernel32" (ByVal CodePage As Integer, ByRef lpCPInfo As CPINFO) As Integer

    'Declare Function GetCurrentProcessId Lib "kernel32" () As Integer

    'Declare Function GetDoubleClickTime Lib "user32" () As Integer

    'Public Const STILL_ACTIVE As Short = 259 ' Return from GetExitCodeProcess if process still live

    'Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Integer, ByRef lpExitCode As Integer) As Integer

    'Declare Function GetKeyboardLayoutName Lib "user32"  Alias "GetKeyboardLayoutNameA"(ByVal pwszKLID As String) As Integer

    'Declare Function GetKeyboardState Lib "user32" (ByRef pbKeyState As Byte) As Integer

    'Declare Function GetKeyboardType Lib "user32" (ByVal nTypeFlag As Integer) As Integer

    'Declare Function GetKeyNameText Lib "user32"  Alias "GetKeyNameTextA"(ByVal lParam As Integer, ByVal lpBuffer As String, ByVal nSize As Integer) As Integer

    'Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Integer) As Short

    'Declare Function GetLocaleInfo Lib "kernel32"  Alias "GetLocaleInfoA"(ByVal Locale As Integer, ByVal LCType As Integer, ByVal lpLCData As String, ByVal cchData As Integer) As Integer

    'Declare Function GetMessageTime Lib "user32" () As Integer

    Public Declare Function GetWindow Lib "user32" Alias "GetWindow" _
        (ByVal hwnd As Integer, ByVal wFlag As Integer) As Integer

    ' GetWindow() Constants
    Public Const GW_HWNDFIRST As Integer = 0
    Public Const GW_HWNDLAST As Integer = 1
    Public Const GW_HWNDNEXT As Integer = 2
    Public Const GW_HWNDPREV As Integer = 3
    Public Const GW_OWNER As Integer = 4
    Public Const GW_CHILD As Integer = 5
    Public Const GW_MAX As Integer = 5

    'Declare Function GetOEMCP Lib "kernel32" () As Integer

    '' For reading INI files
    Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal sSection As String, ByVal sKey As String, ByVal sDefault As String, ByVal sReturn As String, ByVal nSize As Integer, ByVal sFilename As String) As Integer

    'Declare Function GetProcessHeap Lib "kernel32" () As Integer

    'Declare Function GetSysColor Lib "user32" (ByVal nIndex As Integer) As Integer

    ''UPGRADE_WARNING: Structure SYSTEM_INFO may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1050"'
    'Declare Sub GetSystemInfo Lib "kernel32" (ByRef lpSystemInfo As SYSTEM_INFO)

    'Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Integer) As Integer

    'Declare Function GetTempPath Lib "kernel32"  Alias "GetTempPathA"(ByVal nBufferLength As Integer, ByVal lpBuffer As String) As Integer

    'Declare Function GetTickCount Lib "kernel32" () As Integer

    Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Integer) As Integer

    'Declare Function GetUserName Lib "advapi32.dll"  Alias "GetUserNameA"(ByVal lpBuffer As String, ByRef nSize As Integer) As Integer

    'Declare Function GetVersion Lib "kernel32" () As Integer

    ''UPGRADE_WARNING: Structure OSVERSIONINFO may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1050"'
    'Declare Function GetVersionEx Lib "kernel32"  Alias "GetVersionExA"(ByRef lpVersionInformation As OSVERSIONINFO) As Integer

    Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Integer, ByRef lpdwProcessId As Integer) As Integer

    Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
        (ByVal hwnd As Integer, ByVal lpString As String, ByVal cch As Integer) As Integer
    'Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Integer, ByVal dwFlags As Integer, ByVal dwBytes As Integer) As Integer

    ''UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1016"'
    'Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Integer, ByVal dwFlags As Integer, ByRef lpMem As Any) As Integer

    'Declare Function MapVirtualKey Lib "user32"  Alias "MapVirtualKeyA"(ByVal wCode As Integer, ByVal wMapType As Integer) As Integer

    'Declare Function MessageBeep Lib "user32" (ByVal wType As Integer) As Integer

    ''UPGRADE_WARNING: Structure NCB may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1050"'
    'Declare Function Netbios Lib "netapi32.dll" (ByRef pncb As NCB) As Byte

    ''UPGRADE_WARNING: Structure PROCESS_BASIC_INFORMATION may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1050"'
    'Public Declare Function NtQueryInformationProcess Lib "ntdll" (ByVal ProcessHandle As Integer, ByVal ProcessInformationClass As Integer, ByRef ProcessInformation As PROCESS_BASIC_INFORMATION, ByVal lProcessInformationLength As Integer, ByRef lReturnLength As Integer) As Integer

    Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Integer, ByVal bInheritHandle As Integer, ByVal dwProcessId As Integer) As Integer

    Declare Function PostMessage Lib "user32" Alias "PostMessageA" _
        (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) _
        As Long

    'Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Integer) As Integer

    'Declare Function RegDeleteValue Lib "advapi32.dll"  Alias "RegDeleteValueA"(ByVal hKey As Integer, ByVal lpValue As String) As Integer

    'Declare Function RegCreateKeyEx Lib "advapi32.dll"  Alias "RegCreateKeyExA"(ByVal hKey As Integer, ByVal lpSubKey As String, ByVal Reserved As Integer, ByVal lpClass As String, ByVal dwOptions As Integer, ByVal samDesired As Integer, ByVal lpSecurityAttributes As Integer, ByRef phkResult As Integer, ByRef lpdwDisposition As Integer) As Integer

    'Declare Function RegOpenKeyEx Lib "advapi32.dll"  Alias "RegOpenKeyExA"(ByVal hKey As Integer, ByVal lpSubKey As String, ByVal ulOptions As Integer, ByVal samDesired As Integer, ByRef phkResult As Integer) As Integer

    'Declare Function RegQueryValueEx Lib "advapi32"  Alias "RegQueryValueExA"(ByVal hKey As Integer, ByVal lpValueName As String, ByVal lpReserved As Integer, ByRef lpType As Integer, ByVal lpData As String, ByRef lpcbData As Integer) As Integer

    'Declare Function RegQueryValueExString Lib "advapi32.dll"  Alias "RegQueryValueExA"(ByVal hKey As Integer, ByVal lpValueName As String, ByVal lpReserved As Integer, ByRef lpType As Integer, ByVal lpData As String, ByRef lpcbData As Integer) As Integer

    'Declare Function RegQueryValueExLong Lib "advapi32.dll"  Alias "RegQueryValueExA"(ByVal hKey As Integer, ByVal lpValueName As String, ByVal lpReserved As Integer, ByRef lpType As Integer, ByRef lpData As Integer, ByRef lpcbData As Integer) As Integer

    'Declare Function RegQueryValueExByte Lib "advapi32.dll"  Alias "RegQueryValueExA"(ByVal hKey As Integer, ByVal lpValueName As String, ByVal lpReserved As Integer, ByRef lpType As Integer, ByRef lpData As Byte, ByRef lpcbData As Integer) As Integer

    'Declare Function RegQueryValueExNULL Lib "advapi32.dll"  Alias "RegQueryValueExA"(ByVal hKey As Integer, ByVal lpValueName As String, ByVal lpReserved As Integer, ByRef lpType As Integer, ByVal lpData As Integer, ByRef lpcbData As Integer) As Integer

    'Declare Function RegSetValueExString Lib "advapi32.dll"  Alias "RegSetValueExA"(ByVal hKey As Integer, ByVal lpValueName As String, ByVal Reserved As Integer, ByVal dwType As Integer, ByVal lpValue As String, ByVal cbData As Integer) As Integer

    'Declare Function RegSetValueExLong Lib "advapi32.dll"  Alias "RegSetValueExA"(ByVal hKey As Integer, ByVal lpValueName As String, ByVal Reserved As Integer, ByVal dwType As Integer, ByRef lpValue As Integer, ByVal cbData As Integer) As Integer

    Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) _
        As Long

    ' Public Const WM_COMMAND As Integer = &H111
    Public Const MA_ACTIVATE As Integer = 1
    Public Const WM_ACTIVATE As Integer = &H6
    Public Const WM_LBUTTONDOWN As Integer = &H201
    Public Const WM_LBUTTONUP As Integer = &H202
    Public Const WM_LBUTTONDBLCLK As Integer = &H203
    Public Const MK_LBUTTON As Integer = &H1

    'Public Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Integer) As Integer

    'Declare Function SetCaretBlinkTime Lib "user32" (ByVal wMSeconds As Integer) As Integer

    'Declare Function SetDoubleClickTime Lib "user32" (ByVal wCount As Integer) As Integer

    'Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Integer) As Boolean

    'Declare Function SetKeyboardState Lib "user32" (ByRef lppbKeyState As Byte) As Integer

    'Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Integer, ByVal hWndInsertAfter As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer) As Integer

    Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Integer, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer

    ''UPGRADE_WARNING: Structure NOTIFYICONDATA may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1050"'
    'Public Declare Function Shell_NotifyIcon Lib "shell32"  Alias "Shell_NotifyIconA"(ByVal dwMessage As Integer, ByRef pnid As NOTIFYICONDATA) As Boolean

    'Declare Function ShowCaret Lib "user32" (ByVal hwnd As Integer) As Integer

    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Integer)

    'Declare Function sndPlaySound Lib "winmm.dll"  Alias "sndPlaySoundA"(ByVal lpszSoundName As String, ByVal uFlags As Integer) As Integer

    'Declare Function SQLAllocEnv Lib "ODBC32.DLL" (ByRef env As Integer) As Short

    'Declare Function SQLDataSources Lib "ODBC32.DLL" (ByVal henv As Integer, ByVal fDirection As Short, ByVal szDSN As String, ByVal cbDSNMax As Short, ByRef pcbDSN As Short, ByVal szDescription As String, ByVal cbDescriptionMax As Short, ByRef pcbDescription As Short) As Short

    ''UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1016"'
    'Declare Function SystemParametersInfo Lib "user32"  Alias "SystemParametersInfoA"(ByVal uAction As Integer, ByVal uParam As Integer, ByRef lpvParam As Any, ByVal fuWinIni As Integer) As Integer

    Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Integer, ByVal uExitCode As Integer) As Integer

    '' For millisecond timing in trace system
    'Declare Function timeGetTime Lib "winmm.dll" () As Integer

    ' For writing .INI files
    Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal sSection As String, ByVal sKey As String, ByVal sValue As String, ByVal sFilename As String) As Boolean


    '' Private Const STILL_ACTIVE = &H103

    '' Consts for OpenProcess
    'Public Const PROCESS_QUERY_INFORMATION As Short = &H400s
    'Public Const PROCESS_VM_READ As Short = 16
    Public Const PROCESS_TERMINATE As Short = 1

    'Public Structure PROCESS_BASIC_INFORMATION
    '	Dim ExitStatus As Integer
    '	Dim PebBaseAddress As Integer
    '	Dim AffinityMask As Integer
    '	Dim BasePriority As Integer
    '	Dim UniqueProcessId As Integer
    '	Dim InheritedFromUniqueProcessId As Integer 'ParentProcessID'
    'End Structure

    ''-------------------------------------------------------
    ''           API constants
    ''-------------------------------------------------------

    'Public Const GFSR_SYSTEMRESOURCES As Short = 0
    'Public Const GFSR_GDIRESOURCES As Short = 1
    'Public Const GFSR_USERRESOURCES As Short = 2
    'Public Const WF_PMODE As Short = &H1s
    'Public Const WF_CPU286 As Short = &H2s
    'Public Const WF_CPU386 As Short = &H4s
    'Public Const WF_CPU486 As Short = &H8s
    'Public Const WF_STANDARD As Short = &H10s
    'Public Const WF_WIN286 As Short = &H10s
    'Public Const WF_ENHANCED As Short = &H20s
    'Public Const WF_WIN386 As Short = &H20s
    'Public Const WF_CPU086 As Short = &H40s
    'Public Const WF_CPU186 As Short = &H80s
    'Public Const WF_LARGEFRAME As Short = &H100s
    'Public Const WF_SMALLFRAME As Short = &H200s
    'Public Const WF_80x87 As Short = &H400s
    'Public Const VK_LBUTTON As Short = &H1s
    'Public Const VK_RBUTTON As Short = &H2s
    'Public Const VK_CANCEL As Short = &H3s
    'Public Const VK_MBUTTON As Short = &H4s
    'Public Const VK_BACK As Short = &H8s
    'Public Const VK_TAB As Short = &H9s
    'Public Const VK_CLEAR As Short = &HCs
    'Public Const VK_RETURN As Short = &HDs
    'Public Const VK_SHIFT As Short = &H10s
    'Public Const VK_CONTROL As Short = &H11s
    'Public Const VK_MENU As Short = &H12s
    'Public Const VK_PAUSE As Short = &H13s
    'Public Const VK_CAPITAL As Short = &H14s
    'Public Const VK_ESCAPE As Short = &H1Bs
    'Public Const VK_SPACE As Short = &H20s
    'Public Const VK_PRIOR As Short = &H21s
    'Public Const VK_NEXT As Short = &H22s
    'Public Const VK_END As Short = &H23s
    'Public Const VK_HOME As Short = &H24s
    'Public Const VK_LEFT As Short = &H25s
    'Public Const VK_UP As Short = &H26s
    'Public Const VK_RIGHT As Short = &H27s
    'Public Const VK_DOWN As Short = &H28s
    'Public Const VK_SELECT As Short = &H29s
    'Public Const VK_PRINT As Short = &H2As
    'Public Const VK_EXECUTE As Short = &H2Bs
    'Public Const VK_SNAPSHOT As Short = &H2Cs
    'Public Const VK_INSERT As Short = &H2Ds
    'Public Const VK_DELETE As Short = &H2Es
    'Public Const VK_HELP As Short = &H2Fs
    'Public Const VK_NUMPAD0 As Short = &H60s
    'Public Const VK_NUMPAD1 As Short = &H61s
    'Public Const VK_NUMPAD2 As Short = &H62s
    'Public Const VK_NUMPAD3 As Short = &H63s
    'Public Const VK_NUMPAD4 As Short = &H64s
    'Public Const VK_NUMPAD5 As Short = &H65s
    'Public Const VK_NUMPAD6 As Short = &H66s
    'Public Const VK_NUMPAD7 As Short = &H67s
    'Public Const VK_NUMPAD8 As Short = &H68s
    'Public Const VK_NUMPAD9 As Short = &H69s
    'Public Const VK_MULTIPLY As Short = &H6As
    'Public Const VK_ADD As Short = &H6Bs
    'Public Const VK_SEPARATOR As Short = &H6Cs
    'Public Const VK_SUBTRACT As Short = &H6Ds
    'Public Const VK_DECIMAL As Short = &H6Es
    'Public Const VK_DIVIDE As Short = &H6Fs
    'Public Const VK_F1 As Short = &H70s
    'Public Const VK_F2 As Short = &H71s
    'Public Const VK_F3 As Short = &H72s
    'Public Const VK_F4 As Short = &H73s
    'Public Const VK_F5 As Short = &H74s
    'Public Const VK_F6 As Short = &H75s
    'Public Const VK_F7 As Short = &H76s
    'Public Const VK_F8 As Short = &H77s
    'Public Const VK_F9 As Short = &H78s
    'Public Const VK_F10 As Short = &H79s
    'Public Const VK_F11 As Short = &H7As
    'Public Const VK_F12 As Short = &H7Bs
    'Public Const VK_F13 As Short = &H7Cs
    'Public Const VK_F14 As Short = &H7Ds
    'Public Const VK_F15 As Short = &H7Es
    'Public Const VK_F16 As Short = &H7Fs
    'Public Const VK_NUMLOCK As Short = &H90s
    'Public Const VK_SCROLL As Short = &H91s
    'Public Const WM_USER As Short = &H400s
    'Public Const COLOR_SCROLLBAR As Short = 0
    'Public Const COLOR_BACKGROUND As Short = 1
    'Public Const COLOR_ACTIVECAPTION As Short = 2
    'Public Const COLOR_INACTIVECAPTION As Short = 3
    'Public Const COLOR_MENU As Short = 4
    'Public Const COLOR_WINDOW As Short = 5
    'Public Const COLOR_WINDOWFRAME As Short = 6
    'Public Const COLOR_MENUTEXT As Short = 7
    'Public Const COLOR_WINDOWTEXT As Short = 8
    'Public Const COLOR_CAPTIONTEXT As Short = 9
    'Public Const COLOR_ACTIVEBORDER As Short = 10
    'Public Const COLOR_INACTIVEBORDER As Short = 11
    'Public Const COLOR_APPWORKSPACE As Short = 12
    'Public Const COLOR_HIGHLIGHT As Short = 13
    'Public Const COLOR_HIGHLIGHTTEXT As Short = 14
    'Public Const COLOR_BTNFACE As Short = 15
    'Public Const COLOR_BTNSHADOW As Short = 16
    'Public Const COLOR_GRAYTEXT As Short = 17
    'Public Const COLOR_BTNTEXT As Short = 18
    'Public Const COLOR_INACTIVECAPTIONTEXT As Short = 19
    'Public Const COLOR_BTNHIGHLIGHT As Short = 20
    'Public Const COLOR_3DDKSHADOW As Short = 21
    'Public Const COLOR_3DLIGHT As Short = 22
    'Public Const COLOR_INFOBK As Short = 24
    'Public Const COLOR_INFOTEXT As Short = 23
    'Public Const RGB_BLACK As Short = 0 'RGB(0, 0, 0)
    'Public Const RGB_BLUE As Integer = 16711680 'RGB(0, 0, 255)
    'Public Const RGB_GREEN As Integer = 65280 'RGB(0, 255, 0)
    'Public Const RGB_CYAN As Integer = 16776960 'RGB(0, 255, 255)
    'Public Const RGB_RED As Short = 255 'RGB(255, 0, 0)
    'Public Const RGB_MAGENTA As Integer = 16711935 'RGB(255, 0, 255)
    'Public Const RGB_YELLOW As Integer = 65535 'RGB(255, 255, 0)
    'Public Const RGB_WHITE As Integer = 16777215 'RGB(255, 255, 255)
    'Public Const RGB_DARKGREY As Integer = &H80000010
    'Public Const RGB_DARKGREEN As Integer = &H8000
    'Public Const RGB_DARKRED As Integer = &HC0
    'Public Const RGB_DARKYELLOW As Integer = &HC0C0

    'Public Const SM_CXSCREEN As Short = 0
    'Public Const SM_CYSCREEN As Short = 1
    'Public Const SM_CXVSCROLL As Short = 2
    'Public Const SM_CYHSCROLL As Short = 3
    'Public Const SM_CYCAPTION As Short = 4
    'Public Const SM_CXBORDER As Short = 5
    'Public Const SM_CYBORDER As Short = 6
    'Public Const SM_CXDLGFRAME As Short = 7
    'Public Const SM_CYDLGFRAME As Short = 8
    'Public Const SM_CYVTHUMB As Short = 9
    'Public Const SM_CXHTHUMB As Short = 10
    'Public Const SM_CXICON As Short = 11
    'Public Const SM_CYICON As Short = 12
    'Public Const SM_CXCURSOR As Short = 13
    'Public Const SM_CYCURSOR As Short = 14
    'Public Const SM_CYMENU As Short = 15
    'Public Const SM_CXFULLSCREEN As Short = 16
    'Public Const SM_CYFULLSCREEN As Short = 17
    'Public Const SM_CYKANJIWINDOW As Short = 18
    'Public Const SM_MOUSEPRESENT As Short = 19
    'Public Const SM_CYVSCROLL As Short = 20
    'Public Const SM_CXHSCROLL As Short = 21
    'Public Const SM_DEBUG As Short = 22
    'Public Const SM_SWAPBUTTON As Short = 23
    'Public Const SM_RESERVED1 As Short = 24
    'Public Const SM_RESERVED2 As Short = 25
    'Public Const SM_RESERVED3 As Short = 26
    'Public Const SM_RESERVED4 As Short = 27
    'Public Const SM_CXMIN As Short = 28
    'Public Const SM_CYMIN As Short = 29
    'Public Const SM_CXSIZE As Short = 30
    'Public Const SM_CYSIZE As Short = 31
    'Public Const SM_CXFRAME As Short = 32
    'Public Const SM_CYFRAME As Short = 33
    'Public Const SM_CXMINTRACK As Short = 34
    'Public Const SM_CYMINTRACK As Short = 35
    'Public Const SM_CXDOUBLECLK As Short = 36
    'Public Const SM_CYDOUBLECLK As Short = 37
    'Public Const SM_CXICONSPACING As Short = 38
    'Public Const SM_CYICONSPACING As Short = 39
    'Public Const SM_MENUDROPALIGNMENT As Short = 40
    'Public Const SM_PENWINDOWS As Short = 41
    'Public Const SM_DBCSENABLED As Short = 42
    'Public Const SM_CMOUSEBUTTONS As Short = 43
    'Public Const SM_CXFIXEDFRAME As Short = SM_CXDLGFRAME
    'Public Const SM_CYFIXEDFRAME As Short = SM_CYDLGFRAME
    'Public Const SM_CXSIZEFRAME As Short = SM_CXFRAME
    'Public Const SM_CYSIZEFRAME As Short = SM_CYFRAME
    'Public Const SM_SECURE As Short = 44
    'Public Const SM_CXEDGE As Short = 45
    'Public Const SM_CYEDGE As Short = 46
    'Public Const SM_CXMINSPACING As Short = 47
    'Public Const SM_CYMINSPACING As Short = 48
    'Public Const SM_CXSMICON As Short = 49
    'Public Const SM_CYSMICON As Short = 50
    'Public Const SM_CYSMCAPTION As Short = 51
    'Public Const SM_CXSMSIZE As Short = 52
    'Public Const SM_CYSMSIZE As Short = 53
    'Public Const SM_CXMENUSIZE As Short = 54
    'Public Const SM_CYMENUSIZE As Short = 55
    'Public Const SM_ARRANGE As Short = 56
    'Public Const SM_CXMINIMIZED As Short = 57
    'Public Const SM_CYMINIMIZED As Short = 58
    'Public Const SM_CXMAXTRACK As Short = 59
    'Public Const SM_CYMAXTRACK As Short = 60
    'Public Const SM_CXMAXIMIZED As Short = 61
    'Public Const SM_CYMAXIMIZED As Short = 62
    'Public Const SM_NETWORK As Short = 63
    'Public Const SM_CLEANBOOT As Short = 67
    'Public Const SM_CXDRAG As Short = 68
    'Public Const SM_CYDRAG As Short = 69
    'Public Const SM_SHOWSOUNDS As Short = 70
    'Public Const SM_CXMENUCHECK As Short = 71
    'Public Const SM_CYMENUCHECK As Short = 72
    'Public Const SM_SLOWMACHINE As Short = 73
    'Public Const SM_MIDEASTENABLED As Short = 74
    'Public Const SM_CMETRICS As Short = 75


    'Public Const VER_PLATFORM_WIN32_NT As Integer = 2
    'Public Const VER_PLATFORM_WIN32_WINDOWS As Integer = 1

    'Public Const SPI_GETACCESSTIMEOUT As Integer = 60
    'Public Const SPI_GETANIMATION As Integer = 72
    'Public Const SPI_GETBEEP As Integer = 1
    'Public Const SPI_GETBORDER As Integer = 5
    'Public Const SPI_GETDEFAULTINPUTLANG As Integer = 89
    'Public Const SPI_GETDRAGFULLWINDOWS As Integer = 38
    'Public Const SPI_GETFASTTASKSWITCH As Integer = 35
    'Public Const SPI_GETFILTERKEYS As Integer = 50
    'Public Const SPI_GETFONTSMOOTHING As Integer = 74
    'Public Const SPI_GETGRIDGRANULARITY As Integer = 18
    'Public Const SPI_GETHIGHCONTRAST As Integer = 66
    'Public Const SPI_GETICONMETRICS As Integer = 45
    'Public Const SPI_GETICONTITLELOGFONT As Integer = 31
    'Public Const SPI_GETICONTITLEWRAP As Integer = 25
    'Public Const SPI_GETKEYBOARDDELAY As Integer = 22
    'Public Const SPI_GETKEYBOARDPREF As Integer = 68
    'Public Const SPI_GETKEYBOARDSPEED As Integer = 10
    'Public Const SPI_GETLOWPOWERACTIVE As Integer = 83
    'Public Const SPI_GETLOWPOWERTIMEOUT As Integer = 79
    'Public Const SPI_GETMENUDROPALIGNMENT As Integer = 27
    'Public Const SPI_GETMOUSE As Integer = 3
    'Public Const SPI_GETMINIMIZEDMETRICS As Integer = 43
    'Public Const SPI_GETMOUSEKEYS As Integer = 54
    'Public Const SPI_GETMOUSETRAILS As Integer = 94
    'Public Const SPI_GETNONCLIENTMETRICS As Integer = 41
    'Public Const SPI_GETPOWEROFFACTIVE As Integer = 84
    'Public Const SPI_GETPOWEROFFTIMEOUT As Integer = 80
    'Public Const SPI_GETSCREENREADER As Integer = 70
    'Public Const SPI_GETSCREENSAVEACTIVE As Integer = 16
    'Public Const SPI_GETSCREENSAVETIMEOUT As Integer = 14
    'Public Const SPI_GETSERIALKEYS As Integer = 62
    'Public Const SPI_GETSHOWSOUNDS As Integer = 56
    'Public Const SPI_GETSOUNDSENTRY As Integer = 64
    'Public Const SPI_GETSTICKYKEYS As Integer = 58
    'Public Const SPI_GETTOGGLEKEYS As Integer = 52
    'Public Const SPI_GETWINDOWSEXTENSION As Integer = 92
    'Public Const SPI_GETWORKAREA As Integer = 48
    'Public Const SPI_ICONHORIZONTALSPACING As Integer = 13
    'Public Const SPI_ICONVERTICALSPACING As Integer = 24
    'Public Const SPI_LANGDRIVER As Integer = 12
    'Public Const SPI_SCREENSAVERRUNNING As Integer = 97
    'Public Const SPI_SETACCESSTIMEOUT As Integer = 61
    'Public Const SPI_SETANIMATION As Integer = 73
    'Public Const SPI_SETBEEP As Integer = 2
    'Public Const SPI_SETBORDER As Integer = 6
    'Public Const SPI_SETCURSORS As Integer = 87
    'Public Const SPI_SETDEFAULTINPUTLANG As Integer = 90
    'Public Const SPI_SETDESKPATTERN As Integer = 21
    'Public Const SPI_SETDESKWALLPAPER As Integer = 20
    'Public Const SPI_SETDOUBLECLICKTIME As Integer = 32
    'Public Const SPI_SETDOUBLECLKHEIGHT As Integer = 30
    'Public Const SPI_SETDOUBLECLKWIDTH As Integer = 29
    'Public Const SPI_SETDRAGFULLWINDOWS As Integer = 37
    'Public Const SPI_SETDRAGHEIGHT As Integer = 77
    'Public Const SPI_SETDRAGWIDTH As Integer = 76
    'Public Const SPI_SETFASTTASKSWITCH As Integer = 36
    'Public Const SPI_SETFILTERKEYS As Integer = 51
    'Public Const SPI_SETFONTSMOOTHING As Integer = 75
    'Public Const SPI_SETGRIDGRANULARITY As Integer = 19
    'Public Const SPI_SETHANDHELD As Integer = 78
    'Public Const SPI_SETHIGHCONTRAST As Integer = 67
    'Public Const SPI_SETICONMETRICS As Integer = 46
    'Public Const SPI_SETICONS As Integer = 88
    'Public Const SPI_SETICONTITLELOGFONT As Integer = 34
    'Public Const SPI_SETICONTITLEWRAP As Integer = 26
    'Public Const SPI_SETKEYBOARDDELAY As Integer = 23
    'Public Const SPI_SETKEYBOARDPREF As Integer = 69
    'Public Const SPI_SETKEYBOARDSPEED As Integer = 11
    'Public Const SPI_SETLANGTOGGLE As Integer = 91
    'Public Const SPI_SETLOWPOWERACTIVE As Integer = 85
    'Public Const SPI_SETLOWPOWERTIMEOUT As Integer = 81
    'Public Const SPI_SETMENUDROPALIGNMENT As Integer = 28
    'Public Const SPI_SETMINIMIZEDMETRICS As Integer = 44
    'Public Const SPI_SETMOUSE As Integer = 4
    'Public Const SPI_SETMOUSEBUTTONSWAP As Integer = 33
    'Public Const SPI_SETMOUSEKEYS As Integer = 55
    'Public Const SPI_SETMOUSETRAILS As Integer = 93
    'Public Const SPI_SETNONCLIENTMETRICS As Integer = 42
    'Public Const SPI_SETPENWINDOWS As Integer = 49
    'Public Const SPI_SETPOWEROFFACTIVE As Integer = 86
    'Public Const SPI_SETPOWEROFFTIMEOUT As Integer = 82
    'Public Const SPI_SETSCREENREADER As Integer = 71
    'Public Const SPI_SETSCREENSAVEACTIVE As Integer = 17
    'Public Const SPI_SETSCREENSAVETIMEOUT As Integer = 15
    'Public Const SPI_SETSERIALKEYS As Integer = 63
    'Public Const SPI_SETSHOWSOUNDS As Integer = 57
    'Public Const SPI_SETSOUNDSENTRY As Integer = 65
    'Public Const SPI_SETSTICKYKEYS As Integer = 59
    'Public Const SPI_SETTOGGLEKEYS As Integer = 53
    'Public Const SPI_SETWORKAREA As Integer = 47
    'Public Const SPIF_UPDATEINIFILE As Short = 1
    'Public Const SPIF_SENDWININICHANGE As Short = 2

    '' GetWindow constants
    'Public Const GW_HWNDNEXT As Short = 2
    'Public Const GW_HWNDPREV As Short = 3

    ' ShowWindow constants
    Public Const SW_HIDE As Short = 0
    Public Const SW_SHOWNORMAL As Short = 1
    Public Const SW_NORMAL As Short = 1
    Public Const SW_SHOWMINIMIZED As Short = 2
    Public Const SW_SHOWMAXIMIZED As Short = 3
    Public Const SW_MAXIMIZE As Short = 3
    Public Const SW_SHOWNOACTIVATE As Short = 4
    Public Const SW_SHOW As Short = 5
    Public Const SW_MINIMIZE As Short = 6
    Public Const SW_SHOWMINNOACTIVE As Short = 7
    Public Const SW_SHOWNA As Short = 8
    Public Const SW_RESTORE As Short = 9
    Public Const SW_SHOWDEFAULT As Short = 10
    Public Const SW_FORCEMINIMIZE As Short = 11

    'Public Const EWX_LOGOFF As Short = 0
    'Public Const EWX_SHUTDOWN As Short = 1
    'Public Const EWX_REBOOT As Short = 2
    'Public Const EWX_FORCE As Short = 4
    'Public Const EWX_POWEROFF As Short = 8

    'Public Const PROCESSOR_INTEL_386 As Short = 386
    'Public Const PROCESSOR_INTEL_486 As Short = 486
    'Public Const PROCESSOR_INTEL_PENTIUM As Short = 586
    'Public Const PROCESSOR_MIPS_R4000 As Short = 4000
    'Public Const PROCESSOR_ALPHA_21064 As Short = 21064

    'Public Const LOCALE_SYSTEM_DEFAULT As Short = &H800s
    'Public Const LOCALE_USER_DEFAULT As Short = &H400s

    'Public Const LOCALE_ILANGUAGE As Short = &H1s '  language id
    'Public Const LOCALE_SLANGUAGE As Short = &H2s '  localized name of language
    'Public Const LOCALE_SENGLANGUAGE As Short = &H1001s '  English name of language
    'Public Const LOCALE_SABBREVLANGNAME As Short = &H3s '  abbreviated language name
    'Public Const LOCALE_SNATIVELANGNAME As Short = &H4s '  native name of language
    'Public Const LOCALE_ICOUNTRY As Short = &H5s '  country code
    'Public Const LOCALE_SCOUNTRY As Short = &H6s '  localized name of country
    'Public Const LOCALE_SENGCOUNTRY As Short = &H1002s '  English name of country
    'Public Const LOCALE_SABBREVCTRYNAME As Short = &H7s '  abbreviated country name
    'Public Const LOCALE_SNATIVECTRYNAME As Short = &H8s '  native name of country
    'Public Const LOCALE_IDEFAULTLANGUAGE As Short = &H9s '  default language id
    'Public Const LOCALE_IDEFAULTCOUNTRY As Short = &HAs '  default country code
    'Public Const LOCALE_IDEFAULTCODEPAGE As Short = &HBs '  default code page

    'Public Const LOCALE_SLIST As Short = &HCs '  list item separator
    'Public Const LOCALE_IMEASURE As Short = &HDs '  0 = metric, 1 = US

    'Public Const LOCALE_SDECIMAL As Short = &HEs '  decimal separator
    'Public Const LOCALE_STHOUSAND As Short = &HFs '  thousand separator
    'Public Const LOCALE_SGROUPING As Short = &H10s '  digit grouping
    'Public Const LOCALE_IDIGITS As Short = &H11s '  number of fractional digits
    'Public Const LOCALE_ILZERO As Short = &H12s '  leading zeros for decimal
    'Public Const LOCALE_SNATIVEDIGITS As Short = &H13s '  native ascii 0-9

    'Public Const LOCALE_SCURRENCY As Short = &H14s '  local monetary symbol
    'Public Const LOCALE_SINTLSYMBOL As Short = &H15s '  intl monetary symbol
    'Public Const LOCALE_SMONDECIMALSEP As Short = &H16s '  monetary decimal separator
    'Public Const LOCALE_SMONTHOUSANDSEP As Short = &H17s '  monetary thousand separator
    'Public Const LOCALE_SMONGROUPING As Short = &H18s '  monetary grouping
    'Public Const LOCALE_ICURRDIGITS As Short = &H19s '  # local monetary digits
    'Public Const LOCALE_IINTLCURRDIGITS As Short = &H1As '  # intl monetary digits
    'Public Const LOCALE_ICURRENCY As Short = &H1Bs '  positive currency mode
    'Public Const LOCALE_INEGCURR As Short = &H1Cs '  negative currency mode

    'Public Const LOCALE_SDATE As Short = &H1Ds '  date separator
    'Public Const LOCALE_STIME As Short = &H1Es '  time separator
    'Public Const LOCALE_SSHORTDATE As Short = &H1Fs '  short date format string
    'Public Const LOCALE_SLONGDATE As Short = &H20s '  long date format string
    'Public Const LOCALE_STIMEFORMAT As Short = &H1003s '  time format string
    'Public Const LOCALE_IDATE As Short = &H21s '  short date format ordering
    'Public Const LOCALE_ILDATE As Short = &H22s '  long date format ordering
    'Public Const LOCALE_ITIME As Short = &H23s '  time format specifier
    'Public Const LOCALE_ICENTURY As Short = &H24s '  century format specifier
    'Public Const LOCALE_ITLZERO As Short = &H25s '  leading zeros in time field
    'Public Const LOCALE_IDAYLZERO As Short = &H26s '  leading zeros in day field
    'Public Const LOCALE_IMONLZERO As Short = &H27s '  leading zeros in month field
    'Public Const LOCALE_S1159 As Short = &H28s '  AM designator
    'Public Const LOCALE_S2359 As Short = &H29s '  PM designator

    'Public Const LOCALE_SDAYNAME1 As Short = &H2As '  long name for Monday
    'Public Const LOCALE_SDAYNAME2 As Short = &H2Bs '  long name for Tuesday
    'Public Const LOCALE_SDAYNAME3 As Short = &H2Cs '  long name for Wednesday
    'Public Const LOCALE_SDAYNAME4 As Short = &H2Ds '  long name for Thursday
    'Public Const LOCALE_SDAYNAME5 As Short = &H2Es '  long name for Friday
    'Public Const LOCALE_SDAYNAME6 As Short = &H2Fs '  long name for Saturday
    'Public Const LOCALE_SDAYNAME7 As Short = &H30s '  long name for Sunday
    'Public Const LOCALE_SABBREVDAYNAME1 As Short = &H31s '  abbreviated name for Monday
    'Public Const LOCALE_SABBREVDAYNAME2 As Short = &H32s '  abbreviated name for Tuesday
    'Public Const LOCALE_SABBREVDAYNAME3 As Short = &H33s '  abbreviated name for Wednesday
    'Public Const LOCALE_SABBREVDAYNAME4 As Short = &H34s '  abbreviated name for Thursday
    'Public Const LOCALE_SABBREVDAYNAME5 As Short = &H35s '  abbreviated name for Friday
    'Public Const LOCALE_SABBREVDAYNAME6 As Short = &H36s '  abbreviated name for Saturday
    'Public Const LOCALE_SABBREVDAYNAME7 As Short = &H37s '  abbreviated name for Sunday
    'Public Const LOCALE_SMONTHNAME1 As Short = &H38s '  long name for January
    'Public Const LOCALE_SMONTHNAME2 As Short = &H39s '  long name for February
    'Public Const LOCALE_SMONTHNAME3 As Short = &H3As '  long name for March
    'Public Const LOCALE_SMONTHNAME4 As Short = &H3Bs '  long name for April
    'Public Const LOCALE_SMONTHNAME5 As Short = &H3Cs '  long name for May
    'Public Const LOCALE_SMONTHNAME6 As Short = &H3Ds '  long name for June
    'Public Const LOCALE_SMONTHNAME7 As Short = &H3Es '  long name for July
    'Public Const LOCALE_SMONTHNAME8 As Short = &H3Fs '  long name for August
    'Public Const LOCALE_SMONTHNAME9 As Short = &H40s '  long name for September
    'Public Const LOCALE_SMONTHNAME10 As Short = &H41s '  long name for October
    'Public Const LOCALE_SMONTHNAME11 As Short = &H42s '  long name for November
    'Public Const LOCALE_SMONTHNAME12 As Short = &H43s '  long name for December
    'Public Const LOCALE_SABBREVMONTHNAME1 As Short = &H44s '  abbreviated name for January
    'Public Const LOCALE_SABBREVMONTHNAME2 As Short = &H45s '  abbreviated name for February
    'Public Const LOCALE_SABBREVMONTHNAME3 As Short = &H46s '  abbreviated name for March
    'Public Const LOCALE_SABBREVMONTHNAME4 As Short = &H47s '  abbreviated name for April
    'Public Const LOCALE_SABBREVMONTHNAME5 As Short = &H48s '  abbreviated name for May
    'Public Const LOCALE_SABBREVMONTHNAME6 As Short = &H49s '  abbreviated name for June
    'Public Const LOCALE_SABBREVMONTHNAME7 As Short = &H4As '  abbreviated name for July
    'Public Const LOCALE_SABBREVMONTHNAME8 As Short = &H4Bs '  abbreviated name for August
    'Public Const LOCALE_SABBREVMONTHNAME9 As Short = &H4Cs '  abbreviated name for September
    'Public Const LOCALE_SABBREVMONTHNAME10 As Short = &H4Ds '  abbreviated name for October
    'Public Const LOCALE_SABBREVMONTHNAME11 As Short = &H4Es '  abbreviated name for November
    'Public Const LOCALE_SABBREVMONTHNAME12 As Short = &H4Fs '  abbreviated name for December
    'Public Const LOCALE_SABBREVMONTHNAME13 As Short = &H100Fs

    'Public Const LOCALE_SPOSITIVESIGN As Short = &H50s '  positive sign
    'Public Const LOCALE_SNEGATIVESIGN As Short = &H51s '  negative sign
    'Public Const LOCALE_IPOSSIGNPOSN As Short = &H52s '  positive sign position
    'Public Const LOCALE_INEGSIGNPOSN As Short = &H53s '  negative sign position
    'Public Const LOCALE_IPOSSYMPRECEDES As Short = &H54s '  mon sym precedes pos amt
    'Public Const LOCALE_IPOSSEPBYSPACE As Short = &H55s '  mon sym sep by space from pos amt
    'Public Const LOCALE_INEGSYMPRECEDES As Short = &H56s '  mon sym precedes neg amt
    'Public Const LOCALE_INEGSEPBYSPACE As Short = &H57s '  mon sym sep by space from neg amt

    'Public Const HWND_TOPMOST As Short = -1
    'Public Const SWP_NOMOVE As Short = &H2s
    'Public Const SWP_NOSIZE As Short = &H1s
    'Public Const MAX_COMPUTERNAME_LENGTH As Short = 15


    'Public Const SM_CMONITORS As Short = 80 ' to get number of monitors

    ''Constants for the return value when finding a monitor
    'Public Const MONITOR_DEFAULTTONULL As Short = &H0s 'If the monitor is not found, return 0
    'Public Const MONITOR_DEFAULTTOPRIMARY As Short = &H1s 'If the monitor is not found, return the primary monitor
    'Public Const MONITOR_DEFAULTTONEAREST As Short = &H2s 'If the monitor is not found, return the nearest monitor

    ''---------------------------------------------------

    ''              Application Global Variables

    ''---------------------------------------------------

    '' Holder for the original caret blink time
    'Public OriginalCaretBlinkTime As Short

    '' Holder for version information. Set on form load
    'Public myVer As OSVERSIONINFO


    ''------------------------------------------------------------------------------------
    ''---------------------  APIs that allows to manipulate the registry -----------------
    ''------------------------------------------------------------------------------------
    'Public Enum En_RegKeyType
    '	REG_SZ = 1
    '	REG_BINARY = 3 ' Free form binary
    '	REG_DWORD = 4
    '	REG_MULTI_SZ = 7
    'End Enum

    'Public Enum En_RegPredefinedKey
    '	HKEY_CLASSES_ROOT = &H80000000
    '	HKEY_CURRENT_USER = &H80000001
    '	HKEY_LOCAL_MACHINE = &H80000002
    '	HKEY_USERS = &H80000003
    'End Enum

    'Public Const ERROR_NONE As Short = 0
    'Public Const ERROR_BADDB As Short = 1
    'Public Const ERROR_BADKEY As Short = 2
    'Public Const ERROR_CANTOPEN As Short = 3
    'Public Const ERROR_CANTREAD As Short = 4
    'Public Const ERROR_CANTWRITE As Short = 5
    'Public Const ERROR_OUTOFMEMORY As Short = 6
    'Public Const ERROR_ARENA_TRASHED As Short = 7
    'Public Const ERROR_ACCESS_DENIED As Short = 8
    'Public Const ERROR_INVALID_PARAMETERS As Short = 87
    'Public Const ERROR_NO_MORE_ITEMS As Short = 259

    ''reg security
    'Public Const KEY_ALL_ACCESS As Short = &H3Fs
    'Public Const KEY_QUERY_VALUE As Short = &H1s

    'Public Const REG_OPTION_NON_VOLATILE As Short = 0




    ''--------------------------------------------------------------------------
    ''---------------------  APIs for MAC(Network Card Number) -----------------
    ''--------------------------------------------------------------------------
    'Public Const MB_ICONASTERISK As Integer = &H40
    'Public Const MB_ICONEXCLAMATION As Integer = &H30
    'Public Const MB_ICONHAND As Integer = &H10
    'Public Const MB_ICONQUESTION As Integer = &H20


    ''--------------------------------------------------------------------------
    ''---------------------  APIs for MAC(Network Card Number) -----------------
    ''--------------------------------------------------------------------------
    'Public Const NCBASTAT As Short = &H33s
    'Public Const NCBNAMSZ As Short = 16
    'Public Const HEAP_ZERO_MEMORY As Short = &H8s
    'Public Const HEAP_GENERATE_EXCEPTIONS As Short = &H4s
    'Public Const NCBRESET As Short = &H32s

    'Public Structure NCB
    '	Dim ncb_command As Byte 'Integer
    '	Dim ncb_retcode As Byte 'Integer
    '	Dim ncb_lsn As Byte 'Integer
    '	Dim ncb_num As Byte ' Integer
    '	Dim ncb_buffer As Integer 'String
    '	Dim ncb_length As Short
    '	<VBFixedString(NCBNAMSZ),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr,SizeConst:=NCBNAMSZ)> Public ncb_callname As String
    '	<VBFixedString(NCBNAMSZ),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr,SizeConst:=NCBNAMSZ)> Public ncb_name As String
    '	Dim ncb_rto As Byte 'Integer
    '	Dim ncb_sto As Byte ' Integer
    '	Dim ncb_post As Integer
    '	Dim ncb_lana_num As Byte 'Integer
    '	Dim ncb_cmd_cplt As Byte 'Integer
    '	<VBFixedArray(9)> Dim ncb_reserve() As Byte ' Reserved, must be 0
    '	Dim ncb_event As Integer

    '	'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1026"'
    '	Public Sub Initialize()
    '		ReDim ncb_reserve(9)
    '	End Sub
    'End Structure
    'Public Structure ADAPTER_STATUS
    '	<VBFixedArray(5)> Dim adapter_address() As Byte 'As String * 6
    '	Dim rev_major As Byte 'Integer
    '	Dim reserved0 As Byte 'Integer
    '	Dim adapter_type As Byte 'Integer
    '	Dim rev_minor As Byte 'Integer
    '	Dim Duration As Short
    '	Dim frmr_recv As Short
    '	Dim frmr_xmit As Short
    '	Dim iframe_recv_err As Short
    '	Dim xmit_aborts As Short
    '	Dim xmit_success As Integer
    '	Dim recv_success As Integer
    '	Dim iframe_xmit_err As Short
    '	Dim recv_buff_unavail As Short
    '	Dim t1_timeouts As Short
    '	Dim ti_timeouts As Short
    '	Dim Reserved1 As Integer
    '	Dim free_ncbs As Short
    '	Dim max_cfg_ncbs As Short
    '	Dim max_ncbs As Short
    '	Dim xmit_buf_unavail As Short
    '	Dim max_dgram_size As Short
    '	Dim pending_sess As Short
    '	Dim max_cfg_sess As Short
    '	Dim max_sess As Short
    '	Dim max_sess_pkt_size As Short
    '	Dim name_count As Short

    '	'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1026"'
    '	Public Sub Initialize()
    '		ReDim adapter_address(5)
    '	End Sub
    'End Structure
    'Public Structure NAME_BUFFER
    '	<VBFixedString(NCBNAMSZ),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr,SizeConst:=NCBNAMSZ)> Public Name As String
    '	Dim name_num As Short
    '	Dim name_flags As Short
    'End Structure
    'Public Structure ASTAT
    '	Dim adapt As ADAPTER_STATUS
    '	<VBFixedArray(30)> Dim NameBuff() As NAME_BUFFER

    '	'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1026"'
    '	Public Sub Initialize()
    '		adapt.Initialize()
    '		ReDim NameBuff(30)
    '	End Sub
    'End Structure
End Module
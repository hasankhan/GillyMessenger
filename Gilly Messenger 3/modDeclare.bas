Attribute VB_Name = "modDeclare"
Option Explicit

'Activate Window
Public Const SW_RESTORE = 9
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

'Special Folders
Public Const CSIDL_PERSONAL = &H5
Public Type SHITEMID
    cb As Long
    abID As Byte
End Type
Public Type ITEMIDLIST
    mkid As SHITEMID
End Type
Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

'Desktop handle
Public Declare Function GetDesktopWindow Lib "user32" () As Long

'Window dimensions
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'Idle time
Public Type LASTINPUTINFO
  cbSize As Long
  dwTime As Long
End Type
Public Declare Function GetLastInputInfo Lib "user32" (pLII As LASTINPUTINFO) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long

'Number textbox
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_STYLE = (-16)
Public Const ES_NUMBER = &H2000&

'Winamp song
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

'Turn off/restart computer
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Const EWX_FORCE = 4
Public Const EWX_REBOOT = 2
Public Const EWX_SHUTDOWN = 1

'Edges
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Const BF_BOTTOM = &H8
Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BDR_RAISEDINNER = &H4

'Files
Public Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

'Windows Information
Public Declare Function GetVersion Lib "kernel32" () As Long

'Topmost window
Public Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Const HWND_TOPMOST = -1
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40

'Popup
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const SPI_GETWORKAREA = 48
Public NewPopupTop As Double
Public LastPopup As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetFocusWindow Lib "user32" Alias "GetFocus" () As Long
Public Declare Function SetFocusWindow Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Public BlockAlertsOnFullScrApp As Boolean

'Transparency
Public Const LWA_ALPHA = &H2
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

'Disable Close Button
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Public Const MF_BYPOSITION = &H400&
Public Const MF_DISABLED = &H2&

'Directory creator
Public Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long

'Progress Bar
Public Declare Function InvertRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

'Control Communication
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    'Window drag
    Public Declare Function ReleaseCapture Lib "user32" () As Long
    Public Const HTBOTTOMRIGHT = 17
    Public Const WM_NCLBUTTONDOWN = &HA1
    'Combo Box
    Public Const CB_FINDSTRING = &H14C
    Public Const CB_SELECTSTRING = &H14D
    'Rich Edit Control
    Public Const WM_COPY = &H301
    Public Const WM_GETTEXTLENGTH = &HE
    Public Const WM_CUT = &H300
    Public Const EM_GETUNDONAME = &H256
    Public Const EM_UNDO = &HC7
    Public Const WM_PASTE = &H302

'Key capture
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

'TempDir path
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

'SysDir path
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'URL detection
Public Declare Function PathIsURL Lib "shlwapi.dll" Alias "PathIsURLA" (ByVal pszPath As String) As Long

'Flash window
Public Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long

'Contact List
Public ContactList As Collection
Public ContactGroups As Collection
Public ContactProperties As Collection
Public ContactComments As Collection
Public ContactCustomNicks As Collection

Public IgnoreList As Collection
Public HiddenContacts As Collection
Public UserProperties As Collection
Public PendingIM As Collection

Public SortContactsByGroups As Boolean
Public ViewContactsByEmail As Boolean
Public GroupOfflineContactsTogether As Boolean

'IM Window
Public IMFontName As String
Public IMFontSize As Integer
Public IMFontColor As Long
Public IMFontBold As Boolean
Public IMFontItalic As Boolean
Public IMFontStrikethru As Boolean
Public IMFontUnderline As Boolean
Public IMFontRandomFormat As Boolean
Public IMFontRandomColors As Boolean
Public TextStyle As String
Public TimeStamp As Boolean
Public EmoticonFloodControl As Boolean
Public IMWindowWidth As Integer
Public IMWindowHeight As Integer
Public IMWindowMax As Boolean
Public IMWindowBackground As IPictureDisp
Public IMWindowTopLeft As IPictureDisp
Public IMWindowTopMid As IPictureDisp
Public IMWindowTopRight As IPictureDisp

Public IMWindows As Collection
Public Emoticons(89, 1) As String
Public IMWindowCommands(46, 1) As String

Public DpTransfers As Collection
Public SendDisplayPic As Boolean
Public ReceiveDisplayPic As Boolean
Public ShowMyDP As Boolean
Public FTPPort As Integer

'Folders
Public ReceivedFilesFolder As String
Public MessageHistoryFolder As String
Public StatusHistoryFolder As String

'Options
Public AlertOnContactOnline As Boolean
Public AlertOnMessageReceived As Boolean
Public AlertOnEmailReceived As Boolean
Public SoundAlerts As Boolean
Public SaveStatusHistory As Boolean
Public ShowEmoticons As Boolean
Public SaveMessageHistory As Boolean
Public ShowIMWindowOnMsg As Boolean
Public TypingNotification As Boolean
Public HighlightFakeFriends As Boolean

'Sounds
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public SoundFilter As Collection
Public SoundFilterMode As Boolean
Public boolOnlineSound As Boolean
Public strOnlineSound As String
Public boolOfflineSound As Boolean
Public strOfflineSound As String
Public boolTypingSound As Boolean
Public strTypingSound As String
Public boolMessageSound As Boolean
Public strMessageSound As String
Public boolEmailSound As Boolean
Public strEmailSound As String
Public boolAlertSound As Boolean
Public strAlertSound As String

'Browser Options
Public strDefaultBrowser As String
Public strCustomBrowser As String
Public boolUseDefaultBrowser As Boolean

'Email App Options
Public strDefaultEmailApp As String
Public strCustomEmailApp As String
Public strCustomEmailWeb As String
Public boolUseCustomEmailWeb As Boolean
Public boolUseDefaultEmailApp As Boolean

'Other
Public MainWindowWidth As Integer
Public MainWindowHeight As Integer
Public MainWindowMax As Boolean

Public MyStatus As Integer
Public LastStatus As Integer
Public InitialStatus As Integer
Public Transparency As Byte
Public SavePassword As Boolean
Public AutoIdle As Boolean
Public AutoIdle_Interval As Integer
Public Status_AutoIdle As Boolean
Public LastActive As Long
Public PopupFilter As Collection
Public PopupFilterMode As Boolean
Public BlockAlert As Boolean

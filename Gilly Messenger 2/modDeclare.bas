Attribute VB_Name = "modDeclare"
'Key Capture
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'Sounds
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
'Winamp Song
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
'Get Temp Directory
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'Window Highlighter
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetFocus Lib "user32" () As Long
Public Declare Function SetFocusX Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Public FocusWnd As Long
'File Execution
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'Control Communication
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_PASTE = &H302
Public Const CB_FINDSTRING = &H14C
Public Const CB_SELECTSTRING = &H14D
Public Const HTBOTTOMRIGHT = 17
Public Const WM_NCLBUTTONDOWN = &HA1
'Directory creator
Public Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long
'Turn off PC
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Const EWX_SHUTDOWN = 1
Public Const EWX_FORCE = 4
Public Const EWX_REBOOT = 2
'Inbox Management
Public InboxUnread As Integer
Public FolderUnread As Integer
Public NewMail_FromName As String
Public NewMail_FromEmail As String
Public NewMail_Subject As String
Public NewMail_Folder As String
Public DumpMail As String
'Chat Logger
Public ChatLogDir As String
'GM Scripting Engine
Public GMSCount As Integer
Public GMSVars As New Collection
'Contact List Management
Public ContactList As New Collection
Public LstBuddy As Collection
Public Const Lst_FL = 1
Public Const Lst_AL = 2
Public Const Lst_BL = 4
Public Const Lst_RL = 8
Public BuddyIgnore As New Collection
Public BuddyComment As New Collection
Public Type Buddy
    Status  As String
    Email As String
    Nick As String
End Type
Public Contact As Buddy
'Remote Control
Public RcUsername As String
Public RcPassword As String
Public RemoteControl As Boolean
Public RcLoggedIn As Boolean
Public RcUser As String
'Top Window
Public Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Const HWND_TOPMOST = -1
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
'Chat Windows
Public ChatColor As Long
Public ChatRandomColors As Boolean
Public ChatFont As String
Public ChatFontSize As Integer
Public ChatFontBold As Boolean
Public ChatFontItalic As Boolean
Public UseEmoticons As Boolean
Public ChatWindowWidth As Integer
Public ChatWindowHeight As Integer
Public ChatWindowMaximized As Boolean
Public ShowIMWindowOnMsg As Boolean
Public Const IMW_TopBarHeight = 58
'Popup
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Const BF_BOTTOM = &H8
Public Const BF_LEFT = &H1
Public Const BF_RIGHT = &H4
Public Const BF_TOP = &H2
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BDR_SUNKENOUTER = &H2
Public Const BDR_SUNKENINNER As Long = &H8
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Const SPI_GETWORKAREA = 48
Public R As RECT
Public PopupHeight As Integer
Public Popups As Boolean
Public LastPopup As Long
'General Variables
Public Login As String
Public Password As String
Public ViewContactsByEmail As Boolean
Public LoginTime As Date
Public Inbox_Sid As Integer
Public Inbox_Kv As Integer
Public Inbox_Id As Integer
Public Inbox_Rru As String
Public Inbox_MSPAuth As String
Public Inbox_Url As String
Public MsnUrlType As New Collection
Public MsnFile As New Collection
Public Nick As String
Public Status As Integer
Public TrialID As Long
Public MsnError As New Collection
Public Emoticons(89, 1) As String
Public SignedIn As Boolean
Public TempPath As String
Public CallForm As Form
Public CallForms As Collection
Public RingForm As Form
Public TypingNotify As Boolean
Public OpenChats As New Collection
Public LastDir As String
Public Handle As Long
Public Temp As String
Public SignInMode As String
Public InitialStatus As Integer
Public StatusLogDir As String
Public InitStatus As Boolean
Public StatusImage As Integer
Public AddContactFrm As Form
Public LastBlockAlert As String
Public LastDeleteAlert As String
Public LastAddAlert As String
Public TempForm As Form
Public AutoMsg As String
Public LastError As String

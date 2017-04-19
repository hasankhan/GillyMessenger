Attribute VB_Name = "modSysTray"
'Tray Icon
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Long
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONUP = &H205

Dim TrayIcon As NOTIFYICONDATA

Public Sub AddIcon()
TrayIcon.cbSize = Len(TrayIcon)
TrayIcon.hwnd = frmMain.picTrayIcon.hwnd
TrayIcon.uId = 1&
TrayIcon.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
TrayIcon.ucallbackMessage = WM_LBUTTONDOWN
TrayIcon.hIcon = frmMain.picTrayIcon.Picture.Handle
TrayIcon.szTip = "Gilly Messenger" & Chr$(0)
Shell_NotifyIcon NIM_ADD, TrayIcon
End Sub

Public Sub DeleteIcon()
Shell_NotifyIcon NIM_DELETE, TrayIcon
End Sub

Public Sub ChangeTip(Tip As String)
TrayIcon.szTip = Tip & Chr$(0)
Shell_NotifyIcon NIM_MODIFY, TrayIcon
End Sub

Public Sub ChangeIcon(PicHWnd As Long)
TrayIcon.hIcon = PicHWnd
Shell_NotifyIcon NIM_MODIFY, TrayIcon
End Sub

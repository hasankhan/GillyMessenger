Attribute VB_Name = "modSysTray"
Option Explicit

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Long
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONUP = &H205

Private TrayIcon As NOTIFYICONDATA

Public Sub AddTrayIcon()
    TrayIcon.cbSize = Len(TrayIcon)
    TrayIcon.hwnd = frmMain.picTrayIcon.hwnd
    TrayIcon.uId = 1&
    TrayIcon.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    TrayIcon.ucallbackMessage = WM_LBUTTONDOWN
    TrayIcon.hIcon = frmMain.imglstTrayIcons.ListImages(IIf(frmMain.objMSN_NS.State = NsState_SignedIn, Choose(MyStatus + 1, 2, 3, 4, 4, 3, 4, 4, 5, 1), 1)).Picture
    TrayIcon.szTip = "Gilly Messenger - Not Signed In" & Chr$(0)
    Shell_NotifyIcon NIM_ADD, TrayIcon
End Sub

Public Sub DelTrayIcon()
    Shell_NotifyIcon NIM_DELETE, TrayIcon
End Sub

Public Sub UpdateTrayIcon()
    TrayIcon.hIcon = frmMain.imglstTrayIcons.ListImages(IIf(frmMain.objMSN_NS.State = NsState_SignedIn, Choose(MyStatus + 1, 2, 3, 4, 4, 3, 4, 4, 5, 1), 1)).Picture
    TrayIcon.szTip = "Gilly Messenger - " & IIf(frmMain.objMSN_NS.State = NsState_SignedIn, frmMain.objMSN_NS.Login, "Not Signed In") & Chr$(0)
    Shell_NotifyIcon NIM_MODIFY, TrayIcon
End Sub

Public Sub ChangeTrayIcon(PicHWnd As Long)
    TrayIcon.hIcon = PicHWnd
    Shell_NotifyIcon NIM_MODIFY, TrayIcon
End Sub

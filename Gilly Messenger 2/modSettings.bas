Attribute VB_Name = "modSettings"
Public Sub LoadAppSettings()
frmMain.Width = GetSetting("Gilly Messenger", "App Settings", "Width", frmMain.Width)
frmMain.Height = GetSetting("Gilly Messenger", "App Settings", "Height", frmMain.Height)
frmMain.Left = (Screen.Width / 2) - (frmMain.Width / 2)
frmMain.Top = (Screen.Height / 2) - (frmMain.Height / 2)
If GetSetting("Gilly Messenger", "App Settings", "Maximized") = "True" Then
    frmMain.WindowState = vbMaximized
End If
ChatWindowWidth = GetSetting("Gilly Messenger", "App Settings", "ChatWindow Width", 5910)
ChatWindowHeight = GetSetting("Gilly Messenger", "App Settings", "ChatWindow Height", 7065)
ChatWindowMaximized = GetSetting("Gilly Messenger", "App Settings", "ChatWindow Maximized", False)
LastDir = GetSetting("Gilly Messenger", "App Settings", "Last Dir", App.Path)
ChatColor = GetSetting("Gilly Messenger", "App Settings", "Chat Color", vbBlack)
ChatFont = GetSetting("Gilly Messenger", "App Settings", "Chat Font", "Tahoma")
ChatFontSize = GetSetting("Gilly Messenger", "App Settings", "Chat Font Size", 10)
ChatFontBold = GetSetting("Gilly Messenger", "App Settings", "Chat Font Bold", False)
ChatFontItalic = GetSetting("Gilly Messenger", "App Settings", "Chat Font Italic", False)
ChatLogDir = GetSetting("Gilly Messenger", "App Settings", "ChatLog Folder", App.Path & "\Chat Logs\")
StatusLogDir = GetSetting("Gilly Messenger", "App Settings", "StatusLog Folder", App.Path & "\Status Logs\")
Popups = GetSetting("Gilly Messenger", "App Settings", "Popups", "True")
TypingNotify = GetSetting("Gilly Messenger", "App Settings", "Typing Notification", True)
ShowIMWindowOnMsg = GetSetting("Gilly Messenger", "App Settings", "Show IMWindow On Message", False)
UseEmoticons = GetSetting("Gilly Messenger", "App Settings", "Use Emoticons", True)
frmMain.mnuChatLogger.Checked = GetSetting("Gilly Messenger", "App Settings", "Chat Logger", False)
frmMain.mnuStatusLogger.Checked = GetSetting("Gilly Messenger", "App Settings", "Status Logger", False)
End Sub

Public Sub SaveAppSettings()
SaveSetting "Gilly Messenger", "App Settings", "Maximized", (frmMain.WindowState = vbMaximized)
frmMain.WindowState = vbNormal
frmMain.Show
SaveSetting "Gilly Messenger", "App Settings", "Width", frmMain.Width
SaveSetting "Gilly Messenger", "App Settings", "Height", frmMain.Height
SaveSetting "Gilly Messenger", "App Settings", "ChatWindow Width", ChatWindowWidth
SaveSetting "Gilly Messenger", "App Settings", "ChatWindow Height", ChatWindowHeight
SaveSetting "Gilly Messenger", "App Settings", "ChatWindow Maximized", ChatWindowMaximized
SaveSetting "Gilly Messenger", "App Settings", "Last Dir", LastDir
SaveSetting "Gilly Messenger", "App Settings", "Chat Color", ChatColor
SaveSetting "Gilly Messenger", "App Settings", "Chat Font", ChatFont
SaveSetting "Gilly Messenger", "App Settings", "Chat Font Size", ChatFontSize
SaveSetting "Gilly Messenger", "App Settings", "Chat Font Bold", ChatFontBold
SaveSetting "Gilly Messenger", "App Settings", "Chat Font Italic", ChatFontItalic
SaveSetting "Gilly Messenger", "App Settings", "Use Emoticons", UseEmoticons
SaveSetting "Gilly Messenger", "App Settings", "Typing Notification", TypingNotify
SaveSetting "Gilly Messenger", "App Settings", "Chat Logger", frmMain.mnuChatLogger.Checked
SaveSetting "Gilly Messenger", "App Settings", "Status Logger", frmMain.mnuStatusLogger.Checked
End Sub

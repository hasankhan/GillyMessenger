Attribute VB_Name = "modGilly"
Option Explicit
Private hForegroundWnd As Long, hFocusWnd As Long, hActiveWnd As Long

Public Function GetKeyState(KeyCode As Byte) As Boolean
    Dim i As Integer
    i = GetAsyncKeyState(KeyCode)
    If i <> 0 And i <> 1 Then
        GetKeyState = True
    End If
End Function

Public Function InList(Lists As Integer, list As Integer) As Boolean
    If (Lists And list) = list Then
        InList = True
    Else
        InList = False
    End If
End Function

Public Function GetContactAttr(Email As String, Key As String)
    On Error Resume Next
    
    If Key = "nick" Then
        GetContactAttr = GetBuddyCustomNick(Email)
        If Not GetContactAttr = vbNullString Then
            Exit Function
        End If
    End If
    
    GetContactAttr = ContactList(Email).Item(Key)
    
    If GetContactAttr = vbNullString Then
        Select Case Key
        Case "nick"
            GetContactAttr = Email
        Case "status"
            GetContactAttr = msnStatus_Unknown
        End Select
    End If
End Function

Public Sub TerminateGM()
    On Error Resume Next
    
    Close
    Call DelTrayIcon
    
    SaveSettingX "App Settings", "MainWindow Width", MainWindowWidth
    SaveSettingX "App Settings", "MainWindow Height", MainWindowHeight
    SaveSettingX "App Settings", "MainWindow Max", MainWindowMax
    SaveSettingX "App Settings", "MainWindow Left", frmMain.Left
    SaveSettingX "App Settings", "MainWindow Top", frmMain.Top
    
    If frmMain.objMSN_NS.State = NsState_SignedIn Then
        Call SaveUserSettings
        Call LogStatus("Signed out.")
    End If
    
    Dim frmTemp As Form
    For Each frmTemp In Forms
        Unload frmTemp
    Next

    End
End Sub

Public Function InCollection(srcCollection As Collection, Key As String) As Boolean
    On Error GoTo Handler
    srcCollection.Item Key
    InCollection = True
    Exit Function
Handler:
    InCollection = False
End Function

Public Function XorEncrypt(ByVal Data As String, Key As String) As String
    Dim i As Integer, j As Integer
    Data = StrReverse(Data)
    For i = 1 To Len(Data)
        j = j + 1
        XorEncrypt = XorEncrypt & Chr$(128 Xor (Asc(Mid$(Data, i, 1)) Xor Asc(Mid$(Key, j, 1)) Xor (j Mod 255)))
        If j = Len(Key) Then
            j = 1
        End If
    Next
End Function

Public Function XorDecrypt(ByVal Data As String, Key As String) As String
    Dim i As Integer, j As Integer
    For i = 1 To Len(Data)
        j = j + 1
        XorDecrypt = XorDecrypt & Chr$(128 Xor (Asc(Mid$(Data, i, 1)) Xor Asc(Mid$(Key, j, 1)) Xor (j Mod 255)))
        If j = Len(Key) Then
            j = 1
        End If
    Next
    XorDecrypt = StrReverse(XorDecrypt)
End Function

Public Sub ActivateWindow(Window As Form)
    On Error Resume Next
    
    With Window
        If Not .Visible Then
            .Visible = True
            If Not Transparency = 0 Then
                SetTransparency Window, Transparency
            End If
        End If
        
        If Window.WindowState = vbMinimized Then
            ShowWindow .hwnd, SW_RESTORE
        End If
        
        Dim Temp As String
        Temp = .Caption
        .Caption = .Caption & " "
        AppActivate .Caption
        .Caption = Temp
    End With
End Sub

Public Function ProperCase(Text As String) As String
    ProperCase = UCase$(Left$(Text, 1)) & Right$(Text, Len(Text) - 1)
End Function

Public Function GetContactComment(Email As String) As String
    On Error Resume Next
    
    GetContactComment = ContactComments(Email)
End Function

Public Function LoadSettingsInCollection(Section As String, Optional Filter As String) As Collection
    On Error GoTo Handler
    Set LoadSettingsInCollection = New Collection
    Dim strTempAry() As String, i As Integer
    strTempAry = GetAllSettings("Gilly Messenger", Section)
    If Not ArraySize(strTempAry) = -1 Then
        For i = 0 To UBound(strTempAry)
            If Not strTempAry(i, 0) Like Filter Then
                LoadSettingsInCollection.Add strTempAry(i, 1), strTempAry(i, 0)
            End If
        Next
    End If
Handler:
    Erase strTempAry
End Function

Public Sub AddSubMenu(srcMenu As Object, Caption As String, Tag As String)
    On Error Resume Next
    
    Load srcMenu(srcMenu.UBound + 1)
    srcMenu(srcMenu.UBound).Caption = Caption
    srcMenu(srcMenu.UBound).Tag = Tag
    srcMenu(srcMenu.UBound).Enabled = True
    srcMenu(srcMenu.UBound).Visible = True
    srcMenu(srcMenu.LBound).Visible = False
End Sub

Public Sub RemoveSubMenu(srcMenu As Object, Tag As String)
    On Error Resume Next
    
    Unload srcMenu(GetSubMenu(srcMenu, Tag))
End Sub

Public Sub ClearSubMenu(srcMenu As Object)
    On Error Resume Next
    
    Dim i As Integer
    srcMenu(srcMenu.LBound).Visible = True
    For i = 1 To srcMenu.UBound
        Unload srcMenu(i)
    Next
End Sub

Public Sub RenameSubMenu(srcMenu As Object, Tag As String, NewCaption As String)
    On Error Resume Next
    
    Dim i As Integer
    For i = 1 To srcMenu.UBound
        If srcMenu(i).Tag = Tag Then
            srcMenu(i).Caption = NewCaption
            Exit Sub
        End If
    Next
End Sub

Public Sub SetCollectionItem(srcCollection As Collection, Key, NewValue)
    On Error Resume Next
    
    srcCollection.Remove Key
    srcCollection.Add NewValue, Key
End Sub

Public Function StartChat(ByVal Email As String, Optional ByVal Message As String, Optional ByVal File As String, Optional Activate As Boolean) As Double
    On Error Resume Next
    
    If InStr(Email, "@") = 0 Then
        Email = Email & "@hotmail.com"
    End If
    
    If InCollection(IMWindows, Email) Then
        If Activate Then
            ActivateWindow IMWindows(Email)
        End If
        If Not Message = vbNullString Then
            SendMsg IMWindows(Email), Message
        End If
        If Not File = vbNullString Then
            StartChat = SendFile(IMWindows(Email), CStr(Split(File, "|")(1)), CStr(Split(File, "|")(0)))
        End If
    ElseIf InCollection(PendingIM, Email) Then
        If Activate Then
            ActivateWindow PendingIM(Email)
        End If
        If Not Message = vbNullString Then
            SendMsg PendingIM(Email), Message
        End If
        If Not File = vbNullString Then
            StartChat = SendFile(PendingIM(Email), CStr(Split(File, "|")(1)), CStr(Split(File, "|")(0)))
        End If
    Else
        If MyStatus = msnStatus_Offline Then
            MsgBox "Not allowed when offline.", vbExclamation
            Exit Function
        End If
    
        Dim frmIM As New frmChat
        Load frmIM
        
        With frmIM
            Dim strContactNick As String
            strContactNick = GetContactAttr(Email, "nick")

            .Caption = strContactNick & " - Conversation"
            .objMSN_SB.Socket = frmMain.Controls.Add("MSWinsock.Winsock", "wskSB" & Fix(Timer) & frmMain.Controls.Count)
            .objMSN_SB.Contact = Email
            .objMSN_SB.Login = frmMain.objMSN_NS.Login
            .lblBuddies.Caption = Email
            .lblStatus.Caption = "[" & Time$ & "] Connecting..."
            
            If Not (GetContactAttr(Email, "status") = msnStatus_Online Or GetContactAttr(Email, "status") = msnStatus_Unknown) Then
                .lblBuddyStatus.Caption = strContactNick & " may or may not reply because his/her status is set to " & StatusName(GetContactAttr(Email, "status")) & "."
                .picBuddyStatus.Visible = True
                .lblBuddyStatus.Visible = True
            End If
            
            If Not GetSettingX("App Settings\" & frmMain.objMSN_NS.Login & "\Show DP", Email, True) Then
                .imgBuddyDP.Width = 1
                .imgBuddyDP.Height = .imgShowHideBuddyDP.Height
                .imgBuddyDP.BorderStyle = vbBSNone
                Set .imgBuddyDP.Picture = LoadPicture(vbNullString)
                Call .Form_Resize
            End If
            .CallingContact = True
            .Show
        End With
        PendingIM.Add frmIM, Email
        frmMain.objMSN_NS.RequestSB
        
        If Not Message = vbNullString Then
            SendMsg frmIM, Message
        End If
        If Not File = vbNullString Then
            StartChat = SendFile(frmIM, CStr(Split(File, "|")(1)), CStr(Split(File, "|")(0)))
        End If
        
        Call QueScript(frmIM, "imwindowopened", ConvArray(frmMain.objMSN_NS.Login, frmMain.objMSN_NS.Nick))
    End If
End Function

Public Sub BlockContact(Email As String)
    If InList(GetContactAttr(Email, "lists"), msnList_Allow) Then
        frmMain.objMSN_NS.RemoveContact msnList_Allow, Email
    End If
    If Not InList(GetContactAttr(Email, "lists"), msnList_Block) Then
        frmMain.objMSN_NS.AddContact msnList_Block, Email, Email
    End If
End Sub

Public Sub UnblockContact(Email As String)
    If InList(GetContactAttr(Email, "lists"), msnList_Block) Then
        frmMain.objMSN_NS.RemoveContact msnList_Block, Email
    End If
    If Not InList(GetContactAttr(Email, "lists"), msnList_Allow) Then
        frmMain.objMSN_NS.AddContact msnList_Allow, Email, Email
    End If
End Sub

Public Sub IgnoreContact(Email As String)
    On Error Resume Next
    
    IgnoreList.Add "", Email
    SaveSettingX "Ignore List\" & frmMain.objMSN_NS.Login, Email, vbNullString
    If InList(GetContactAttr(Email, "lists"), msnList_Forward) Then
        frmMain.RefreshContact Email
    End If
End Sub

Public Sub UnignoreContact(Email As String)
    On Error Resume Next
    
    IgnoreList.Remove Email
    DeleteSetting "Gilly Messenger", "Ignore List\" & frmMain.objMSN_NS.Login, Email
    If InList(GetContactAttr(Email, "lists"), msnList_Forward) Then
        frmMain.RefreshContact Email
    End If
End Sub

Public Sub HideContact(Email As String)
    On Error Resume Next
    
    HiddenContacts.Add "", Email
    SaveSettingX "Hide List\" & frmMain.objMSN_NS.Login, Email, vbNullString
    Call frmMain.RemoveContact(Email)
End Sub

Public Sub UnhideContact(Email As String)
    On Error Resume Next
    
    HiddenContacts.Remove Email
    DeleteSetting "Gilly Messenger", "Hide List\" & frmMain.objMSN_NS.Login, Email
    Call frmMain.RefreshContact(Email)
End Sub

Public Function GetUserFile(Optional Filter As String, Optional DialogTitle As String, Optional Action As Integer) As String
    On Error GoTo Handler
    With frmMain.CommonDialog
        .FileName = vbNullString
        .DialogTitle = DialogTitle
        .Filter = Filter
        Select Case Action
        Case 0
            .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
            .ShowOpen
        Case 1
            .Flags = cdlOFNHideReadOnly
            .ShowSave
        End Select
        GetUserFile = .FileName
    End With
    Exit Function
Handler:
    MsgBox Err.Description, vbExclamation
End Function

Public Function SendFile(IMWindow As frmChat, FilePath As String, Optional FileTitle As String, Optional Cookie As Double) As Double
    On Error Resume Next
    If IMWindow.ChatBuddies.Count > 1 Then
        IMWindow.Comment "You can not send a file in multiple contact conversation."
        Cookie = 0
    ElseIf FileExists(FilePath) Then
        If FileTitle = vbNullString Then
            FileTitle = Right$(FilePath, Len(FilePath) - InStrRev(FilePath, "\"))
        End If
        If Cookie = 0 Then
            Cookie = Fix(Timer)
        End If
        With IMWindow
            .Comment "Waiting for " & .BuddyNick & " to accept """ & FileTitle & """."
    
            Select Case .objMSN_SB.State
            Case SbState_Connected
                If Not (.ChatBuddies.Count = 0) Then
                    Dim TransferForm As New frmTransfer
                    Load TransferForm
                    With TransferForm
                        Set .Parent = IMWindow
                        If IMWindow.WindowState = vbMinimized Then
                            ShowWindow IMWindow.hwnd, SW_RESTORE
                        End If
                        .Cookie = Cookie
                        .objMSN_FTP.FilePath = FilePath
                        .objMSN_FTP.File = FileTitle
                        .objMSN_FTP.FileSize = FileLen(FilePath)
                        .objMSN_FTP.Login = IMWindow.objMSN_SB.Contact
                        .objMSN_FTP.TransferType = FtpTransferType_Send
                        .objMSN_FTP.AuthCookie = Fix(Rnd(Timer) * 4294967295#) + 1
                        .lblFileName = "File name: " & .objMSN_FTP.File
                        .lblFileSize = "File size: " & ConvertBytes(.objMSN_FTP.FileSize)
                        .lblTransfer = "Sending to: " & IMWindow.objMSN_SB.Contact
                        .cmdAccept.Visible = False
                        .Show , IMWindow
                        Call SetRandomPos(IMWindow, TransferForm)
                    End With
                    IMWindow.Invitations.Add TransferForm, "Cookie " & IMWindow.objMSN_SB.TrialID
                    IMWindow.objMSN_SB.SendInvitation "File Transfer", "{5D3E02AB-6190-11d3-BBBB-00C04F795683}", IMWindow.objMSN_SB.TrialID, "Application-File: " & FileTitle, "Application-FileSize: " & TransferForm.objMSN_FTP.FileSize
                    Set TransferForm = Nothing
                    
                ElseIf Not .CallingContact Then
                    .CallingContact = True
                    .lblStatus.Caption = "[" & Time$ & "] Reconnecting..."
                    .FileQue.Add FileTitle & "|" & FilePath & "|" & Cookie
                    .objMSN_SB.InviteContact .objMSN_SB.Contact
                End If
            Case SbState_Connecting
                .FileQue.Add FileTitle & "|" & FilePath & "|" & Cookie
            Case SbState_Disconnected
                .FileQue.Add FileTitle & "|" & FilePath & "|" & Cookie
                If Not .CallingContact Then
                    .CallingContact = True
                    PendingIM.Add IMWindow, IMWindow.objMSN_SB.Contact
                    frmMain.objMSN_NS.RequestSB
                End If
            End Select
        End With
        Call QueScript(IMWindow, "transferrequest", ConvArray(frmMain.objMSN_NS.Login, FileTitle, FileLen(FilePath), Cookie))
    End If
    SendFile = Cookie
End Function

Public Function ConvertBytes(Bytes As Double, Optional Rate As Boolean) As String
  If Bytes >= 1073741824 Then
      ConvertBytes = Format(Bytes / 1024 / 1024 / 1024, IIf(Rate, "#0.0", "#0.00")) & " " & IIf(Rate, "Gbps", "GB")
  ElseIf Bytes >= 1048576 Then
      ConvertBytes = Format(Bytes / 1024 / 1024, IIf(Rate, "#0.0", "#0.00")) & " " & IIf(Rate, "Mbps", "MB")
  ElseIf Bytes >= 1024 Then
      ConvertBytes = Format(Bytes / 1024, IIf(Rate, "#0.0", "#0.00")) & " " & IIf(Rate, "Kbps", "KB")
  Else
      ConvertBytes = Fix(Bytes) & " " & IIf(Rate, "bps", "Bytes")
  End If
End Function

Public Sub DisableCloseButton(hwnd As Long)
    Dim hMenu As Long
    Dim nCount As Long
    hMenu = GetSystemMenu(hwnd, 0)
    nCount = GetMenuItemCount(hMenu)
    RemoveMenu hMenu, nCount - 1, MF_DISABLED Or MF_BYPOSITION
    DrawMenuBar hwnd
End Sub

Public Sub SendMsg(IMWindow As frmChat, Message As String, Optional Visible As Boolean = True)
    On Error Resume Next
    
    With IMWindow
    
        If Not LCase$(Left$(Message, 6)) = "/text " Then
            Message = Alias(Message, IMWindow)
        End If
        
        Dim CmdParams() As String, Param As String, i As Integer
        CmdParams = Split(Message)
        If UBound(CmdParams) > 0 Then
            Param = Right$(Message, Len(Message) - InStr(Message, " "))
        End If
        
        Select Case LCase$(CmdParams(0))
        Case "cls"
            .txtChat.Text = vbNullString
        Case "/vanish"
            .Visible = False
        Case "/invite"
            If Not Param = vbNullString Then
                .objMSN_SB.InviteContact CmdParams(1)
            End If
        Case "/block"
            If Param = vbNullString Then
                BlockContact .objMSN_SB.Contact
            Else
                BlockContact CmdParams(1)
            End If
        Case "/unblock"
            If Param = vbNullString Then
                UnblockContact .objMSN_SB.Contact
            Else
                UnblockContact CmdParams(1)
            End If
        Case "/ignore"
            If Param = vbNullString Then
                IgnoreContact .objMSN_SB.Contact
            Else
                IgnoreContact CmdParams(1)
            End If
        Case "/unignore"
            If Param = vbNullString Then
                UnignoreContact .objMSN_SB.Contact
            Else
                UnignoreContact CmdParams(1)
            End If
        Case "/profile"
            If Param = vbNullString Then
                Call WebNavigate("http://members.msn.com/" & .objMSN_SB.Contact)
            Else
                Call WebNavigate("http://members.msn.com/" & CmdParams(1))
            End If
        Case "/properties"
            If Param = vbNullString Then
                ShowBuddyProperties IMWindow, .objMSN_SB.Contact, GetCustomNick(.ChatBuddies(1).Item("email"), .ChatBuddies(1).Item("nick"))
            Else
                If InCollection(.ChatBuddies, CmdParams(1)) Then
                    ShowBuddyProperties IMWindow, CmdParams(1), GetCustomNick(.ChatBuddies(CmdParams(1)).Item("email"), .ChatBuddies(CmdParams(1)).Item("nick"))
                Else
                    ShowBuddyProperties IMWindow, CmdParams(1)
                End If
            End If
        Case "/list"
            If Param = vbNullString Then
                If Not .ChatBuddies.Count = 0 Then
                    Dim strList As String
                    strList = .ChatBuddies(1).Item("email") & " - " & .ChatBuddies(1).Item("nick")
                    For i = 2 To .ChatBuddies.Count
                        strList = strList & vbCrLf & .ChatBuddies(i).Item("email") & " - " & .ChatBuddies(i).Item("nick")
                    Next
                    .Comment strList, , False
                Else
                    .Comment "No one is in the conversation.", , False
                End If
            ElseIf LCase$(CmdParams(1)) = "online" Then
                .Comment OnlineList, , False
            End If
        Case "/view"
            If Not Param = vbNullString Then
                Select Case LCase$(CmdParams(1))
                Case "log"
                    If UBound(CmdParams) = 1 Then
                        ShellExecute 0, "open", MessageHistoryFolder & "\" & .objMSN_SB.Contact & ".txt", vbNullString, vbNullString, 1
                    ElseIf IsEmail(CmdParams(2)) And UBound(CmdParams) = 2 Then
                        ShellExecute 0, "open", MessageHistoryFolder & "\" & CmdParams(2) & ".txt", vbNullString, vbNullString, 1
                    End If
                Case "comment"
                    If UBound(CmdParams) = 1 Then
                        .Comment GetBuddyComment(.objMSN_SB.Contact), , False
                    ElseIf IsEmail(CmdParams(2)) And UBound(CmdParams) = 2 Then
                        .Comment GetBuddyComment(CmdParams(2)), , False
                    Else
                        .Comment GetBuddyComment(.objMSN_SB.Contact), , False
                    End If
                Case "info"
                    If UBound(CmdParams) = 1 Then
                        .Comment GetBuddyInfo(.objMSN_SB.Contact), , False
                    ElseIf IsEmail(CmdParams(2)) And UBound(CmdParams) = 2 Then
                        .Comment GetBuddyInfo(CmdParams(2)), , False
                    Else
                        .Comment GetBuddyInfo(.objMSN_SB.Contact), , False
                    End If
                End Select
            End If
        Case "/comment"
            If UBound(CmdParams) = 1 Then
                SetBuddyComment .objMSN_SB.Contact, Param
            ElseIf IsEmail(CmdParams(1)) And UBound(CmdParams) > 1 Then
                If UBound(CmdParams) = 2 Then
                    SetBuddyComment CmdParams(1), Param
                Else
                    SetBuddyComment CmdParams(1), Right$(Param, Len(Param) - InStr(Param, " "))
                End If
            Else
                SetBuddyComment .objMSN_SB.Contact, Param
            End If
        Case "/customnick"
            If UBound(CmdParams) <= 1 Then
                SetBuddyCustomNick .objMSN_SB.Contact, Param
            ElseIf IsEmail(CmdParams(2)) And UBound(CmdParams) > 1 Then
                If UBound(CmdParams) = 2 Then
                    SetBuddyCustomNick CmdParams(1), Param
                Else
                    SetBuddyCustomNick CmdParams(1), Right$(Param, Len(Param) - InStr(Param, " "))
                End If
            Else
                SetBuddyCustomNick .objMSN_SB.Contact, Param
            End If
        Case "/addcontact"
            If UBound(CmdParams) = 0 Then
                AddContact .objMSN_SB.Contact
            ElseIf IsEmail(CmdParams(1)) Then
                If UBound(CmdParams) = 1 Then
                    AddContact CmdParams(1)
                Else
                    AddContact CmdParams(1), Right$(Param, Len(Param) - InStr(Param, " "))
                End If
            Else
                AddContact .objMSN_SB.Contact, Param
            End If
        Case "/find"
            If Not Param = vbNullString Then
                .txtChat.Find Param, Val(.txtChat.Tag)
            End If
        Case "/email"
            If Not Param = vbNullString Then
                If IsEmail(Param) Then
                    Call SendEmail(Param)
                Else
                    Call SendEmail(Param & "@hotmail.com")
                End If
            Else
                Call SendEmail(.objMSN_SB.Contact)
            End If
        Case "/nick"
            If Not Param = vbNullString Then
                frmMain.objMSN_NS.ChangeNick Param
            End If
        Case "/fakenick"
            If Param = vbNullString Then
                .mnuTools_FakeNick.Checked = False
            Else
                .mnuTools_FakeNick.Tag = Param
                .mnuTools_FakeNick.Checked = True
            End If
        Case "/imitate"
            If Not Param = vbNullString Then
                If Param = "off" Then
                    Call .RestoreMsgFont
                    .mnuTools_FakeNick.Checked = False
                Else
                    Call .Imitate(Param)
                End If
            Else
                Call .Imitate(.objMSN_SB.Contact)
            End If
        Case "/font"
            .txtMessage.FontBold = False
            .txtMessage.FontItalic = False
            .txtMessage.FontStrikethru = False
            .txtMessage.FontUnderline = False
            Dim FontAttrs() As String
            FontAttrs = Split(Param)
            For i = 0 To ArraySize(FontAttrs)
                Select Case LCase$(FontAttrs(i))
                Case "bold"
                    .txtMessage.FontBold = True
                Case "italic"
                    .txtMessage.FontItalic = True
                Case "strikethru"
                    .txtMessage.FontStrikethru = True
                Case "underline"
                    .txtMessage.FontUnderline = True
                Case Else
                    If Not FontAttrs(i) = vbNullString Then
                        .txtMessage.FontName = Right$(Param, Len(Param) - InStr(Param, FontAttrs(i)) + 1)
                        Exit For
                    End If
                End Select
            Next
            IMFontName = .txtMessage.FontName
            IMFontBold = .txtMessage.FontBold
            IMFontItalic = .txtMessage.FontItalic
            IMFontStrikethru = .txtMessage.FontStrikethru
            IMFontUnderline = .txtMessage.FontUnderline
        Case "/color"
            If Not Param = vbNullString Then
                .txtMessage.ForeColor = ColorConv(CmdParams(1))
            End If
        Case "/automsg"
            If Not Param = vbNullString Then
                frmMain.mnuTools_AutoMessage.Tag = CmdParams(1)
                frmMain.mnuTools_AutoMessage.Checked = True
            Else
                frmMain.mnuTools_AutoMessage.Checked = False
            End If
        Case "/online"
            frmMain.objMSN_NS.ChangeStatus msnStatus_Online
        Case "/busy"
            frmMain.objMSN_NS.ChangeStatus msnStatus_Busy
        Case "/brb"
            frmMain.objMSN_NS.ChangeStatus msnStatus_BeRightBack
        Case "/away"
            frmMain.objMSN_NS.ChangeStatus msnStatus_Away
            If Not Param = vbNullString Then
                frmMain.mnuTools_AutoMessage.Tag = CmdParams(1)
                frmMain.mnuTools_AutoMessage.Checked = True
            End If
        Case "/phone"
            frmMain.objMSN_NS.ChangeStatus msnStatus_OneThePhone
        Case "/lunch"
            frmMain.objMSN_NS.ChangeStatus msnStatus_OutToLunch
        Case "/idle"
            frmMain.objMSN_NS.ChangeStatus msnStatus_Idle
        Case "/hide"
            frmMain.objMSN_NS.ChangeStatus msnStatus_Offline
        Case "/chat"
            If Not Param = vbNullString Then
                StartChat CmdParams(1)
            End If
        Case "/msg"
            If UBound(CmdParams) > 1 Then
                StartChat CmdParams(1), Right$(Param, Len(Param) - InStr(Param, " "))
            End If
        Case "/msgall"
            If Not Param = vbNullString Then
                Call MessageAll(Param)
            End If
        Case "/sendfile"
            SendFile IMWindow, Param
        Case "/signout"
            Call Signout
        Case "/signin"
            Call SignIn(CmdParams(1), CmdParams(2))
        Case "/ver"
            .Comment "Gilly Messenger " & App.Major & "." & App.Minor & "." & App.Revision
        Case "/msgr"
            ActivateWindow frmMain
        Case "/close"
            Unload IMWindow
        Case "/script"
            If Not Param = vbNullString Then
                Dim ScriptParams() As String
                ScriptParams = Split(Param, ",")
                For i = 0 To UBound(ScriptParams)
                    ScriptParams(i) = TrimX(ScriptParams(i))
                Next
                If UBound(ScriptParams) = 0 Then
                    Call ToggleScript(Param)
                Else
                    Call ToggleScript(ScriptParams(0), , SubArray(ScriptParams, 1, UBound(ScriptParams)))
                End If
            End If
        Case "/bot"
            Call LoadBot(vbNullString)
        Case "/execute"
            Call ShellExecuteEx(Param)
        Case "/buzz"
            Dim TempFont(5) As String
            TempFont(0) = .txtMessage.FontName
            TempFont(1) = .txtMessage.ForeColor
            TempFont(2) = .txtMessage.FontBold
            TempFont(3) = .txtMessage.FontItalic
            TempFont(4) = .txtMessage.FontStrikethru
            TempFont(5) = .txtMessage.FontUnderline
            .txtMessage.FontName = "Verdana"
            .txtMessage.ForeColor = vbRed
            .txtMessage.FontBold = True
            .txtMessage.FontItalic = False
            .txtMessage.FontStrikethru = False
            .txtMessage.FontUnderline = False
            SendMsg IMWindow, "BUZZ!!!", True
            .txtMessage.FontName = TempFont(0)
            .txtMessage.ForeColor = Val(TempFont(1))
            .txtMessage.FontBold = TempFont(2)
            .txtMessage.FontItalic = TempFont(3)
            .txtMessage.FontStrikethru = TempFont(4)
            .txtMessage.FontUnderline = TempFont(5)
            Erase TempFont
        Case "/exit"
            Call TerminateGM
        Case Else
            If LCase$(Left$(Message, 6)) = "/text " Then
                Message = Right$(Message, Len(Message) - 6)
            End If
            
            If Visible Then
                .LastMsg = Message
                If .mnuTools_RandomFormat.Checked Then
                    Randomize Timer
                    .txtMessage.FontBold = CBool(Fix(Rnd() * 2))
                    .txtMessage.FontItalic = CBool(Fix(Rnd() * 2))
                End If
                If .mnuTools_RandomColors.Checked Then
                    Randomize Timer
                    .txtMessage.ForeColor = RGB(Fix(Rnd() * 256), Fix(Rnd() * 256), Fix(Rnd() * 256))
                End If
            End If
            
            If Not .mnuTools_TextStyler.Tag = vbNullString Then
                Dim strStyle As String
                For i = 1 To .TextStyler.Count
                    strStyle = .TextStyler(i)
                    Message = Replace$(Message, Left$(strStyle, InStr(strStyle, "=") - 1), Right$(strStyle, Len(strStyle) - InStr(strStyle, "=")), , , vbBinaryCompare)
                Next
                Message = Alias(Message, IMWindow)
            End If
            
            If Visible Then
                .AddChat IIf(.mnuTools_FakeNick.Checked, Alias(.mnuTools_FakeNick.Tag, IMWindow), frmMain.objMSN_NS.Nick), .txtMessage.FontName, .txtMessage.ForeColor, .txtMessage.FontBold, .txtMessage.FontItalic, .txtMessage.FontStrikethru, .txtMessage.FontUnderline, Message
            End If
            
            Select Case .objMSN_SB.State
            Case SbState_Connected
                If Not (.ChatBuddies.Count = 0) Then
                    If .mnuTools_Encryption.Checked Then
                        .objMSN_SB.SendMessage XorEncrypt(Message, frmMain.objMSN_NS.Login), .txtMessage.FontName, .txtMessage.ForeColor, .txtMessage.FontBold, .txtMessage.FontItalic, .txtMessage.FontStrikethru, .txtMessage.FontUnderline, IIf(.mnuTools_FakeNick.Checked, Alias(.mnuTools_FakeNick.Tag, IMWindow), vbNullString)
                    Else
                        .objMSN_SB.SendMessage Message, .txtMessage.FontName, .txtMessage.ForeColor, .txtMessage.FontBold, .txtMessage.FontItalic, .txtMessage.FontStrikethru, .txtMessage.FontUnderline, IIf(.mnuTools_FakeNick.Checked, Alias(.mnuTools_FakeNick.Tag, IMWindow), vbNullString)
                    End If
                    
                Else
                    .MessageQue.Add Message
                    If Not .CallingContact Then
                        .CallingContact = True
                        .lblStatus.Caption = "[" & Time$ & "] Reconnecting..."
                        Call LogChat(.objMSN_SB.Contact, "----" & vbCrLf & "[" & Now & "] Reconnecting..." & vbCrLf & "----")
                        .objMSN_SB.InviteContact .objMSN_SB.Contact
                    End If
                End If
            Case SbState_Connecting
                .MessageQue.Add Message
            Case SbState_Disconnected
                .MessageQue.Add Message
                If Not .CallingContact Then
                    .CallingContact = True
                    .lblStatus.Caption = "[" & Time$ & "] Reconnecting..."
                    Call LogChat(.objMSN_SB.Contact, "----" & vbCrLf & "[" & Now & "] Reconnecting..." & vbCrLf & "----")
                    PendingIM.Add IMWindow, .objMSN_SB.Contact
                    frmMain.objMSN_NS.RequestSB
                End If
            End Select
            
            If Visible Then
                Call LogChat(.objMSN_SB.Contact, "[" & Now & "] " & frmMain.objMSN_NS.Nick & " : " & vbCrLf & Space$(3) & Message)
            
                If Not .MsgSentProc Then
                    Call QueScript(IMWindow, "messagesent", ConvArray(IMWindow.objMSN_SB.Contact, IMWindow.txtMessage.FontName, IMWindow.txtMessage.ForeColor, IMWindow.txtMessage.FontBold, IMWindow.txtMessage.FontItalic, IMWindow.txtMessage.FontStrikethru, IMWindow.txtMessage.FontUnderline, Message))
                End If
                
                If Not .FirstMsgReceived Then
                    .FirstMsgReceived = True
                End If
            End If
        End Select
    End With
End Sub

Public Sub AddContact(Email As String, Optional Nick As String)
    If Nick = vbNullString Then
        Nick = GetContactAttr(Email, "nick")
    End If
    If Not InList(GetContactAttr(Email, "lists"), msnList_Allow) Then
        frmMain.objMSN_NS.AddContact msnList_Allow, Email, Email
    End If
    If Not InList(GetContactAttr(Email, "lists"), msnList_Forward) Then
        frmMain.objMSN_NS.AddContact msnList_Forward, Email, Nick, 0
    End If
End Sub

Public Sub SetTransparency(Window As Form, ByVal Value As Byte)
    On Error Resume Next
    
    Value = 255 - Value
    Dim lngStyle As Long
    If Value = 255 Then
        lngStyle = GetWindowLong(Window.hwnd, GWL_EXSTYLE)
        SetWindowLong Window.hwnd, GWL_EXSTYLE, lngStyle And Not WS_EX_LAYERED
    Else
        lngStyle = GetWindowLong(Window.hwnd, GWL_EXSTYLE)
        SetWindowLong Window.hwnd, GWL_EXSTYLE, lngStyle Or WS_EX_LAYERED
        SetLayeredWindowAttributes Window.hwnd, 0, Value, LWA_ALPHA
    End If
End Sub

Public Sub ShowPopup(Source As Form, Action As String, Message As String)
    On Error Resume Next
    
    If Not FullScrApp Then
        Dim DesktopRect As RECT, Popup As New frmPopup, i As Integer, j As Integer
        
        SystemParametersInfo SPI_GETWORKAREA, 0, DesktopRect, 0
        
        If (NewPopupTop - 1755) < DesktopRect.Top Then
            NewPopupTop = DesktopRect.Bottom * Screen.TwipsPerPixelY
        End If
        
        Popup.Left = (DesktopRect.Right * Screen.TwipsPerPixelX) - 2730
        Popup.Top = NewPopupTop
        Popup.Height = 1
        Popup.lblMessage.MouseIcon = frmMain.picSignIn.MouseIcon
        
        LastPopup = Popup.hwnd
        NewPopupTop = NewPopupTop - 1755
        
        Popup.Tag = Action
        Popup.lblMessage.Caption = Message
        
        Call SaveFocus
        
        Set Popup.Source = Source
        
        Popup.Show
        SetWindowPos Popup.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
        
        Call RestoreFocus
        
        j = Popup.Top
        For i = 1 To 1755 Step 25
            Popup.Height = i
            Popup.Top = j - i
            DoEvents
            Sleep 1
        Next
    End If
End Sub

Public Sub OpenMsnURL(rru As String, URL As String, ID As Integer)
    'On Error Resume Next
    
    Call WebNavigate(URL & "&mode=ttl" & _
    "&login=" & SubString(frmMain.objMSN_NS.Login, "", "@") & _
    "&username=" & frmMain.objMSN_NS.Login & _
    "&sid=" & frmMain.objMSN_NS.sid & _
    "&kv=" & frmMain.objMSN_NS.kv & _
    "&id=" & ID & _
    "&sl=" & frmMain.objMSN_NS.sl & _
    "&rru=" & rru & _
    "&auth=" & URL_Encode(frmMain.objMSN_NS.MSPAuth) & _
    "&creds=" & MD5Encrypt(frmMain.objMSN_NS.MSPAuth & frmMain.objMSN_NS.sl & frmMain.objMSN_NS.Password) & _
    "&js=yes")
End Sub

Public Sub UpdateContactStatusInIMWindow(Email As String)
    Dim Status As Integer, strContactNick As String
    Status = GetContactAttr(Email, "status")
    strContactNick = GetContactAttr(Email, "nick")
    
    If InCollection(IMWindows, Email) Then
        With IMWindows(Email)
            If Not Status = msnStatus_Online Then
                If Status = msnStatus_Offline Then
                    .lblBuddyStatus.Caption = strContactNick & " appears to be offline and may not reply."
                Else
                    .lblBuddyStatus.Caption = strContactNick & " may or may not reply because his/her status is set to " & StatusName(Status) & "."
                End If
                .picBuddyStatus.Visible = True
                .lblBuddyStatus.Visible = True
            Else
                .picBuddyStatus.Visible = False
                .lblBuddyStatus.Visible = False
            End If
            Call .Form_Resize
        End With
    
    ElseIf InCollection(PendingIM, Email) Then
        With PendingIM(Email)
            If Not Status = msnStatus_Online Then
                If Status = msnStatus_Offline Then
                    .lblBuddyStatus.Caption = strContactNick & " appears to be offline and may not reply."
                Else
                    .lblBuddyStatus.Caption = strContactNick & " may or may not reply because his/her status is set to " & StatusName(Status) & "."
                End If
                .picBuddyStatus.Visible = True
                .lblBuddyStatus.Visible = True
            Else
                .picBuddyStatus.Visible = False
                .lblBuddyStatus.Visible = False
            End If
            Call .Form_Resize
        End With
    End If
End Sub

Public Sub FlashWindowEx(hwnd As Long)
    If Not GetForegroundWindow = hwnd Then
        FlashWindow hwnd, True
    End If
End Sub

Public Function OnlineList() As String
    Dim i As Integer, strEmail As String, intContactCount As Integer
    For i = 1 To ContactList.Count
        strEmail = ContactList(i).Item("email")
        
        If IsOnline(ContactList(i).Item("email")) Then
            intContactCount = intContactCount + 1
            OnlineList = OnlineList & intContactCount & ". " & strEmail & " - " & ContactList(i).Item("nick") & " " & GetContactState(strEmail) & vbCrLf
        End If
    Next
    If Right(OnlineList, 2) = vbCrLf Then
        OnlineList = Left$(OnlineList, Len(OnlineList) - 2)
    End If
End Function

Public Function ChatList() As String
    On Error Resume Next
    
    Dim i As Integer, strEmail As String, intContactCount As Integer
    For i = 1 To PendingIM.Count
        strEmail = PendingIM(i).objMSN_SB.Contact
        ChatList = ChatList & "Email: " & strEmail & vbCrLf & _
        "Nick: " & GetContactAttr(strEmail, "nick") & vbCrLf & _
        "State: " & GetContactState(strEmail) & vbCrLf & _
        "Status: " & PendingIM(i).lblStatus.Caption & vbCrLf & vbCrLf
    Next
    For i = 1 To IMWindows.Count
        strEmail = IMWindows(i).objMSN_SB.Contact
        ChatList = ChatList & "Email: " & strEmail & vbCrLf & _
        "Nick: " & IMWindows(i).ChatBuddies(strEmail).Item("nick") & vbCrLf & _
        "State: " & GetContactState(strEmail) & vbCrLf & _
        "Status: " & IMWindows(i).lblStatus.Caption & vbCrLf & vbCrLf
    Next
    ChatList = Left$(ChatList, Len(ChatList) - 4)
End Function

Public Function IsOnline(Email As String) As Boolean
    Dim intStatus As Integer
    intStatus = GetContactAttr(Email, "status")
    If Not (intStatus = msnStatus_Offline Or intStatus = msnStatus_Unknown) Then
        IsOnline = True
    End If
End Function

Public Function ColorConv(HexCode As String) As Long
    On Error Resume Next
    
    If Len(HexCode) < 6 Then
        HexCode = HexCode & String$(6 - Len(HexCode), "0")
    ElseIf Len(HexCode) > 6 Then
        HexCode = Left$(HexCode, 6)
    End If
    ColorConv = RGB(Val("&H" & Left$(HexCode, 2)), Val("&H" & Mid$(HexCode, 3, 2)), Val("&H" & Right$(HexCode, 2)))
End Function

Public Sub MessageAll(Message As String, Optional Visible As Boolean = True)
    Dim i As Integer
    For i = 1 To IMWindows.Count
        SendMsg IMWindows(i), Message, Visible
    Next
    For i = 1 To PendingIM.Count
        SendMsg PendingIM(i), Message, Visible
    Next
End Sub

Public Sub ResizeScrollLabelSet(srcPictureBox As PictureBox, srcLabel As Label, srcVScrollBar As VScrollBar)
    srcVScrollBar.Left = srcPictureBox.Width - srcVScrollBar.Width
    
    Dim intWidth As Integer, intHeight As Integer
    intWidth = srcPictureBox.TextWidth(srcLabel.Caption)
    intHeight = srcPictureBox.TextHeight(srcLabel.Caption)
    
    If intWidth > srcPictureBox.Width Then
        srcLabel.Width = srcPictureBox.Width - srcVScrollBar.Width - 4
        srcLabel.Height = ((intWidth \ srcPictureBox.Width) * intHeight) + IIf(((intWidth / srcPictureBox.Width) Mod intHeight) = 0, 0, intHeight)
        srcVScrollBar.Max = (srcLabel.Height / intHeight) - 1
        srcVScrollBar.Visible = True
    Else
        srcVScrollBar.Value = 0
        srcLabel.Width = srcPictureBox.Width
        srcLabel.Height = intHeight
        srcVScrollBar.Visible = False
    End If
End Sub

Public Function PopupMessage(ByVal strText As String, Multiline As Boolean) As String
    On Error Resume Next
    
    If Multiline Then
        If Len(strText) > 50 Then
            strText = Left$(strText, 50)
            PopupMessage = Left$(strText, 25) & vbCrLf & Right$(strText, 25) & "..."
        ElseIf Len(strText) > 25 Then
            PopupMessage = Left$(strText, 25) & vbCrLf & Right$(strText, Len(strText) - 25)
        Else
            PopupMessage = strText
        End If
    Else
        If Len(strText) > 25 Then
            PopupMessage = Left$(strText, 22) & "..."
        Else
            PopupMessage = strText
        End If
    End If
End Function

Public Function CropText(Window As Form, Width As Integer, Text As String, Optional Tail As String)
    If Window.TextWidth(Text & Tail) > Width Then
        Dim i As Integer
        For i = Len(Text) To 1 Step -3
            If Window.TextWidth(Left$(Text, i) & "..." & Tail) <= Width Then
                CropText = Left$(Text, i) & "..." & Tail
                Exit Function
            End If
        Next
    Else
        CropText = Text & Tail
    End If
End Function

Public Sub LoadFileMenu(FileMenu As Object, Path As String, Pattern As String)
    Dim strFile As String
    strFile = Dir$(Path & Pattern)
    Do Until strFile = vbNullString
        AddSubMenu FileMenu, Left$(strFile, InStrRev(strFile, ".") - 1), LCase$(Path & strFile)
        strFile = Dir$()
    Loop
End Sub

Public Function Alias(Text As String, Optional IMWindow As frmChat) As String
    Alias = Text
    Alias = Replace$(Alias, "(time)", Time, , , vbTextCompare)
    Alias = Replace$(Alias, "(stime)", Format(Time$, "HH:MM AM/PM"), , , vbTextCompare)
    Alias = Replace$(Alias, "(date)", Date$, , , vbTextCompare)
    Alias = Replace$(Alias, "(now)", Now, , , vbTextCompare)
    Alias = Replace$(Alias, "(day)", WeekdayName(Weekday(Now)), , , vbTextCompare)
    Alias = Replace$(Alias, "(email)", frmMain.objMSN_NS.Login, , , vbTextCompare)
    Alias = Replace$(Alias, "(nick)", frmMain.objMSN_NS.Nick, , , vbTextCompare)
    Alias = Replace$(Alias, "(status)", MyStatus, , , vbTextCompare)
    Alias = Replace$(Alias, "(myip)", frmMain.wskNS.LocalIP, , , vbTextCompare)
    Alias = Replace$(Alias, "(crlf)", vbCrLf, , , vbTextCompare)
    Alias = Replace$(Alias, "(ver)", "Gilly Messenger " & App.Major & "." & App.Minor & "." & App.Revision, , , vbTextCompare)
    If Not InStr(1, Alias, "(song)", vbTextCompare) = 0 Then
        Alias = Replace$(Alias, "(song)", GetCurrentSong, , , vbTextCompare)
    End If
    Dim FontColor As String
    If IMWindow Is Nothing Then
        Alias = Replace$(Alias, "(fontname)", IMFontName, , , vbTextCompare)
        FontColor = Hex(IMFontColor)
        FontColor = String$(6 - Len(FontColor), "0") & FontColor
        FontColor = Right$(FontColor, 2) & Mid$(FontColor, 3, 2) & Left$(FontColor, 2)
        Alias = Replace$(Alias, "(font)", IIf(IMFontBold, "Bold ", vbNullString) & IIf(IMFontItalic, "Italic ", vbNullString) & IIf(IMFontStrikethru, "Strikethru ", vbNullString) & IIf(IMFontUnderline, "Underline ", vbNullString) & IMFontName, , , vbTextCompare)
        Alias = Replace$(Alias, "(fontcolor)", FontColor, , , vbTextCompare)
        Alias = Replace$(Alias, "(fontbold)", IMFontBold, , , vbTextCompare)
        Alias = Replace$(Alias, "(fontitalic)", IMFontItalic, , , vbTextCompare)
        Alias = Replace$(Alias, "(fontstrikethru)", IMFontStrikethru, , , vbTextCompare)
        Alias = Replace$(Alias, "(fontunderline)", IMFontUnderline, , , vbTextCompare)
    Else
        With IMWindow
            Alias = Replace$(Alias, "(fontname)", .txtMessage.FontName, , , vbTextCompare)
            FontColor = Hex(.txtMessage.ForeColor)
            FontColor = String$(6 - Len(FontColor), "0") & FontColor
            FontColor = Right$(FontColor, 2) & Mid$(FontColor, 3, 2) & Left$(FontColor, 2)
            Alias = Replace$(Alias, "(font)", IIf(.txtMessage.FontBold, "Bold ", vbNullString) & IIf(.txtMessage.FontItalic, "Italic ", vbNullString) & IIf(.txtMessage.FontStrikethru, "Strikethru ", vbNullString) & IIf(.txtMessage.FontUnderline, "Underline ", vbNullString) & .txtMessage.FontName, , , vbTextCompare)
            Alias = Replace$(Alias, "(fontcolor)", FontColor, , , vbTextCompare)
            Alias = Replace$(Alias, "(fontbold)", .txtMessage.FontBold, , , vbTextCompare)
            Alias = Replace$(Alias, "(fontitalic)", .txtMessage.FontItalic, , , vbTextCompare)
            Alias = Replace$(Alias, "(fontstrikethru)", .txtMessage.FontStrikethru, , , vbTextCompare)
            Alias = Replace$(Alias, "(fontunderline)", .txtMessage.FontUnderline, , , vbTextCompare)
            Alias = Replace$(Alias, "(buddycount)", .ChatBuddies.Count, , , vbTextCompare)
            If .ChatBuddies.Count = 0 Then
                Alias = Replace$(Alias, "(buddyemail)", .objMSN_SB.Contact, , , vbTextCompare)
                Alias = Replace$(Alias, "(buddynick)", GetContactAttr(.objMSN_SB.Contact, "nick"), , , vbTextCompare)
                Alias = Replace$(Alias, "(buddycomment)", GetBuddyComment(.objMSN_SB.Contact), , , vbTextCompare)
                Alias = Replace$(Alias, "(buddycustomnick)", GetBuddyCustomNick(.objMSN_SB.Contact), , , vbTextCompare)
            Else
                Alias = Replace$(Alias, "(buddyemail)", .ChatBuddies(1).Item("email"), , , vbTextCompare)
                Alias = Replace$(Alias, "(buddynick)", .ChatBuddies(1).Item("nick"), , , vbTextCompare)
                Alias = Replace$(Alias, "(buddycomment)", GetBuddyComment(.ChatBuddies(1).Item("email")), , , vbTextCompare)
                Alias = Replace$(Alias, "(buddycustomnick)", GetBuddyCustomNick(.ChatBuddies(1).Item("email")), , , vbTextCompare)
                If IsNumeric(SubString(Alias, "(buddy", ")")) Then
                    Dim i As Integer
                    For i = 1 To .ChatBuddies.Count
                        Alias = Replace$(Alias, "(buddy" & i & "email)", .ChatBuddies(i).Item("email"), , , vbTextCompare)
                        Alias = Replace$(Alias, "(buddy" & i & "nick)", .ChatBuddies(i).Item("nick"), , , vbTextCompare)
                        Alias = Replace$(Alias, "(buddy" & i & "comment)", GetBuddyComment(.ChatBuddies(i).Item("email")), , , vbTextCompare)
                        Alias = Replace$(Alias, "(buddy" & i & "customnick)", GetBuddyCustomNick(.ChatBuddies(i).Item("email")), , , vbTextCompare)
                    Next
                End If
            End If
        End With
    End If
End Function

Public Function GetSubMenu(srcMenu As Object, Tag As String) As Integer
    Dim i As Integer
    For i = srcMenu.LBound To srcMenu.UBound
        If srcMenu(i).Tag = Tag Then
            GetSubMenu = i
            Exit Function
        End If
    Next
End Function

Public Sub EnableControl(srcControl As Control)
    If TypeOf srcControl Is TextBox Then
        srcControl.BackColor = vbWindowBackground
        srcControl.Enabled = True
    End If
End Sub

Public Sub DisableControl(srcControl As Control)
    If TypeOf srcControl Is TextBox Then
        srcControl.BackColor = vbButtonFace
    End If
    srcControl.Enabled = False
End Sub

Public Function GetWindowsVersion() As Integer
    GetWindowsVersion = GetVersion Mod 256
End Function

Public Sub SaveUserSettings()
    On Error Resume Next
    
    SaveSettingX "App Settings\" & frmMain.objMSN_NS.Login, "Show MyDP", ShowMyDP
    SaveSettingX "App Settings\" & frmMain.objMSN_NS.Login, "IMWindow Width", IMWindowWidth
    SaveSettingX "App Settings\" & frmMain.objMSN_NS.Login, "IMWindow Height", IMWindowHeight
    SaveSettingX "App Settings\" & frmMain.objMSN_NS.Login, "IMWindow Max", IMWindowMax
    SaveSettingX "App Settings\" & frmMain.objMSN_NS.Login, "Text Style", TextStyle
    SaveSettingX "App Settings\" & frmMain.objMSN_NS.Login, "Emoticon FloodControl", EmoticonFloodControl
    SaveSettingX "App Settings\" & frmMain.objMSN_NS.Login, "IMFont Name", IMFontName
    SaveSettingX "App Settings\" & frmMain.objMSN_NS.Login, "IMFont Size", IMFontSize
    SaveSettingX "App Settings\" & frmMain.objMSN_NS.Login, "IMFont Color", IMFontColor
    SaveSettingX "App Settings\" & frmMain.objMSN_NS.Login, "IMFont Bold", IMFontBold
    SaveSettingX "App Settings\" & frmMain.objMSN_NS.Login, "IMFont Italic", IMFontItalic
    SaveSettingX "App Settings\" & frmMain.objMSN_NS.Login, "IMFont Strikethru", IMFontStrikethru
    SaveSettingX "App Settings\" & frmMain.objMSN_NS.Login, "IMFont Underline", IMFontUnderline
    SaveSettingX "App Settings\" & frmMain.objMSN_NS.Login, "IMFont RandomFormat", IMFontRandomFormat
    SaveSettingX "App Settings\" & frmMain.objMSN_NS.Login, "IMFont RandomColors", IMFontRandomColors
    SaveSettingX "App Settings\" & frmMain.objMSN_NS.Login, "Time Stamp", TimeStamp
End Sub

Public Sub LoadUserSettings()
    On Error Resume Next

    ShowMyDP = GetSettingX("App Settings\" & frmMain.objMSN_NS.Login, "Show MyDP", True)
    SendDisplayPic = GetSettingX("App Settings\" & frmMain.objMSN_NS.Login, "Send DisplayPic", True)
    ReceiveDisplayPic = GetSettingX("App Settings\" & frmMain.objMSN_NS.Login, "Receive DisplayPic", True)
    SaveStatusHistory = GetSettingX("App Settings\" & frmMain.objMSN_NS.Login, "Save StatusHistory", True)
    StatusHistoryFolder = GetSettingX("App Settings\" & frmMain.objMSN_NS.Login, "StatusHistory Folder", App.Path & "\" & "Status History\")
    SaveMessageHistory = GetSettingX("App Settings\" & frmMain.objMSN_NS.Login, "Save MessageHistory", True)
    AutoIdle = GetSettingX("App Settings\" & frmMain.objMSN_NS.Login, "AutoIdle", IIf(GetWindowsVersion < 5, False, True))
    AutoIdle_Interval = Val(GetSettingX("App Settings\" & frmMain.objMSN_NS.Login, "AutoIdle Interval", 5))
    MessageHistoryFolder = GetSettingX("App Settings\" & frmMain.objMSN_NS.Login, "MessageHistory Folder", App.Path & "\Message History\" & frmMain.objMSN_NS.Login & "\")
    ShowIMWindowOnMsg = GetSettingX("App Settings\" & frmMain.objMSN_NS.Login, "Show IMWindow OnMsg", False)
    TypingNotification = Not GetSettingX("App Settings\" & frmMain.objMSN_NS.Login, "Disable MsgTypingNotification", False)
    HighlightFakeFriends = GetSettingX("App Settings\" & frmMain.objMSN_NS.Login, "Highlight FakeFriends", True)
    EmoticonFloodControl = GetSettingX("App Settings\" & frmMain.objMSN_NS.Login, "Emoticon FloodControl", False)
    IMWindowWidth = Val(GetSettingX("App Settings\" & frmMain.objMSN_NS.Login, "IMWindow Width", 5985))
    IMWindowWidth = IIf(IMWindowWidth = 0, 5985, IMWindowWidth)
    IMWindowHeight = Val(GetSettingX("App Settings\" & frmMain.objMSN_NS.Login, "IMWindow Height", 5520))
    IMWindowHeight = IIf(IMWindowHeight = 0, 5520, IMWindowHeight)
    IMWindowMax = GetSettingX("App Settings\" & frmMain.objMSN_NS.Login, "IMWindow Max", False)
    TextStyle = GetSettingX("App Settings\" & frmMain.objMSN_NS.Login, "Text Style")
    ShowEmoticons = GetSettingX("App Settings\" & frmMain.objMSN_NS.Login, "Show Emoticons", True)
    IMFontName = GetSettingX("App Settings\" & frmMain.objMSN_NS.Login, "IMFont Name", "Tahoma")
    IMFontSize = GetSettingX("App Settings\" & frmMain.objMSN_NS.Login, "IMFont Size", 2)
    IMFontColor = Val(GetSettingX("App Settings\" & frmMain.objMSN_NS.Login, "IMFont Color", vbBlack))
    IMFontBold = GetSettingX("App Settings\" & frmMain.objMSN_NS.Login, "IMFont Bold", False)
    IMFontItalic = GetSettingX("App Settings\" & frmMain.objMSN_NS.Login, "IMFont Italic", False)
    IMFontStrikethru = GetSettingX("App Settings\" & frmMain.objMSN_NS.Login, "IMFont Strikethru", False)
    IMFontUnderline = GetSettingX("App Settings\" & frmMain.objMSN_NS.Login, "IMFont Underline", False)
    IMFontRandomFormat = GetSettingX("App Settings\" & frmMain.objMSN_NS.Login, "IMFont RandomFormat", False)
    IMFontRandomColors = GetSettingX("App Settings\" & frmMain.objMSN_NS.Login, "IMFont RandomColors", False)
    TimeStamp = GetSettingX("App Settings\" & frmMain.objMSN_NS.Login, "Time Stamp", True)
    boolUseDefaultEmailApp = GetSettingX("App Settings\" & frmMain.objMSN_NS.Login, "Use DefaultEmailApp", True)
    strCustomEmailApp = GetSettingX("App Settings\" & frmMain.objMSN_NS.Login, "Custom EmailApp")
    strCustomEmailWeb = GetSettingX("App Settings\" & frmMain.objMSN_NS.Login, "Custom EmailWeb")
    boolUseCustomEmailWeb = GetSettingX("App Settings\" & frmMain.objMSN_NS.Login, "Use CustomEmailWeb", False)
End Sub

Public Sub LogChat(Contact As String, Message As String)
    On Error Resume Next
    
    If SaveMessageHistory And Not (Contact = vbNullString Or Message = vbNullString) Then
        Dim FileNum As Integer
        FileNum = FreeFile
        MakeSureDirectoryPathExists MessageHistoryFolder
        Open MessageHistoryFolder & "\" & Contact & ".txt" For Append As #FileNum
            Print #FileNum, Message
        Close FileNum
    End If
End Sub

Public Sub ChangeGMStatus(Status As String, Optional Log As Boolean)
    On Error Resume Next
    
    frmMain.lblStatus.Caption = "[" & Time$ & "] " & Status
    If Log Then
        Call LogStatus(Status)
    End If
End Sub

Public Sub LogStatus(Status As String)
    If SaveStatusHistory And Not Status = vbNullString Then
        Dim FileNum As Integer
        FileNum = FreeFile
        MakeSureDirectoryPathExists StatusHistoryFolder
        Open StatusHistoryFolder & "\" & frmMain.objMSN_NS.Login & ".txt" For Append As #FileNum
            Print #FileNum, "[" & Now & "] " & Status
        Close FileNum
    End If
End Sub

Public Sub ShowBuddyProperties(OwnerForm As Form, Email As String, Optional Nick As String)
    On Error Resume Next
    
    With frmContactProperties
        .txtCustomNick = GetBuddyCustomNick(Email)
        .txtEmail.Text = Email
        If Not InCollection(ContactList, Email) Then
            .txtNick.Locked = True
            .txtNick.Text = IIf(Nick = vbNullString, GetContactAttr(Email, "nick"), Nick)
            .txtStatus.Text = "Unknown"
            .txtGroups.Text = "_"
        Else
            .txtNick.Text = IIf(Nick = vbNullString, GetContactAttr(Email, "nick"), Nick)
            .txtStatus.Text = StatusName(GetContactAttr(Email, "status"))
            Dim i As Integer
            .txtGroups.Text = ContactGroups("GRP " & ContactList(Email).Item("groups").Item(1)).Item("name")
            For i = 2 To ContactList(Email).Item("groups").Count
                .txtGroups.Text = .txtGroups.Text & ", " & ContactGroups("GRP " & ContactList(Email).Item("groups").Item(1)).Item("name")
            Next
        End If
        .txtComment = GetBuddyComment(Email)
        .lblLastOnline.Caption = "Last seen online on " & GetSettingX("Statistics\" & Email, "Last Online", "_")
        .lblLastConversationStarted.Caption = "Last started conversation on " & GetSettingX("Statistics\" & Email, "Last ConversationStarted", "_")
        .lblLastConversationJoined.Caption = "Last joined conversation on " & GetSettingX("Statistics\" & Email, "Last ConversationJoined", "_")
        .lblLastMessageSent.Caption = "Last sent message on " & GetSettingX("Statistics\" & Email, "Last MessageSent", "_")
        .lblLastIP.Caption = "Last IP address: " & GetSettingX("Statistics\" & Email, "Last IP", "_")
        .Tag = Email
        .Show vbModal, OwnerForm
    End With
End Sub

Public Function GetSettingX(Section As String, Key As String, Optional DefaultVal)
    On Error Resume Next
    
    GetSettingX = GetSetting(App.Title, Section, Key)
    If (CStr(GetSettingX)) = vbNullString And Not IsMissing(DefaultVal) Then
        GetSettingX = DefaultVal
    End If
End Function

Public Sub SaveSettingX(Section As String, Key As String, Setting)
    On Error Resume Next
    
    SaveSetting App.Title, Section, Key, Setting
End Sub

Public Function GetBuddyComment(Email) As String
    On Error Resume Next
    
    GetBuddyComment = ContactComments(Email)
End Function

Public Sub SetBuddyComment(Email As String, Comment As String)
    On Error Resume Next
    
    If Not Comment = vbNullString Then
        SetCollectionItem ContactComments, Email, Comment
        SaveSetting "Gilly Messenger", "Comments\" & frmMain.objMSN_NS.Login, Email, Comment
    Else
        ContactComments.Remove Email
        DeleteSetting "Gilly Messenger", "Comments\" & frmMain.objMSN_NS.Login, Email
    End If
End Sub

Public Function GetBuddyCustomNick(Email As String) As String
    On Error Resume Next
    
    GetBuddyCustomNick = ContactCustomNicks(Email)
End Function

Public Sub SetBuddyCustomNick(Email As String, CustomNick As String)
    On Error Resume Next
    
    If Not CustomNick = vbNullString Then
        SetCollectionItem ContactCustomNicks, Email, CustomNick
        SaveSetting "Gilly Messenger", "Custom Nicks\" & frmMain.objMSN_NS.Login, Email, CustomNick
    Else
        ContactCustomNicks.Remove Email
        DeleteSetting "Gilly Messenger", "Custom Nicks\" & frmMain.objMSN_NS.Login, Email
    End If
    
    If InList(GetContactAttr(Email, "lists"), msnList_Forward) Then
        frmMain.RefreshContact (Email)
    End If
End Sub

Public Sub Signout()
    frmMain.lblStatus.Caption = vbNullString
    
    If frmMain.objMSN_NS.State = NsState_SignedIn Then
        Call SaveUserSettings
    End If
    
    frmMain.objMSN_NS.Disconnect
    Call LogStatus("Signed out.")
End Sub

Public Function GetContactState(Email As String)
    GetContactState = "(" & StatusName(GetContactAttr(Email, "status")) & ")"
    If InList(GetContactAttr(Email, "lists"), msnList_Block) Then
        GetContactState = GetContactState & " (Blocked)"
    End If
    If InCollection(IgnoreList, Email) Then
        GetContactState = GetContactState & " (Ignored)"
    End If
End Function

Public Sub SignIn(Login As String, Password As String)
    If Not frmMain.objMSN_NS.State = NsState_Disconnected Then
        frmMain.objMSN_NS.Disconnect
    End If
    frmMain.mnuFile_SignIn.Caption = "Cancel S&ign In"
    frmMain.picSignIn.Visible = False
    frmMain.picSignInProgress.Cls
    frmMain.picSignInProgress.Visible = True
    frmMain.objMSN_NS.Login = Login
    frmMain.objMSN_NS.Password = Password
    frmMain.objMSN_NS.Server = GetSettingX("Server Settings", "IPAddress", "messenger.hotmail.com")
    frmMain.objMSN_NS.Port = Val(GetSettingX("Server Settings", "Port", 1863))
    frmMain.objMSN_NS.Connect
    frmMain.tmrReconnect.Enabled = False
End Sub


Public Sub PlaySound(File As String, ID As String)
    mciSendString "close gm_" & ID, vbNullString, 0, 0
    mciSendString "open " & Chr(34) & File & Chr(34) & " alias gm_" & ID, vbNullString, 0, 0
    mciSendString "play gm_" & ID, vbNullString, 0, 0
End Sub

Public Sub ContactSound(Contact As String, Sound As String, ID As String)
    If PatternSearch(SoundFilter, Contact, SoundFilterMode) Then
        Call PlaySound(Sound, ID)
    End If
End Sub

Public Function ArraySize(srcArray() As String, Optional Dimension As Integer) As Integer
    On Error Resume Next
    
    ArraySize = -1
    If Dimension = 0 Then
        ArraySize = UBound(srcArray)
    Else
        ArraySize = UBound(srcArray, Dimension)
    End If
End Function

Public Function GetString(Key As String) As String
    GetString = TrimX(Key)
    If Right$(GetString, 1) = """" And Left$(GetString, 1) = """" Then
        GetString = Mid$(GetString, 2, Len(GetString) - 2)
    End If
End Function

Public Function PositiveInt(Num As Integer) As Integer
    If Num < 0 Then
        PositiveInt = 0
    Else
        PositiveInt = Num
    End If
End Function

Public Sub Swap(Var1 As String, Var2 As String)
    Dim Temp As String
    Temp = Var2
    Var2 = Var1
    Var1 = Temp
End Sub

Public Function IsEmail(strEmail As String) As Boolean
    Dim intCharPos As Integer
    intCharPos = InStr(strEmail, "@")
    If Not intCharPos = 0 Then
        intCharPos = InStr(intCharPos + 1, strEmail, ".")
        If Not intCharPos = 0 Then
            IsEmail = True
        End If
    End If
End Function

Public Function CharCount(Text As String, Char As String) As Integer
    Dim CharPos As Integer
    Do
        CharPos = InStr(CharPos + 1, Text, Char)
        If CharPos = 0 Then
            Exit Do
        Else
            CharCount = CharCount + 1
        End If
    Loop
End Function

Public Function TrimX(Text As String) As String
    TrimX = Trim$(Text)
    Do Until (Left$(TrimX, 1) <> Chr$(vbKeyTab) And Left$(TrimX, 1) <> Chr$(vbKeySpace)) Or TrimX = vbNullString
        TrimX = Right$(TrimX, Len(TrimX) - 1)
    Loop
End Function

Public Function PatternSearch(srcCollection As Collection, Key As String, Mode As Boolean) As Boolean
    Dim i As Integer
    If Not Mode Then
        For i = 1 To srcCollection.Count
            If Key Like srcCollection(i) Then
                PatternSearch = True
                Exit Function
            End If
        Next
    Else
        PatternSearch = True
        For i = 1 To srcCollection.Count
            If Key Like srcCollection(i) Then
                PatternSearch = False
                Exit Function
            End If
        Next
    End If
End Function

Public Sub VibrateWindow(Window As Form)
    Dim i As Integer
    For i = 1 To 4
        Window.Left = Window.Left - 200
        Sleep 10
        Window.Top = Window.Top - 200
        Sleep 10
        Window.Left = Window.Left + 400
        Sleep 10
        Window.Top = Window.Top + 400
        Sleep 10
        Window.Left = Window.Left - 200
        Sleep 10
        Window.Top = Window.Top - 200
    Next
End Sub

Public Function GetTempDir() As String
    Dim intLen As Integer
    GetTempDir = String(255, Chr$(0))
    intLen = GetTempPath(255, GetTempDir)
    GetTempDir = Left$(GetTempDir, intLen)
End Function

Public Function GetSysDir() As String
    Dim intLen As Integer
    GetSysDir = String(255, Chr$(0))
    intLen = GetSystemDirectory(GetSysDir, 255)
    GetSysDir = Left$(GetSysDir, intLen)
End Function

Public Function GetCurrentSong() As String
    Dim strWinampSong As String, strRealSong As String
    
    strWinampSong = GetWinampSong
    strRealSong = GetRealSong
    
    Dim strCurrentSong As String
    
    If Not strWinampSong = vbNullString Then
        strCurrentSong = strWinampSong
    Else
        If Not strRealSong = vbNullString Then
            strCurrentSong = strRealSong
        Else
            Exit Function
        End If
    End If
    
    Dim i As Integer
    i = 1
    Do While IsNumeric(Mid$(strCurrentSong, i, 1))
        i = i + 1
    Loop
    
    Do Until Not (Mid$(strCurrentSong, i, 1) = " " Or Mid$(strCurrentSong, i, 1) = "-" Or Mid$(strCurrentSong, i, 1) = ".")
        i = i + 1
    Loop
    
    If i < Len(strCurrentSong) Then
        strCurrentSong = Mid$(strCurrentSong, i)
    End If
    
    If strCurrentSong Like "*.???" Then
        strCurrentSong = Left$(strCurrentSong, Len(strCurrentSong) - 1)
    End If
    
    GetCurrentSong = strCurrentSong
End Function

Private Function GetWinampSong() As String
    Dim hWinamp As Long, strSongTitle As String, i As Integer
    
    hWinamp = FindWindow("Winamp v1.x", vbNullString)
    If Not hWinamp = 0 Then
        strSongTitle = GetWindowCaption(hWinamp)
        i = InStr(strSongTitle, " - Winamp")
        If Not i = 0 Then
            If Not Right$(strSongTitle, 9) = "[Stopped]" Then
                strSongTitle = Left$(strSongTitle, i - 1)
                GetWinampSong = strSongTitle
            End If
        End If
    End If
End Function

Private Function GetRealSong() As String
    On Error Resume Next
    
    Dim hRealPlayer As Long, i As Integer, Caption As String
    Do
        hRealPlayer = FindWindowEx(GetDesktopWindow, hRealPlayer, "GeminiWindowClass", vbNullString)
        If hRealPlayer = 0 Then
            Exit Do
        Else
            Caption = GetWindowCaption(hRealPlayer)
            If Left$(Caption, 12) = "RealPlayer: " Then
                GetRealSong = Mid$(Caption, 13)
                Exit Do
            End If
        End If
    Loop
End Function

Public Function GetWindowCaption(hwnd As Long) As String
    Dim intTitleLen As Integer
    intTitleLen = GetWindowTextLength(hwnd)
    GetWindowCaption = String$(intTitleLen, Chr$(0))
    GetWindowText hwnd, GetWindowCaption, intTitleLen + 1
End Function

Public Function GetBuddyInfo(Email As String) As String
    On Error Resume Next
    
    GetBuddyInfo = GetBuddyInfo & "Email: " & Email & vbCrLf
    GetBuddyInfo = GetBuddyInfo & "Nick: " & GetContactAttr(Email, "nick") & vbCrLf
    GetBuddyInfo = GetBuddyInfo & "Custom Nick: " & GetBuddyCustomNick(Email) & vbCrLf
    GetBuddyInfo = GetBuddyInfo & "Status: " & StatusName(GetContactAttr(Email, "status")) & vbCrLf
    Dim Groups As String
    If InCollection(ContactList, Email) Then
        Dim i As Integer
        Groups = Groups & ContactGroups("GRP " & ContactList(Email).Item("groups").Item(1)).Item("name")
        For i = 2 To ContactList(Email).Item("groups").Count
            Groups = Groups & ", " & ContactGroups("GRP " & ContactList(Email).Item("groups").Item(1)).Item("name")
        Next
    End If
    GetBuddyInfo = GetBuddyInfo & "Groups: " & Groups & vbCrLf
    GetBuddyInfo = GetBuddyInfo & "Comment: " & GetBuddyComment(Email) & vbCrLf
    GetBuddyInfo = GetBuddyInfo & "Last seen online on " & GetSettingX("Statistics\" & Email, "Last Online", "-") & vbCrLf
    GetBuddyInfo = GetBuddyInfo & "Last started conversation on " & GetSettingX("Statistics\" & Email, "Last ConversationStarted", "-") & vbCrLf
    GetBuddyInfo = GetBuddyInfo & "Last joined conversation on " & GetSettingX("Statistics\" & Email, "Last ConversationJoined", "-") & vbCrLf
    GetBuddyInfo = GetBuddyInfo & "Last sent message on " & GetSettingX("Statistics\" & Email, "Last MessageSent", "-") & vbCrLf
    GetBuddyInfo = GetBuddyInfo & "Last IP address: " & GetSettingX("Statistics\" & Email, "Last IP", "-")
    
    GetBuddyInfo = Replace$(GetBuddyInfo, ": " & vbCrLf, ": -" & vbCrLf)
End Function

Public Sub LoadDataIntoFile(DataName As String, FileName As String)
    Dim myArray() As Byte
    Dim myFile As Long
    If Dir(FileName) = "" Then
        myArray = LoadResData(DataName, "CUSTOM")
        myFile = FreeFile
        Open FileName For Binary Access Write As #myFile
        Put #myFile, , myArray
        Close #myFile
    End If
End Sub

Public Sub LoadDP(Email As String, picDP As Control)
    On Error Resume Next
    Dim DpId As String, DpPath As String
    DpId = GetSettingX("Display Pics", Email)
    DpPath = App.Path & "\Display Pics\" & Email & ".dat"
    If Not DpId = vbNullString And FileExists(DpPath) Then
        If ShowMyDP Or Not picDP.Name = "imgMyDP" Then
            Set picDP.Picture = LoadPicture(DpPath)
        End If
        picDP.Visible = True
    End If
End Sub

Public Function WordCount(SourceString, SearchString) As Integer
    If Not (SourceString = vbNullString Or SearchString = vbNullString) Then
        Dim i As Integer
        For i = 1 To Len(SourceString)
            If StrComp(Mid$(SourceString, i, Len(SearchString)), SearchString, vbTextCompare) = 0 Then
                WordCount = WordCount + 1
            End If
        Next
    End If
End Function

Public Sub SetNumbered(hTextBox As Long, SetState As Boolean)
    Dim lngStyle As Long
    If SetState Then
        lngStyle = GetWindowLong(hTextBox, GWL_STYLE)
        SetWindowLong hTextBox, GWL_STYLE, lngStyle Or ES_NUMBER
    Else
        lngStyle = GetWindowLong(hTextBox, GWL_STYLE)
        SetWindowLong hTextBox, GWL_STYLE, lngStyle And Not ES_NUMBER
    End If
End Sub

Public Function GetCustomNick(Email As String, Nick As String) As String
    GetCustomNick = GetBuddyCustomNick(Email)
    If GetCustomNick = vbNullString Then
        GetCustomNick = Nick
    End If
End Function

Public Function FileExists(Path As String) As Boolean
    FileExists = CBool(PathFileExists(Path))
End Function

Public Function GetIdleTime() As Double
    On Error GoTo Handler:
    If GetVersion < 5 Then
        GetIdleTime = Fix(Timer - LastActive)
    Else
        Dim InputInfo As LASTINPUTINFO
        InputInfo.cbSize = Len(InputInfo)
        GetLastInputInfo InputInfo
        GetIdleTime = (GetTickCount - InputInfo.dwTime) \ 1000
    End If
    Exit Function
Handler:
    On Error Resume Next
    GetIdleTime = Fix(Timer - LastActive)
End Function

Public Sub OpenMailBox()
    If boolUseDefaultEmailApp Then
        If frmMain.objMSN_NS.EmailEnabled Then
            frmMain.objMSN_NS.RequestURL "INBOX"
        Else
            ShellExecute 0, "open", strDefaultEmailApp, vbNullString, vbNullString, 1
        End If
    Else
        If boolUseCustomEmailWeb Then
            Call WebNavigate(strCustomEmailWeb)
        Else
            ShellExecute 0, "open", strCustomEmailApp, vbNullString, vbNullString, 1
        End If
    End If
End Sub

Public Sub WebNavigate(URL As String)
    If (boolUseDefaultBrowser And strDefaultBrowser = vbNullString) Or (boolUseDefaultBrowser = False And strCustomBrowser = vbNullString) Then
        ShellExecute 0, "open", URL, vbNullString, vbNullString, 1
    Else
        If boolUseDefaultBrowser Then
            ShellExecute 0, "open", strDefaultBrowser, URL, vbNullString, 1
        Else
            ShellExecute 0, "open", strCustomBrowser, URL, vbNullString, 1
        End If
    End If
End Sub

Public Sub SendEmail(Target As String)
    If boolUseDefaultEmailApp Then
        If frmMain.objMSN_NS.EmailEnabled Then
            frmMain.objMSN_NS.RequestURL "COMPOSE " & Target
        Else
            ShellExecute 0, "open", "mailto:" & Target, vbNullString, vbNullString, 1
        End If
    Else
        If boolUseCustomEmailWeb Then
            Call WebNavigate(strCustomEmailWeb)
        Else
            ShellExecute 0, "open", strCustomEmailApp, Target, vbNullString, 1
        End If
    End If
End Sub

Public Function ConvArray(ParamArray Elements())
    Dim tmpArray()
    ReDim tmpArray(UBound(Elements))
    Dim i As Integer
    For i = 0 To UBound(Elements)
        tmpArray(i) = Elements(i)
    Next
    ConvArray = tmpArray()
End Function

Public Function FullScrApp() As Boolean
    Dim hActiveWnd As Long
    hActiveWnd = GetForegroundWindow
    Dim WndRect As RECT, DesktopRect As RECT
    If Not hActiveWnd = GetDesktopWindow Then
        GetWindowRect hActiveWnd, WndRect
        GetWindowRect GetDesktopWindow, DesktopRect
        If ((WndRect.Right >= DesktopRect.Right And WndRect.Left = 0 And WndRect.Bottom >= DesktopRect.Bottom And WndRect.Top = 0) Or (WndRect.Right <= 0 And WndRect.Bottom <= 0)) And Not GetWindowCaption(hActiveWnd) = "Program Manager" Then
            If BlockAlertsOnFullScrApp Then
                FullScrApp = True
            End If
        End If
    End If
End Function

Public Function GetSpecialFolder(CSIDL As Long) As String
    Dim IDL As ITEMIDLIST
    SHGetSpecialFolderLocation 100, CSIDL, IDL
    GetSpecialFolder = String$(512, Chr$(0))
    SHGetPathFromIDList ByVal IDL.mkid.cb, ByVal GetSpecialFolder
    GetSpecialFolder = Left$(GetSpecialFolder, InStr(GetSpecialFolder, Chr$(0)) - 1)
End Function

Public Function SubArray(SourceArray() As String, IndexStart As Integer, IndexStop As Integer) As String()
    Dim TempArray() As String
    ReDim TempArray(IndexStop - IndexStart)
    Dim i As Integer
    For i = 0 To UBound(TempArray)
        TempArray(i) = SourceArray(IndexStart + i)
    Next
    SubArray = TempArray
End Function

Public Function ShellExecuteEx(File As String) As Boolean
    File = TrimX(File)
    Dim intResult As Integer
    If FileExists(File) Then
        intResult = ShellExecute(0, "open", File, vbNullString, vbNullString, 1)
    Else
        Dim i As Integer, FileParams() As String
        FileParams = Split(File)
        For i = UBound(FileParams) - 1 To 0 Step -1
            If FileExists(Join(SubArray(FileParams, 0, i))) Then
                intResult = ShellExecute(0, "open", Join(SubArray(FileParams, 0, i)), Join(SubArray(FileParams, i + 1, UBound(FileParams))), vbNullString, 1)
                Exit For
            End If
        Next
    End If
    If intResult <= 32 Then
        intResult = ShellExecute(0, "open", File, vbNullString, vbNullString, 1)
    End If
    ShellExecuteEx = (intResult > 32)
End Function

Public Function SubString(strSource As String, strStart As String, strStop As String)
    Dim i As Integer, KeyCount As Integer
    Dim intStart As Integer, intStop As Integer
    Dim Key_Start_Len As Integer, Key_Stop_Len As Integer
    Key_Start_Len = Len(strStart)
    Key_Stop_Len = Len(strStop)
    If strStart = vbNullString Then
        intStop = InStr(strSource, strStop)
        If Not intStop = 0 Then
            SubString = Left$(strSource, intStop - 1)
        End If
    ElseIf strStop = vbNullString Then
        intStart = InStr(strSource, strStart)
        If Not intStart = 0 Then
            SubString = Right$(strSource, Len(strSource) - intStart - Key_Start_Len + 1)
        End If
    Else
        For i = 1 To Len(strSource)
            If Mid$(strSource, i, Key_Start_Len) = strStart Then
                If KeyCount = 0 Then
                    intStart = i + Key_Start_Len
                End If
                KeyCount = KeyCount + 1
                i = i + Key_Start_Len - 1 'subtracting 1 to compensate the FOR loop's increment
            ElseIf Mid$(strSource, i, Key_Stop_Len) = strStop And KeyCount > 0 Then
                KeyCount = KeyCount - 1
                If KeyCount = 0 Then
                    SubString = Mid$(strSource, intStart, i - intStart)
                    Exit Function
                End If
            End If
        Next
        If KeyCount > 0 Then
            intStop = InStrRev(strSource, strStop)
            If intStop > intStart Then
                SubString = Mid$(strSource, intStart, intStop - intStart)
            End If
        End If
    End If
End Function

Public Sub SetRandomPos(Owner As Form, Child As Form)
    On Error Resume Next
    
    Randomize Timer
    Child.Left = Owner.Left + (Owner.Width / 3) + Fix(Rnd * (Owner.Width / 3)) - (Child.Width / 2)
    Child.Top = Owner.Top + (Owner.Height / 3) + Fix(Rnd * (Owner.Height / 3)) - (Child.Height / 2)
End Sub

Public Sub SaveFocus()
    hForegroundWnd = GetForegroundWindow()
    hFocusWnd = GetFocusWindow()
    hActiveWnd = GetActiveWindow()
End Sub

Public Sub RestoreFocus()
    SetForegroundWindow hForegroundWnd
    SetFocusWindow hFocusWnd
    SetActiveWindow hActiveWnd
End Sub

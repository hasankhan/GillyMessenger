Attribute VB_Name = "modChatCommands"
Public Function ProcessChatCommand(ChatWindow As Form, ChatCommand As String) As Boolean
On Error Resume Next
If UCase$(ChatCommand) = "CLS" Then
    ChatWindow.txtChat.Text = vbNullString
    ProcessChatCommand = True
ElseIf UCase$(ChatCommand) = "/RECONNECT" Then
    If ChatWindow.ChatBuddies.Count = 0 Then
        Invite ChatWindow.lblStatus.Tag, ChatWindow.cTrialID, ChatWindow.wskChat
        ChatWindow.lblStatus.Caption = "Reconnecting..."
        Call LogChat("Reconnecting...", ChatWindow.lblStatus.Tag)
    End If
    ProcessChatCommand = True
ElseIf UCase$(ChatCommand) = "/VANISH" Then
    ChatWindow.Visible = False
    ProcessChatCommand = True
ElseIf UCase$(Left$(ChatCommand, 6)) = "/MIMIC" Then
    ChatCommand = Trim$(ChatCommand)
    If Len(ChatCommand) > 6 Then
        ChatWindow.Mimic = Right$(ChatCommand, Len(ChatCommand) - 7)
        ChatWindow.Caption = GetBuddyNick(ChatWindow.Mimic)
        ChatWindow.lblStatus.Caption = Replace$(ChatWindow.lblStatus.Caption, ChatWindow.lblStatus.Tag, ChatWindow.Mimic)
    Else
        ChatWindow.lblStatus.Caption = Replace$(ChatWindow.lblStatus.Caption, ChatWindow.Mimic, ChatWindow.lblStatus.Tag)
        ChatWindow.Caption = GetBuddyNick(ChatWindow.lblStatus.Tag)
        ChatWindow.Mimic = vbNullString
    End If
    Call UpdateBuddies(ChatWindow)
    ProcessChatCommand = True
ElseIf UCase$(ChatCommand) = "/VIEW LOG" Then
    ShellExecute 0, "open", ChatLogDir & "\" & Login & "\" & ChatWindow.lblStatus.Tag & ".txt", vbNullString, vbNullString, 1
    ProcessChatCommand = True
ElseIf UCase$(ChatCommand) = "/IGNORE" Then
    Call ChatWindow.mnuIgnore_Click
    ProcessChatCommand = True
ElseIf UCase$(ChatCommand) = "/PROFILE" Then
    Call ChatWindow.mnuProfile_Click
    ProcessChatCommand = True
ElseIf UCase$(ChatCommand) = "/MASK" Then
    Call ChatWindow.mnuMask_Click
    ProcessChatCommand = True
ElseIf UCase$(Left$(ChatCommand, 7)) = "/INVITE" Then
    ChatWindow.LastCall = Split(ChatCommand)(1)
    Invite ChatWindow.LastCall, ChatWindow.cTrialID, ChatWindow.wskChat
    ProcessChatCommand = True
ElseIf UCase$(ChatCommand) = "/BLOCK" Then
    Call ChatWindow.mnuBlock_Click
    ProcessChatCommand = True
ElseIf UCase$(ChatCommand) = "/CHATLOGGER" Then
    Call frmMain.mnuChatLogger_Click
    ProcessChatCommand = True
ElseIf UCase$(ChatCommand) = "/LIST" Then
    Comment ChatWindow, vbNullString
    If ChatWindow.ChatBuddies.Count > 0 Then
        For X = 1 To ChatWindow.ChatBuddies.Count
            If ChatWindow.ChatBuddies(X) = ChatWindow.lblStatus.Tag And ChatWindow.Mimic <> vbNullString Then
                Comment ChatWindow, ChatWindow.Mimic & " - " & GetBuddyNick(ChatWindow.Mimic)
            Else
                Comment ChatWindow, ChatWindow.ChatBuddies(X) & " - " & ChatWindow.ChatBuddyNick(ChatWindow.ChatBuddies(X))
            End If
        Next
    Else
        Comment ChatWindow, "No one is in the conversation."
    End If
    Comment ChatWindow, vbNullString
    ProcessChatCommand = True
ElseIf UCase$(ChatCommand) = "/LIST ONLINE" Then
    Comment ChatWindow, vbNullString
    Comment ChatWindow, ListOnline
    Comment ChatWindow, vbNullString
    ProcessChatCommand = True
ElseIf UCase$(Left$(ChatCommand, 6)) = "/FIND " Then
    ChatCommand = Right$(ChatCommand, Len(ChatCommand) - 6)
    ChatWindow.txtChat.Find ChatCommand, 1
    ProcessChatCommand = True
ElseIf UCase$(Left$(ChatCommand, 6)) = "/NICK " Then
    ChatCommand = Right$(ChatCommand, Len(ChatCommand) - 6)
    ChangeNick ChatCommand
    ProcessChatCommand = True
ElseIf UCase$(Left$(ChatCommand, 7)) = "/COLOR " Then
    ChatCommand = Right$(ChatCommand, Len(ChatCommand) - 7)
    ChatWindow.cdHeader.Color = ColorConv(ChatCommand, True)
    ChatWindow.txtMessage.ForeColor = ChatWindow.cdHeader.Color
    ChatWindow.txtMask.ForeColor = ChatWindow.txtMessage.ForeColor
    Call ChatWindow.CreateHeader
    ChatWindow.ChatColor = ChatWindow.cdHeader.Color
    ProcessChatCommand = True
ElseIf UCase$(ChatCommand) = "/BOLD" Then
    ChatWindow.cdHeader.FontBold = Not ChatWindow.cdHeader.FontBold
    ChatWindow.ChatFontBold = ChatWindow.cdHeader.FontBold
    ChatWindow.txtMessage.FontBold = ChatWindow.cdHeader.FontBold
    ChatWindow.txtMask.FontBold = ChatWindow.txtMessage.FontBold
    Call ChatWindow.CreateHeader
    ProcessChatCommand = True
ElseIf UCase$(ChatCommand) = "/ITALIC" Then
    ChatWindow.cdHeader.FontItalic = Not ChatWindow.cdHeader.FontItalic
    ChatWindow.ChatFontItalic = ChatWindow.cdHeader.FontItalic
    ChatWindow.txtMessage.FontItalic = ChatWindow.cdHeader.FontItalic
    ChatWindow.txtMask.FontItalic = ChatWindow.txtMessage.FontItalic
    Call ChatWindow.CreateHeader
    ProcessChatCommand = True
ElseIf UCase$(Left$(ChatCommand, 8)) = "/AUTOMSG" Then
    X = InStr(ChatCommand, Chr$(vbKeySpace))
    If X > 0 Then
        SetAutoMsg Right$(ChatCommand, Len(ChatCommand) - X)
    ElseIf frmMain.mnuAutoMessage.Checked = True Then
        Call frmMain.mnuAutoMessage_Click
    End If
    ProcessChatCommand = True
ElseIf UCase$(Left$(ChatCommand, 6)) = "/AWAY " Then
    ChatCommand = Right$(ChatCommand, Len(ChatCommand) - 6)
    SetAutoMsg ChatCommand
    ChangeStatus 3
    ProcessChatCommand = True
ElseIf UCase$(ChatCommand) = "/ONLINE" Then
    ChangeStatus 0
    ProcessChatCommand = True
ElseIf UCase$(ChatCommand) = "/BUSY" Then
    ChangeStatus 1
    ProcessChatCommand = True
ElseIf UCase$(ChatCommand) = "/BRB" Then
    ChangeStatus 2
    ProcessChatCommand = True
ElseIf UCase$(ChatCommand) = "/AWAY" Then
    ChangeStatus 3
    ProcessChatCommand = True
ElseIf UCase$(ChatCommand) = "/PHONE" Then
    ChangeStatus 4
    ProcessChatCommand = True
ElseIf UCase$(ChatCommand) = "/LUNCH" Then
    ChangeStatus 5
    ProcessChatCommand = True
ElseIf UCase$(ChatCommand) = "/IDLE" Then
    ChangeStatus 6
    ProcessChatCommand = True
ElseIf UCase$(ChatCommand) = "/HIDE" Then
    ChangeStatus 7
    ProcessChatCommand = True
ElseIf UCase$(Left$(ChatCommand, 6)) = "/CHAT " Then
    ChatCommand = Right$(ChatCommand, Len(ChatCommand) - 6)
    StartChat ChatCommand
    ProcessChatCommand = True
ElseIf UCase$(Left$(ChatCommand, 5)) = "/MSG " Then
    ChatCommand = Right$(ChatCommand, Len(ChatCommand) - 5)
    X = InStr(ChatCommand, " ")
    MessageUser Left$(ChatCommand, X - 1), Right$(ChatCommand, Len(ChatCommand) - X)
    ProcessChatCommand = True
ElseIf UCase$(Left$(ChatCommand, 8)) = "/MSGALL " Then
    ChatCommand = Right$(ChatCommand, Len(ChatCommand) - 8)
    MessageAll ChatCommand
    ProcessChatCommand = True
ElseIf UCase$(Left$(ChatCommand, 8)) = "/SIGNIN " Then
    SignIn Split(ChatCommand, " ")(1), Split(ChatCommand, " ")(2)
    ProcessChatCommand = True
ElseIf UCase$(ChatCommand) = "/SIGNOUT" Then
    Call Signout
    ProcessChatCommand = True
ElseIf UCase$(ChatCommand) = "/MSGR" Then
    ShowMe frmMain
    ProcessChatCommand = True
ElseIf UCase$(ChatCommand) = "/EXIT" Then
    Call frmMain.mnuExit_Click
    ProcessChatCommand = True
End If
End Function

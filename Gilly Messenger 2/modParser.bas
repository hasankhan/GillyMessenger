Attribute VB_Name = "modParser"
Dim bContact As String, Lst As String
Dim cfSet As Boolean

Public Sub Parse(Data As String)
On Error Resume Next
Dim Command As String
'Parse Commands
Do
    'Extract a command
    Command = Left$(Data, InStr(Data, vbCrLf) - 1)
    Debug.Print "<-- " & Command
    'If protocol agreed
    'VER 1 MSNP8
    If Left$(Command, 3) = "VER" And Right$(Command, 5) = "MSNP8" Then
        'Get incryption information
        MsnSend "CVR " & TrialID & " 0x0413 winnt 5.2 i386 MSNMSGR 6.0.0602 MSMSGS " & Login, TrialID, frmMain.wskMSN
        
    'CVR 2 6.0.0602 6.0.0602 6.0.0268 http://download.microsoft.com/download/4/d/e/4defe3d0-b2e3-4de2-ab23-8bd74be402ea/SetupDl.exe http://messenger.msn.com/nl
    ElseIf Left$(Command, 3) = "CVR" Then
        MsnSend "USR " & TrialID & " TWN I " & Login, TrialID, frmMain.wskMSN
        
    'USR 3 TWN S lc=1033,id=507,tw=40,fs=1,ru=http%3A%2F%2Fmessenger%2Emsn%2Ecom,ct=1066301339,kpp=1,kv=5,ver=2.1.0173.1,tpf=5cdba66478a2d79f7f635f26371ddef2
    ElseIf Left$(Command, 3) = "USR" And InStr(Command, "TWN") > 0 Then
        'Send username
        frmMain.lblStatus.Caption = "Authenticating..."
        ' Get the hash supplied by the DS:
        Dim h As Integer, strHashParams As String, strResponse As String
        h = InStr(LCase$(Command), " lc")
        strHashParams = Right$(Command, Len(Command) - h)
            
        ' Start the SSL-procedure:
        strResponse = frmMain.DoSSL(strHashParams)
        If strResponse = "False" Then
            If RcLoggedIn = False Then
                ShowMe frmMain
                MsgBox "Invalid username or password.", vbInformation
                Call frmMain.picSignIn_Click
            End If
            Exit Sub
        ElseIf strResponse = "Error" Then
            Exit Sub
        End If
        ' Pass authentication result back to the DS:
        MsnSend "USR " & TrialID & " TWN S " & strResponse, TrialID, frmMain.wskMSN
        
    'If asks to change the server
    ElseIf Left$(Command, 3) = "XFR" And InStr(Command, "NS") > 0 Then
        'Connect to new server
        frmMain.wskMSN.Close
        TrialID = 1
        frmMain.wskMSN.RemoteHost = Mid$(Command, InStr(Command, "N") + 3, (InStr(Command, ":") - (InStr(Command, "NS") + 3)))
        frmMain.wskMSN.RemotePort = Mid$(Command, InStr(Command, ":") + 1, 4)
        frmMain.wskMSN.Connect
        
    'If password accepted
    'USR 4 OK hasankhan01@hotmail.com %43%72%61%63%6B%53%6F%66%74 1
    ElseIf Left$(Command, 3) = "USR" And InStr(Command, "OK") > 0 Then
        LoginTime = Now()
        If LoginCached(frmSignIn.cmbLogin.Text) = False Then
            cmbLogin.AddItem cmbLogin.Text
            SaveSetting "Gilly Messenger", "Login Cache", Login, vbNullString
        End If
        frmSettings.cmdClearIgnoreList.Enabled = True
        frmSettings.cmdClearContactComments.Enabled = True
        SignedIn = True
        frmMain.tmrPing.Enabled = True
        Status = InitialStatus
        Call UpdateStatusImage
        frmMain.mnuStatusList(InitialStatus).Checked = True
        frmMain.picMask.Visible = False
        'Change tray icon tip
        ChangeTip "Gilly Messenger - " & Login
        'Request users's contact list
        frmMain.lblStatus.Caption = "Retrieving contact list..."
        If SignInMode <> "Online" Then
            MsnSend "SYN " & TrialID & " 0", TrialID, frmMain.wskMSN
        End If
        'Change user status to initial status
        MsnSend "CHG " & TrialID & " " & StatusCode(InitialStatus), TrialID, frmMain.wskMSN
        'Extract nick
        Nick = Split(Command, " ")(4)
        Nick = DeMorph(Nick, True)
        If Status <> 0 Then
            frmMain.lblNick.Caption = Nick & " (" & StatusConv(StatusCode(InitialStatus)) & ")"
        Else
            frmMain.lblNick.Caption = Nick
        End If
        frmMain.txtNick.Text = Nick
        'Update messenger window
        'Get cached ignore list
        Dim IgnoreList() As String
        IgnoreList = GetAllSettings("Gilly Messenger", "Ignore List\" & Login)
        Temp = Str(UBound(IgnoreList))
        For X = 0 To Val(Temp)
            BuddyIgnore.Add IgnoreList(X, 0), IgnoreList(X, 0)
            DoEvents
        Next X
        'Get cached comments
        Dim Comments() As String
        Comments = GetAllSettings("Gilly Messenger", "Comments\" & Login)
        Temp = Str(UBound(IgnoreList))
        For X = 0 To Val(Temp)
            BuddyComment.Add Comments(X, 1), Comments(X, 0)
        Next
        
        ViewContactsByEmail = GetSetting("Gilly Messenger", "App Settings\" & Login, "View Contacts By Email", False)
        If ViewContactsByEmail = True Then
            frmMain.mnuViewContactsByEmail.Checked = True
            frmMain.mnuViewContactsByDisplayName.Checked = False
        Else
            frmMain.mnuViewContactsByDisplayName.Checked = True
            frmMain.mnuViewContactsByEmail.Checked = False
        End If
        
        frmMain.tvwBuddies.Nodes.Add , , "Online", "Online"
        frmMain.tvwBuddies.Nodes("Online").Expanded = True
        frmMain.tvwBuddies.Nodes("Online").Bold = True
        frmMain.tvwBuddies.Nodes("Online").ForeColor = vbBlue
        frmMain.tvwBuddies.Nodes("Online").Sorted = True
        frmMain.tvwBuddies.Nodes("Online").Image = 5
        
        frmMain.tvwBuddies.Nodes.Add , , "Offline", "Not Online"
        frmMain.tvwBuddies.Nodes("Offline").Expanded = True
        frmMain.tvwBuddies.Nodes("Offline").Bold = True
        frmMain.tvwBuddies.Nodes("Offline").ForeColor = vbBlue
        frmMain.tvwBuddies.Nodes("Offline").Sorted = True
        frmMain.tvwBuddies.Nodes("Offline").Image = 5
        
        frmMain.mnuStatus.Enabled = True
        frmMain.mnuSignIn.Caption = "Sign &Out"
        frmMain.mnuSendMessage.Enabled = True
        frmMain.mnuOpenInbox.Enabled = True
        frmMain.mnuChatRooms.Enabled = True
        frmMain.mnuEditPassport.Enabled = True
        frmMain.mnuEditProfile.Enabled = True
        frmMain.mnuAddContact.Enabled = True
        frmMain.mnuViewContactsBy.Enabled = True
        frmMain.mnuSaveContactList.Enabled = True
        frmMain.mnuImportContactList.Enabled = True
        frmMain.mnuSearch.Enabled = True
        frmMain.mnuMessageAll.Enabled = True
        frmMain.mnuIgnoreAll.Enabled = True
        frmMain.mnuGMScript.Enabled = True
        frmMain.lblEmail.Visible = True
        'Save server information
        SaveSetting "Gilly Messenger", "Server Settings", "IP Address", frmMain.wskMSN.RemoteHost
        SaveSetting "Gilly Messenger", "Server Settings", "Port", frmMain.wskMSN.RemotePort
        
    'If error occurs
    ElseIf IsNumeric(Left$(Command, 3)) = True Then
        If RcLoggedIn = False And LastError <> Left$(Command, 3) Then
            LastError = Left$(Command, 3)
            MsgBox MsnError(Left$(Command, 3)), vbInformation
        End If
        If Left$(Command, 3) = "911" Then
            Call ResetSockets
            Call frmMain.picSignIn_Click
        End If
        
    'If challange key requested
    ElseIf Left$(Command, 3) = "CHL" Then
        'Encrypt and send challenge key
        Command = Right$(Command, Len(Command) - InStr(Command, " "))
        Command = Right$(Command, Len(Command) - InStr(Command, " "))
        frmMain.wskMSN.SendData "QRY " & TrialID & " msmsgs@msnmsgr.com 32" & vbCrLf & MD5Encrypt(Command & "Q1P7W2E4J9R8U3S5")
        TrialID = TrialID + 1
        
    'If user signs in from another location
    'OUT OTH
    ElseIf Command = "OUT OTH" Then
        If RcLoggedIn = False Then
            LogStatus "You have been signed out because you signed in from another location."
            Call ResetSockets
            MsgBox "You have been signed out because you signed in from another location.", vbInformation
        End If
        
    'If server is going down
    'OUT SSD
    ElseIf Command = "OUT SSD" Then
        If RcLoggedIn = False Then
            MsgBox "Server is going down for maintenance.", vbInformation
        End If
        
    'If initial status of buddy received
    'ILN 7 NLN hasankhan85@msn.com Someone
    ElseIf Left$(Command, 3) = "ILN" Then
        Contact = BuddyConv(Command)
        SetBuddyProperty Contact.Email, "status", Contact.Status
        SetBuddyProperty Contact.Email, "nick", Contact.Nick
        SetBuddyProperty Contact.Email, "forward", True
        Call UpdateList(Contact.Email, True)
        
    'If contact changes state or nick
    ElseIf Left$(Command, 3) = "NLN" Then
        Contact = BuddyConv(Command)
        If GetBuddyStatus(Contact.Email) = "Offline" Then
            LogStatus Contact.Email & " has signed in."
            frmMain.lblStatus.Caption = Contact.Email & " has signed in."
            If Popups = True Then
                If ViewContactsByEmail = True Then
                    ShowPopup PopupBreak(Contact.Email, True) & vbCrLf & "has signed in.", Contact.Email
                Else
                    ShowPopup PopupBreak(Contact.Nick, True) & vbCrLf & "has signed in.", Contact.Email
                End If
            End If
            frmMain.tmrAnimator.Enabled = True
        End If
        If GetBuddyNick(Contact.Email) <> Contact.Nick Then
            SetBuddyProperty Contact.Email, "nick", Contact.Nick
            UpdateChatCaption Contact.Email, Contact.Nick
        End If
        If GetBuddyStatus(Contact.Email) <> Contact.Status Then
            SetBuddyProperty Contact.Email, "status", Contact.Status
        End If
        Call UpdateList(Contact.Email, True)
        
    'If contact goes offline
    ElseIf Left$(Command, 3) = "FLN" Then
        Command = Right$(Command, Len(Command) - InStr(Command, " "))
        LogStatus Command & " appears to be offline."
        frmMain.lblStatus.Caption = Command & " appears to be offline."
        SetBuddyProperty Command, "status", "Offline"
        Call UpdateList(Command, True)
        If Status <> 7 And LastBlockAlert <> Command Then
            X = 0
            X = OpenChats(Command)
            If X > 0 Then
                Temp = vbNullString
                Temp = FindChat(, CLng(X)).ChatBuddies(Command)
                If Temp = vbNullString Then
                    StartChat Command, , True
                End If
            Else
                StartChat Command, , True
            End If
        End If
        
    'If list user received
    'LST kal_saleem@hotmail.com Kanwal 11 0
    ElseIf Left$(Command, 3) = "LST" Then
        eml = Split(Command, " ")(1)
        Lst = Split(Command, " ")(3)
        
        Set LstBuddy = New Collection
        LstBuddy.Add eml, "email"
        LstBuddy.Add DeMorph(Split(Command)(2), True), "nick"
        LstBuddy.Add "Offline", "status"
        LstBuddy.Add CBool(Lst And Lst_FL), "forward"
        LstBuddy.Add CBool(Lst And Lst_AL), "allow"
        LstBuddy.Add CBool(Lst And Lst_BL), "block"
        LstBuddy.Add CBool(Lst And Lst_RL), "reverse"
        ContactList.Add LstBuddy, eml
        If Lst = 8 Then
            ShowAddContactForm CStr(eml), GetBuddyNick(eml)
        End If
        If LstBuddy("forward") = True Then
            Call UpdateList(CStr(eml), True)
        End If
        
    'If switch bord server address received
    'XFR 9 SB 64.4.12.161:1863 CKI 17154350.1033580721.7348
    ElseIf Left$(Command, 3) = "XFR" And InStr(Command, "SB") > 0 Then
        Set CallForm = CallForms.Item(1)
        CallForm.wskChat.Close
        Temp = Split(Command, " ")(3)
        CallForm.wskChat.RemoteHost = Split(Temp, ":")(0)
        CallForm.wskChat.RemotePort = Split(Temp, ":")(1)
        CallForm.wskChat.Tag = Split(Command, " ")(5)
        CallForm.wskChat.Connect
        CallForms.Remove 1
    'If contact starts conversation
    'RNG 389850 64.4.12.158:1863 CKI 1036849607.27207 hasankhan01@hotmail.com %48%61%73%61%6E
    ElseIf Left$(Command, 3) = "RNG" Then
        bContact = Split(Command, " ")(5)
        If GetBuddyStatus(bContact) = "Offline" And IsInList(bContact) = True And AppIgnored(bContact) = False And RcLoggedIn = False Then
            If LastBlockAlert <> bContact Then
                LastBlockAlert = bContact
                MsgBox bContact & " has blocked you.", vbInformation, "Block Alert!"
            End If
        ElseIf IsIgnored(bContact) = False And (IsInList(bContact) = True Or AppIgnored(bContact) = False) Then
            X = 0
            X = OpenChats(bContact)
            cfSet = False
            If X > 0 Then
                Set TempForm = FindChat(, CLng(X))
                X = 0
                X = TempForm.ChatBuddies.Count
                If X = 0 Then
                    Set RingForm = TempForm
                    cfSet = True
                ElseIf X = 1 Then
                    Temp = vbNullString
                    Temp = TempForm.ChatBuddies(bContact)
                    If Temp = bContact Then
                        Set RingForm = TempForm
                        cfSet = True
                    End If
                End If
            End If
            If cfSet = False Then
                Set RingForm = New frmChat
                OpenChats.Remove bContact
                OpenChats.Add RingForm.hwnd, bContact
            End If
            RingForm.wskChat.Close
            RingForm.Tag = "RING" & Split(Command, " ")(1)
            Temp = Split(Command, " ")(2)
            RingForm.wskChat.RemoteHost = Split(Temp, ":")(0)
            RingForm.wskChat.RemotePort = Split(Temp, ":")(1)
            RingForm.wskChat.Tag = Split(Command, " ")(4)
            RingForm.lblStatus.Tag = bContact
            Call LogChat(bContact & " has opened your window.", bContact)
            If RingForm.Mimic = vbNullString Then
                RingForm.lblStatus.Caption = bContact & " has opened your window."
                RingForm.Caption = DeMorph(Split(Command, " ")(6), True)
                RingForm.lblBuddy.Caption = bContact
            Else
                RingForm.lblStatus.Caption = RingForm.Mimic & " has opened your window."
                RingForm.Caption = GetBuddyNick(RingForm.Mimic)
                RingForm.lblBuddy.Caption = RingForm.Mimic
            End If
            RingForm.wskChat.Connect
            If ShowIMWindowOnMsg = False Then
                Handle = GetForegroundWindow()
                FocusWnd = GetFocus()
                RingForm.Show
                SetForegroundWindow Handle
                SetFocusX FocusWnd
            End If
        End If
    'If inbox email status received
    ElseIf Left$(Command, 14) = "Inbox-Unread: " Then
        InboxUnread = InboxUnread + Val(Right$(Command, Len(Command) - InStr(Command, " ")))
        Call UpdateEmail
        
    'If folders email status received
    ElseIf Left$(Command, 16) = "Folders-Unread: " Then
        FolderUnread = FolderUnread + Val(Right$(Command, Len(Command) - InStr(Command, " ")))
        Call UpdateEmail
        
    'If nick changed
    ElseIf Left$(Command, 3) = "REA" Then
        If Split(Command)(3) = Login Then
            Command = Split(Command, " ")(4)
            Nick = DeMorph(Command, True)
            If Status <> 0 Then
                frmMain.lblNick.Caption = Nick & " (" & StatusConv(StatusCode(Status)) & ")"
            Else
                frmMain.lblNick.Caption = Nick
            End If
            frmMain.txtNick.Text = Nick
        End If
    'If contact removed from a list
    'REM 9 FL 5237 hasankhan1@msn.com
    ElseIf Left$(Command, 3) = "REM" Then
        Lst = Split(Command)(2)
        Command = Split(Command)(4)
        'If contact deleted
        If Lst = "FL" Then
            frmMain.tvwBuddies.Nodes.Remove Command
            SetBuddyProperty Command, "forward", False
            Call UpdateListCount
        ElseIf Lst = "AL" Then
            SetBuddyProperty Command, "allow", False
        'If contact is unblocked
        ElseIf Lst = "BL" Then
            SetBuddyProperty Command, "block", False
            Call UpdateList(Command)
        ElseIf Lst = "RL" Then
            SetBuddyProperty Command, "reverse", False
            If IsInList(Command) = True Then
                frmMain.tvwBuddies.Nodes(Command).BackColor = 16119285 'RGB(245,245,245)
            End If
            If RcLoggedIn = False Then
                If LastDeleteAlert <> Command Then
                    LastDeleteAlert = Command
                    MsgBox Command & " has deleted you.", vbInformation
                End If
            End If
        End If
        
    'If contact added to a list
    'ADD 9 FL 1477 born_intelligent@msn.com born_intelligent@msn.com
    ElseIf Left$(Command, 3) = "ADD" Then
        Lst = Split(Command)(2)
        eml = Split(Command)(4)
        If Lst = "FL" Then
            SetBuddyProperty CStr(eml), "forward", True
            Call UpdateList(CStr(eml), True)
        'If someone adds you
        ElseIf Lst = "RL" Then
            SetBuddyProperty CStr(eml), "reverse", True
            If GetBuddyProperty(CStr(eml), "forward") <> "True" Then
                ShowAddContactForm CStr(eml), DeMorph(Split(Command)(5), True)
            Else
                frmMain.tvwBuddies.Nodes(eml).BackColor = -2147483643
                If LastAddAlert <> eml Then
                    LastAddAlert = eml
                    MsgBox eml & " has added you to his/her contact list.", vbInformation
                End If
            End If
        'if contact is allowed
        ElseIf Lst = "AL" Then
            SetBuddyProperty CStr(eml), "allow", True
        'If contact is blocked
        ElseIf Lst = "BL" Then
            SetBuddyProperty CStr(eml), "block", True
            Call UpdateList(CStr(eml))
        End If
        
    'If new mail message received
    ElseIf Left$(Command, 6) = "From: " Then
        NewMail_FromName = Right$(Command, Len(Command) - 6)
        
    'Destination folder of new mail
    ElseIf Left$(Command, 13) = "Dest-Folder: " Then
        If NewMail_FromName <> vbNullString Then
            NewMail_Folder = Split(Command)(1)
            If NewMail_Folder = "ACTIVE" Then
                InboxUnread = InboxUnread + 1
                NewMail_Folder = "Inbox"
            Else
                FolderUnread = FolderUnread + 1
            End If
            Call UpdateEmail
        End If
        
    ElseIf Left$(Command, 9) = "Subject: " Then
        NewMail_Subject = Right$(Command, Len(Command) - 9)
        
    'New mail sender
    ElseIf Left$(Command, 11) = "From-Addr: " Then
        NewMail_FromEmail = Split(Command)(1)
        If Popups = True Then
            ShowPopup "You have received a new e-mail message " & vbCrLf & "from " & NewMail_FromName & vbCrLf & vbCrLf & PopupBreak(NewMail_Subject, False), "Email"
        End If
        LogStatus "New mail in " & NewMail_Folder & " from " & NewMail_FromEmail
        frmMain.lblStatus.Caption = "New mail in " & NewMail_Folder & " from " & NewMail_FromEmail
        NewMail_FromName = vbNullString
        NewMail_Folder = vbNullString
        frmMain.tmrAnimator.Enabled = True
        
    'If mail unread changed
    ElseIf Left$(Command, 12) = "Src-Folder: " Then
        DumpMail = Right$(Command, Len(Command) - 12)
        
    ElseIf Left$(Command, 15) = "Message-Delta: " Then
        If DumpMail = "ACTIVE" Then
            InboxUnread = InboxUnread - Val(Right$(Command, Len(Command) - 15))
            If InboxUnread < 0 Then InboxUnread = 0
            DumpMail = "Inbox"
        ElseIf UCase$(DumpMail) <> "TRASH" Then
            FolderUnread = FolderUnread - Val(Right$(Command, Len(Command) - 15))
            If FolderUnread < 0 Then FolderUnread = 0
        End If
        Call UpdateEmail
        DumpMail = vbNullString
        
    ElseIf Left$(Command, 4) = "CHG " And InitStatus = True Then
        LogStatus "Signed in."
        frmMain.lblStatus.Caption = "Signed in."
        frmMain.tvwBuddies.Nodes(1).EnsureVisible
        InitStatus = False
        Call UpdateListCount
        
    ElseIf Left$(Command, 4) = "CHG " Then
        Command = Left$(Command, Len(Command) - 2)
        frmMain.mnuStatusList(Status).Checked = False
        Status = StatusCode(Right$(Command, 3))
        frmMain.mnuStatusList(Status).Checked = True
        Call UpdateStatusImage
        If Status <> 0 Then
            frmMain.lblNick.Caption = Nick & " (" & StatusConv(StatusCode(Status)) & ")"
        Else
            frmMain.lblNick.Caption = Nick
        End If
        
    ElseIf UCase$(Left$(Command, 5)) = "SID: " Then
        Inbox_Sid = Val(Right$(Command, Len(Command) - 5))
        
    ElseIf UCase$(Left$(Command, 4)) = "KV: " Then
        Inbox_Kv = Val(Right$(Command, Len(Command) - 4))
    
    ElseIf Left$(Command, 9) = "MSPAuth: " Then
        Inbox_MSPAuth = Right$(Command, Len(Command) - 9)
        
    ElseIf Left$(Command, 4) = "URL " Then
        Inbox_Rru = Split(Command)(2)
        Inbox_Url = Split(Command)(3)
        Inbox_Id = Split(Command)(4)
        Call OpenMsnUrl
    End If
    Data = Right$(Data, Len(Data) - InStr(Data, vbCrLf) - 1)
Loop Until Data = vbNullString
End Sub

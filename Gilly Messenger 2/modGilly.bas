Attribute VB_Name = "modGilly"
Public Sub ResetCollection(SrcCollection As Collection)
    Set SrcCollection = Nothing
    Set SrcCollection = New Collection
End Sub

Public Function LoginCached(cLogin As String) As Boolean
For X = 0 To frmSignIn.cmbLogin.ListCount
    If frmSignIn.cmbLogin.List(X) = cLogin Then
        LoginCached = True
        Exit Function
    End If
    DoEvents
Next X
LoginCached = False
End Function

Public Sub MsnSend(Data As String, TID As Long, Connection As Winsock)
On Error Resume Next
Connection.SendData Data & vbCrLf
TID = TID + 1
Debug.Print "--> " & Data
End Sub

Public Function StatusConv(Data As String) As String
Select Case Data
Case "NLN"
    StatusConv = "Online"
Case "Online"
    StatusConv = "NLN"
Case "BSY"
    StatusConv = "Busy"
Case "Busy"
    StatusConv = "BSY"
Case "AWY"
    StatusConv = "Away"
Case "Away"
    StatusConv = "AWY"
Case "PHN"
    StatusConv = "On The Phone"
Case "On The Phone"
    StatusConv = "PHN"
Case "BRB"
    StatusConv = "Be Right Back"
Case "Be Right Back"
    StatusConv = "BRB"
Case "LUN"
    StatusConv = "Out To Lunch"
Case "Out To Lunch"
    StatusConv = "LUN"
Case "IDL"
    StatusConv = "Idle"
Case "Idle"
    StatusConv = "IDL"
Case "FLN"
    StatusConv = "Offline"
Case "HDN"
    StatusConv = "Offline"
Case "Appear Offline"
    StatusConv = "HDN"
End Select
End Function

Public Function Morph(Text As String, Full As Boolean) As String
On Error Resume Next

If Full = True Then
    Text = Replace$(Text, Chr$(vbKeySpace) & Chr$(vbKeySpace), Chr$(160) & Chr$(vbKeySpace))
    If Right$(Text, 1) = Chr$(vbKeySpace) Then
        Mid$(Text, 1, 1) = Chr$(160)
    End If
    If Left$(Text, 1) = Chr$(vbKeySpace) Then
        Mid$(Text, Len(Text), 1) = Chr$(160)
    End If
End If

Dim Char As String, Code As Long
For X = 1 To Len(Text)
    Char = Mid$(Text, X, 1)
    Code = Asc(Char)
    If Code > 191 Then
        Char = Chr$(195) & Chr$(192 Or Code)
    ElseIf Code > 159 Then
        Char = Chr$(194) & Char
    ElseIf Code > 157 Then
        Char = Chr$(197) & Chr$(194 Or Code)
    ElseIf Code > 156 Then
        Char = Chr$(194) & Char
    ElseIf Code > 153 Then
        Char = Chr$(197) & Chr$(194 Or Code)
    ElseIf Code > 142 Then
        Char = Chr$(194) & Char
    ElseIf Code > 141 Then
        Char = Chr$(197) & Chr$(194 Or Code)
    ElseIf Code > 140 Then
        Char = Chr$(194) & Char
    ElseIf Code > 137 Then
        Char = Chr$(197) & Chr$(194 Or Code)
    ElseIf Full = True Then
        If Not (Code > 32 And Code < 127) Then
            Char = Hex(Code)
            Char = String$(2 - Len(Char), "0") & Char
            Char = "%" & Char
        End If
    End If
    If Char = "ƒ" Then Char = "Æ’"
    If Char = "†" Then Char = "â€ "
    Morph = Morph & Char
Next
End Function

Public Function DeMorph(Text, Full As Boolean) As String
On Error Resume Next
Dim Char As String, Key As Long
For X = 1 To Len(Text)
    Char = Mid$(Text, X, 1)
    Key = Asc(Char)
    Code = Asc(Mid$(Text, X + 1, 1))
    If Char = "%" And Full = True Then
        DeMorph = DeMorph & Chr$(Val("&H" & Mid$(Text, X + 1, 2)))
        X = X + 2
    ElseIf Key = 195 Then
        DeMorph = DeMorph & Chr$(192 Or Code)
        X = X + 1
    ElseIf Key = 197 Then
        DeMorph = DeMorph & Chr$(194 Or Code)
        X = X + 1
    Else
        DeMorph = DeMorph & Char
    End If
Next
DeMorph = Replace$(DeMorph, Chr$(194), vbNullString)
DeMorph = Replace$(DeMorph, "Æ’", "ƒ")
DeMorph = Replace$(DeMorph, "â€ ", "†")
End Function

Public Function IconConv(Status As String) As Long
If Status = "Online" Then
    IconConv = 1
ElseIf Status = "Offline" Then
    IconConv = 2
ElseIf Status = "Busy" Or Status = "On The Phone" Then
    IconConv = 3
Else
    IconConv = 4
End If
End Function

Public Sub UpdateList(Email As String, Optional Full As Boolean)
Dim TempBlock As String, TempIgnore As String
On Error Resume Next
Temp = vbNullString
Temp = frmMain.tvwBuddies.SelectedItem.Key
If Full = True Then
    frmMain.tvwBuddies.Nodes.Remove Email
    If GetBuddyStatus(Email) = "Offline" Then
        frmMain.tvwBuddies.Nodes.Add "Offline", tvwChild, Email, Email
    Else
        frmMain.tvwBuddies.Nodes.Add "Online", tvwChild, Email, Email
    End If
    If InitStatus = False Then
        Call UpdateListCount
    End If
End If
frmMain.tvwBuddies.Nodes(Temp).Selected = True
Temp = vbNullString
If IsIgnored(Email) = True Then
    TempIgnore = "(Ignored)"
Else
    TempIgnore = vbNullString
End If
If GetBuddyStatus(Email) = "Online" Or GetBuddyStatus(Email) = "Offline" Then
    If ViewContactsByEmail = False Then
        frmMain.tvwBuddies.Nodes(Email).Text = GetBuddyNick(Email) & RTrim$(" " & GetBuddyBlock(Email)) & RTrim$(" " & TempIgnore)
    Else
        frmMain.tvwBuddies.Nodes(Email).Text = Email & RTrim$(" " & GetBuddyBlock(Email)) & RTrim$(" " & TempIgnore)
    End If
Else
    If ViewContactsByEmail = False Then
        frmMain.tvwBuddies.Nodes(Email).Text = GetBuddyNick(Email) & " (" & GetBuddyStatus(Email) & ")" & RTrim$(" " & GetBuddyBlock(Email)) & RTrim$(" " & TempIgnore)
    Else
        frmMain.tvwBuddies.Nodes(Email).Text = Email & " (" & GetBuddyStatus(Email) & ")" & RTrim$(" " & GetBuddyBlock(Email)) & RTrim$(" " & TempIgnore)
    End If
End If
frmMain.tvwBuddies.Nodes(Email).Image = IconConv(GetBuddyStatus(Email))
If GetBuddyProperty(Email, "reverse") <> "True" And SignInMode <> "Online" Then
    frmMain.tvwBuddies.Nodes(Email).BackColor = 16119285 'RGB(245,245,245)
End If
End Sub

Public Sub ChangeNick(NewNick As String)
If SignedIn = False Or NewNick = vbNullString Then
    Exit Sub
End If
MsnSend "REA " & TrialID & " " & Login & " " & Morph(NewNick, True), TrialID, frmMain.wskMSN
End Sub

Public Sub Invite(Contact As String, TID As Long, Connection As Winsock)
MsnSend "CAL " & TID & " " & Contact, TID, Connection
End Sub

Public Sub MessageAll(Message As String)
On Error Resume Next
For Each Form In Forms
    If Form.Name = "frmChat" Then
        Form.txtMessage.Text = Message
        Form.cmdSend.Value = 1
    End If
    DoEvents
Next
End Sub

Public Sub StartChat(Email As String, Optional Message As String, Optional BlockCheck As Boolean = False, Optional BlockCheckByUser As Boolean = False)
Dim Target As String
On Error Resume Next
Target = Email
If BlockCheck = False Then
    X = 0
    X = OpenChats(Target)
    If X > 0 Then
        Set TempForm = FindChat(, CLng(X))
        Temp = vbNullString
        Temp = TempForm.ChatBuddies(Email)
        X = TempForm.ChatBuddies.Count
        If (Temp <> vbNullString And X = 1) Or (X = 0) Then
            TempForm.Visible = True
            SetForegroundWindow TempForm.hwnd
            SetFocusX TempForm.hwnd
            If Message <> vbNullString Then
                MessageUser Email, Message
            End If
            Exit Sub
        Else
            OpenChats.Remove Target
        End If
    End If
End If

If Status = 7 And RcLoggedIn = False Then
    MsgBox "Not allowed when offline.", vbExclamation
    Exit Sub
End If

Set CallForm = New frmChat
CallForm.Tag = "CALL" & Target
CallForm.lblStatus.Tag = Target

If BlockCheck = True Then
    CallForm.BlockCheck = True
    CallForm.BlockCheckByUser = BlockCheckByUser
Else
    CallForm.lblBuddy.Caption = Target
    CallForm.Caption = GetBuddyNick(Target)
    CallForm.Show
    OpenChats.Add CallForm.hwnd, Target
    If Message <> vbNullString Then
        CallForm.txtMessage.Text = Message
        CallForm.cmdSend.Value = 1
    End If
End If

CallForms.Add CallForm
MsnSend "XFR " & TrialID & " SB", TrialID, frmMain.wskMSN
End Sub

Public Sub Block(Email As String)
MsnSend "REM " & TrialID & " AL " & Email, TrialID, frmMain.wskMSN
MsnSend "ADD " & TrialID & " BL " & Email & " " & Email, TrialID, frmMain.wskMSN
End Sub

Public Sub UnBlock(Email As String)
MsnSend "REM " & TrialID & " BL " & Email, TrialID, frmMain.wskMSN
MsnSend "ADD " & TrialID & " AL " & Email & " " & Email, TrialID, frmMain.wskMSN
End Sub

Public Sub AddContact(Email As String)
If InStr(Email, "@") = 0 Then
    Email = Email & "@hotmail.com"
End If
MsnSend "ADD " & TrialID & " FL " & Email & " " & Email, TrialID, frmMain.wskMSN
If GetBuddyBlock(Email) = "" And GetBuddyProperty(Email, "allow") <> "True" Then
    MsnSend "ADD " & TrialID & " AL " & Email & " " & Email, TrialID, frmMain.wskMSN
End If
End Sub

Public Sub ChatMsgSend(Header As String, Message As String, TID As Long, Connection As Winsock)
Connection.SendData "MSG " & TID & " N " & Len(Header) + 4 + Len(Message) & vbCrLf & Header & vbCrLf & vbCrLf & Message
TID = TID + 1
End Sub

Public Function BuddyConv(Data As String) As Buddy
On Error Resume Next
'ILN 7 NLN hasankhan1@msn.com %42%75%5A%7A%7A%5A%5A%7A
'NLN NLN hasankhan1@msn.com %42%75%5A%7A%7A%5A%5A%7A
If Left$(Data, 3) = "ILN" Then
    Data = Right$(Data, Len(Data) - InStr(Data, " "))
End If
Temp = Split(Data)(1)
BuddyConv.Status = StatusConv(Temp)
BuddyConv.Email = Split(Data)(2)
BuddyConv.Nick = Split(Data)(3)
BuddyConv.Nick = DeMorph(BuddyConv.Nick, True)
End Function

Public Sub UpdateEmail()
If InboxUnread = 0 Then
    frmMain.lblEmail.Caption = "No new e-mail messages"
Else
    frmMain.lblEmail.Caption = InboxUnread & " new e-mail messages"
End If
frmMain.lblEmail.ToolTipText = "Inbox : " & InboxUnread & Space$(4) & "Folders : " & FolderUnread
End Sub

Public Sub SetAutoMsg(Msg As String)
If frmMain.mnuAutoMessage.Checked = False Then frmMain.mnuAutoMessage.Checked = True
AutoMsg = Msg
frmMain.lblStatus.Caption = "Auto Message set."
End Sub

Public Sub LoadFileMenu(FileMenu As Object, Path As String, Pattern As String)
Temp = Dir$(Path & "\" & Pattern)
Do Until Temp = vbNullString
    Load FileMenu(FileMenu.Count)
    FileMenu(FileMenu.Count - 1).Caption = "&" & Left$(Temp, InStrRev(Temp, ".") - 1)
    FileMenu(FileMenu.Count - 1).Enabled = True
    FileMenu(FileMenu.Count - 1).Tag = Path & "\" & Temp
    Temp = Dir$
Loop
If FileMenu.Count > 1 Then
    FileMenu(0).Visible = False
End If
End Sub

Public Function Alias(Text As String) As String
Alias = Text
Alias = Replace$(Alias, "(Time)", Time, , , vbTextCompare)
Alias = Replace$(Alias, "(Date)", Date$, , , vbTextCompare)
Alias = Replace$(Alias, "(Now)", Now, , , vbTextCompare)
Alias = Replace$(Alias, "(Day)", WeekdayName(Weekday(Now)), , , vbTextCompare)
Alias = Replace$(Alias, "(Nick)", Nick, , , vbTextCompare)
Alias = Replace$(Alias, "(IP)", frmMain.wskMSN.LocalIP, , , vbTextCompare)
Alias = Replace$(Alias, "(CRLF)", vbCrLf, , , vbTextCompare)
Alias = Replace$(Alias, "(VER)", App.Major & "." & App.Minor & "." & App.Revision, , , vbTextCompare)
Temp = vbNullString
Temp = GetSong
If Temp = vbNullString Then Temp = Nick
Alias = Replace$(Alias, "(Song)", Temp, , , vbTextCompare)
End Function

Public Sub Signout()
If frmMain.mnuSignIn.Caption = "Sign &Out" Then
    Call frmMain.mnuSignIn_Click
End If
End Sub

Public Sub SignIn(sLogin, sPassword)
On Error Resume Next
Call ResetSockets
Login = LCase$(sLogin)
If Len(sPassword) > 16 Then
    Password = Left$(sPassword, 16)
Else
    Password = sPassword
End If
frmMain.picSignIn.Visible = False
frmMain.mnuSignIn.Caption = "&Cancel Sign In"
frmMain.lblStatus.Caption = "Connecting..."
frmMain.wskMSN.Connect
End Sub

Public Sub UpdateChatCaption(Email As String, NewNick As String)
On Error Resume Next
FindChat(Email).Caption = NewNick
End Sub

Public Sub MessageUser(Email As String, Message As String)
On Error Resume Next
Set TempForm = FindChat(Email)
Temp = TempForm.Name
If Temp = "frmChat" Then
    TempForm.txtMessage.Text = Message
    TempForm.cmdSend.Value = 1
Else
    StartChat Email, Message
End If
End Sub

Public Function ListOnline() As String
On Error Resume Next
Dim TempEmail As String
For X = 3 To frmMain.tvwBuddies.Nodes.Count
    If frmMain.tvwBuddies.Nodes(X).Key <> "NoneOnline" And frmMain.tvwBuddies.Nodes(X).Key <> "NoneOffline" Then
        TempEmail = frmMain.tvwBuddies.Nodes(X).Key
        If frmMain.tvwBuddies.Nodes(X).Parent.Key <> "Offline" Then
            If IsIgnored(TempEmail) = True Then
                TempIgnore = "(Ignored)"
            Else
                TempIgnore = vbNullString
            End If
            If frmMain.tvwBuddies.Nodes(X).Parent.Key = "Online" Then
                ListOnline = ListOnline & TempEmail & " - " & GetBuddyNick(TempEmail) & RTrim$(" " & GetBuddyBlock(TempEmail)) & RTrim$(" " & TempIgnore) & vbCrLf
            Else
                ListOnline = ListOnline & TempEmail & " - " & GetBuddyNick(TempEmail) & " (" & GetBuddyStatus(TempEmail) & ")" & RTrim$(" " & GetBuddyBlock(TempEmail)) & RTrim$(" " & TempIgnore) & vbCrLf
            End If
        End If
        DoEvents
    End If
Next
If ListOnline <> vbNullString Then
    ListOnline = Left$(ListOnline, Len(ListOnline) - 2)
Else
    ListOnline = "None of contacts is online."
End If
End Function

Public Function ListChats() As String
    For Each Form In Forms
        If Form.Name = "frmChat" Then
            ListChats = ListChats & Form.lblStatus.Tag & " - " & Form.lblStatus.Caption & vbCrLf
        End If
    Next
End Function

Public Sub ShowPopup(Message As String, Tag As String)
On Error Resume Next
Dim Popup As Form
If PopupHeight >= frmPopup.Height Then
    Set Popup = New frmPopup
    Popup.lblMessage.Caption = Message
    Popup.Picture = frmMain.imglstPictures.ListImages("POPUP").Picture
    Popup.Tag = Tag
    Handle = GetForegroundWindow()
    FocusWnd = GetFocus()
    Popup.hActiveWnd = Handle
    Popup.hFocusWnd = FocusWnd
    Popup.Show
End If
End Sub

Public Sub DeleteContact(Email As String)
If MsgBox("Are you sure you want to delete " & Email & " from your contact list?", vbQuestion Or vbYesNo, "Delete Contact") = vbYes Then
    MsnSend "REM " & TrialID & " FL " & Email, TrialID, frmMain.wskMSN
End If
End Sub

Public Function GetSong() As String
Dim Buffer As String
Handle = FindWindow("Winamp v1.x", vbNullString)
If Handle = 0 Then Exit Function
Buffer = Space$(100)
GetWindowText Handle, Buffer, 100
If InStr(Buffer, " - Winamp") = 0 Then Exit Function
Buffer = Left$(Buffer, InStrRev(Buffer, " - Winamp") - 1)
Buffer = Right$(Buffer, Len(Buffer) - 3)
If Right$(Buffer, 4) Like ".???" Then
    Buffer = Left$(Buffer, Len(Buffer) - 4)
End If
GetSong = Buffer
End Function

Public Function ColorConv(HexCode As String, Optional Invert As Boolean) As Long
On Error Resume Next
If Len(HexCode) > 6 Then HexCode = Left$(HexCode, 6)
If Invert = True Then
    If Len(HexCode) < 6 Then HexCode = HexCode & String$(6 - Len(HexCode), "0")
    ColorConv = RGB(Val("&H" & Left$(HexCode, 2)), Val("&H" & Mid$(HexCode, 3, 2)), Val("&H" & Right$(HexCode, 2)))
Else
    If Len(HexCode) < 6 Then HexCode = String$(6 - Len(HexCode), "0") & HexCode
    ColorConv = RGB(Val("&H" & Right$(HexCode, 2)), Val("&H" & Mid$(HexCode, 3, 2)), Val("&H" & Left$(HexCode, 2)))
End If
End Function

Public Sub ChangeStatus(Index As Integer)
If frmMain.wskMSN.State = sckConnected Then
    MsnSend "CHG " & TrialID & " " & StatusCode(Index), TrialID, frmMain.wskMSN
End If
End Sub

Public Function StatusCode(strCode As Variant) As Variant
Select Case strCode
    Case "NLN"
        StatusCode = 0
    Case 0
        StatusCode = "NLN"
    Case "BSY"
        StatusCode = 1
    Case 1
        StatusCode = "BSY"
    Case "BRB"
        StatusCode = 2
    Case 2
        StatusCode = "BRB"
    Case "AWY"
        StatusCode = 3
    Case 3
        StatusCode = "AWY"
    Case "PHN"
        StatusCode = 4
    Case 4
        StatusCode = "PHN"
    Case "LUN"
        StatusCode = 5
    Case 5
        StatusCode = "LUN"
    Case "IDL"
        StatusCode = 6
    Case 6
        StatusCode = "IDL"
    Case "HDN"
        StatusCode = 7
    Case 7
        StatusCode = "HDN"
End Select
End Function

Public Function IsIgnored(Email As String) As Boolean
On Error Resume Next
Temp = vbNullString
Temp = BuddyIgnore(Email)
If Temp = Email Then IsIgnored = True Else IsIgnored = False
End Function

Public Sub Ignore(Email As String)
On Error Resume Next
BuddyIgnore.Add Email, Email
SaveSetting "Gilly Messenger", "Ignore List\" & Login, Email, ""
Call UpdateList(Email)
End Sub

Public Sub Unignore(Email As String)
On Error Resume Next
BuddyIgnore.Remove Email
DeleteSetting "Gilly Messenger", "Ignore List\" & Login, Email
Call UpdateList(Email)
End Sub

Public Function FindChat(Optional Email As String, Optional Handle As Long) As Form
For Each Form In Forms
    If Form.Name = "frmChat" Then
        If Email <> vbNullString Then
            If Form.lblStatus.Tag = Email Then
                Set FindChat = Form
                Exit For
            End If
        ElseIf Handle <> 0 Then
            If Form.hwnd = Handle Then
                Set FindChat = Form
                Exit For
            End If
        End If
    End If
Next
End Function

Public Sub UpdateStatusImage()
StatusImage = IconConv(StatusConv(StatusCode(Status)))
Call frmMain.Form_Resize
End Sub

Public Function PopupBreak(ByVal strText As String, Multiline As Boolean) As String
On Error Resume Next
If Multiline = True Then
    If Len(strText) > 50 Then
        strText = Left$(strText, 50)
        PopupBreak = Left$(strText, 25) & vbCrLf & Right$(strText, 25) & "..."
    ElseIf Len(strText) > 25 Then
        PopupBreak = Left$(strText, 25) & vbCrLf & Right$(strText, Len(strText) - 25)
    Else
        PopupBreak = strText
    End If
Else
    If Len(strText) > 25 Then
        PopupBreak = Left$(strText, 22) & "..."
    Else
        PopupBreak = strText
    End If
End If
End Function

Public Function GetBuddyStatus(Email) As String
On Error Resume Next
GetBuddyStatus = ContactList.Item(Email).Item("status")
If GetBuddyStatus = "" Then
    GetBuddyStatus = "Offline"
End If
End Function

Public Function GetBuddyComment(Email) As String
On Error Resume Next
GetBuddyComment = BuddyComment(Email)
End Function

Public Sub SetBuddyComment(Email As String, BdyComment As String)
On Error Resume Next
BuddyComment.Remove Email
If BdyComment <> vbNullString Then
    BuddyComment.Add BdyComment, Email
    SaveSetting "Gilly Messenger", "Comments\" & Login, Email, BdyComment
Else
    DeleteSetting "Gilly Messenger", "Comments\" & Login, Email
End If
End Sub

Public Function GetBuddyNick(Email) As String
On Error Resume Next
GetBuddyNick = ContactList.Item(Email).Item("nick")
If GetBuddyNick = vbNullString Then
    GetBuddyNick = Email
End If
End Function

Public Sub LogChat(Text As String, Title As String)
    If Login = vbNullString Or frmMain.mnuChatLogger.Checked = False Or Title = vbNullString Then
        Exit Sub
    End If
    On Error GoTo Handler
    If MakeSureDirectoryPathExists(ChatLogDir & "\" & Login & "\") = 0 Then
        Call frmMain.mnuChatLogger_Click
        MsgBox "Invalid chat log directory.", vbCritical, "Error!"
        Exit Sub
    End If
    Open ChatLogDir & "\" & Login & "\" & Title & ".txt" For Append As #3
    Print #3, "[" & Now & "] " & Text
    Close #3
Handler:
End Sub

Public Sub LogStatus(Text As String)
    If Login = vbNullString Or frmMain.mnuStatusLogger.Checked = False Then Exit Sub
    On Error GoTo Handler
    If MakeSureDirectoryPathExists(StatusLogDir) = 0 Then
        frmMain.mnuStatusLogger.Checked = False
        MsgBox "Invalid status log directory.", vbCritical, "Error!"
        Exit Sub
    End If
    Open StatusLogDir & "\" & Login & ".txt" For Append As #3
    Print #3, "[" & Now & "] " & Text
    Close #3
Handler:
End Sub

Public Sub ShowBuddyPropPage(Email As String)
On Error Resume Next
frmContactProperties.txtNick = GetBuddyNick(Email)
frmContactProperties.txtEmail = Email
frmContactProperties.txtStatus = GetBuddyStatus(Email)
frmContactProperties.txtComment = GetBuddyComment(Email)
frmContactProperties.txtComment.SelStart = Len(frmContactProperties.txtComment.Text)
frmContactProperties.Show , frmMain
End Sub

Public Function GetBuddyBlock(Email As String) As String
On Error Resume Next
Dim TempEml As String
TempEml = Email
Temp = ContactList.Item(Email).Item("block")
Email = TempEml
If Temp = "True" Then
    GetBuddyBlock = "(Blocked)"
End If
End Function

Public Function IsOnline(Email As String) As Boolean
If GetBuddyStatus(Email) <> "Offline" Then
    IsOnline = True
Else
    IsOnline = False
End If
End Function

Public Function IsInList(Email As String) As Boolean
    On Error GoTo Handler
    Temp = frmMain.tvwBuddies.Nodes(Email).Key
    IsInList = True
    Exit Function
Handler:
    IsInList = False
End Function

Public Sub UpdateListCount()
    On Error Resume Next
    If frmMain.tvwBuddies.Nodes("Online").Children > 0 Then
        Temp = vbNullString
        Temp = frmMain.tvwBuddies.Nodes("NoneOnline").Text
        If Temp <> vbNullString And frmMain.tvwBuddies.Nodes("Online").Children > 1 Then
            frmMain.tvwBuddies.Nodes.Remove ("NoneOnline")
        End If
        Temp = vbNullString
        Temp = frmMain.tvwBuddies.Nodes("NoneOnline").Text
        If Temp = vbNullString Then
            frmMain.tvwBuddies.Nodes("Online").Text = "Online (" & frmMain.tvwBuddies.Nodes("Online").Children & ")"
        End If
    Else
        frmMain.tvwBuddies.Nodes.Add "Online", tvwChild, "NoneOnline", "None of your contacts are online"
        frmMain.tvwBuddies.Nodes("NoneOnline").ForeColor = Val("&H808080")
        frmMain.tvwBuddies.Nodes("Online").Text = "Online"
    End If
    If frmMain.tvwBuddies.Nodes("Offline").Children > 0 Then
        Temp = vbNullString
        Temp = frmMain.tvwBuddies.Nodes("NoneOffline").Text
        If Temp <> vbNullString And frmMain.tvwBuddies.Nodes("Offline").Children > 1 Then
            frmMain.tvwBuddies.Nodes.Remove ("NoneOffline")
        End If
        Temp = vbNullString
        Temp = frmMain.tvwBuddies.Nodes("NoneOffline").Text
        If Temp = vbNullString Then
            frmMain.tvwBuddies.Nodes("Offline").Text = "Not Online (" & frmMain.tvwBuddies.Nodes("Offline").Children & ")"
        End If
    Else
        frmMain.tvwBuddies.Nodes.Add "Offline", tvwChild, "NoneOffline", "None of your contacts are offline"
        frmMain.tvwBuddies.Nodes("NoneOffline").ForeColor = Val("&H808080")
        frmMain.tvwBuddies.Nodes("Offline").Text = "Not Online"
    End If
End Sub

Public Sub ResetSockets()
    frmMain.wskSSL.Close
    Call frmMain.wskSSL_Close
    frmMain.wskMSN.Close
    Call frmMain.wskMSN_Close
End Sub

Public Function GetSl() As Integer
GetSl = DateDiff("s", LoginTime, Now())
End Function

Public Sub OpenMsnUrl()
    Dim strTemp As String, FileNum As Integer
    strTemp = String(100, Chr$(0))
    'Get the temporary path
    GetTempPath 100, strTemp
    'strip the rest of the buffer
    strTemp = Left$(strTemp, InStr(strTemp, Chr$(0)) - 1)
    strTemp = strTemp & "\gm" & MsnUrlType(1) & MsnUrlType.Count & ".htm"
    FileNum = FreeFile
    Open strTemp For Output As #FileNum
    Print #FileNum, "<html>"
    Print #FileNum, " <head>"
    Print #FileNum, "  <noscript>"
    Print #FileNum, "   <meta http-equiv=Refresh content='0; url=http://www.hotmail.com'>"
    Print #FileNum, "  </noscript>"
    Print #FileNum, " </head>"
    Print #FileNum, ""
    Print #FileNum, " <body onload='document.pform.submit(); '>"
    Print #FileNum, "  <form name='pform' action='" & Inbox_Url & "' method='POST'>"
    Print #FileNum, "   <input type='hidden' name='mode' value='ttl'>"
    Print #FileNum, "   <input type='hidden' name='login' value='" & Split(Login, "@")(0) & "'>"
    Print #FileNum, "   <input type='hidden' name='username' value='" & Login & "'>"
    Print #FileNum, "   <input type='hidden' name='sid' value='" & Inbox_Sid & "'>"
    Print #FileNum, ""
    Print #FileNum, "   <input type='hidden' name='kv' value='" & Inbox_Kv & "'>"
    Print #FileNum, "   <input type='hidden' name='id' value='" & Inbox_Id & "'>"
    Print #FileNum, "   <input type='hidden' name='sl' value='" & GetSl & "'>"
    Print #FileNum, "   <input type='hidden' name='rru' value='" & Inbox_Rru & "'>"
    Print #FileNum, "   <input type='hidden' name='auth' value='" & Inbox_MSPAuth & "'>"
    Print #FileNum, "   <input type='hidden' name='creds' value='" & MD5Encrypt(Inbox_MSPAuth & GetSl & Password) & "'>"
    Print #FileNum, ""
    If MsnUrlType(1) = "00" Then
        Print #FileNum, "   <input type='hidden' name='svc' value='mail'>"
    End If
    Print #FileNum, "   <input type='hidden' name='js' value='yes'>"
    Print #FileNum, "  </form>"
    Print #FileNum, " </body>"
    Print #FileNum, "</html>"
    Close #FileNum
    ShellExecute frmMain.hwnd, "open", strTemp, vbNullString, vbnullstirng, 1
    MsnFile.Add strTemp
    frmMain.tmrMsnFileKiller.Enabled = True
    'Kill strTemp
    MsnUrlType.Remove 1
End Sub

Public Sub RefreshList()
    For X = 3 To frmMain.tvwBuddies.Nodes.Count
        If frmMain.tvwBuddies.Nodes(X).Key <> "NoneOnline" And frmMain.tvwBuddies.Nodes(X).Key <> "NoneOffline" Then
            UpdateList frmMain.tvwBuddies.Nodes(X).Key, False
        End If
    Next
End Sub

Public Sub SetBuddyProperty(Email As String, Property As String, Value As String)
On Error Resume Next
X = 0
X = ContactList.Item(Email).Count
If X > 0 Then
    ContactList.Item(Email).Remove Property
    ContactList.Item(Email).Add Value, Property
Else
    Set LstBuddy = New Collection
    LstBuddy.Add Email, "email"
    LstBuddy.Add Value, Property
    ContactList.Add LstBuddy, Email
End If
End Sub

Public Function GetBuddyProperty(Email As String, Property As String) As String
On Error Resume Next
X = 0
X = ContactList.Item(Email).Count
If X > 0 Then
    GetBuddyProperty = ContactList.Item(Email).Item(Property)
End If
End Function

Public Sub ShowAddContactForm(ContactEmail As String, ContactNick As String)
    Set AddContactFrm = New frmAddContact
    AddContactFrm.lblEmail = Replace$(AddContactFrm.lblEmail, "[Email]", ContactNick & " [" & ContactEmail & "]")
    AddContactFrm.Tag = ContactEmail
    AddContactFrm.Show
End Sub

Public Sub ShowMe(srcForm As Form)
    If srcForm.WindowState <> vbMaximized Then
        srcForm.WindowState = vbNormal
    End If
    srcForm.Visible = True
    Temp = srcForm.Caption
    srcForm.Caption = Temp & "   "
    AppActivate srcForm.Caption
    srcForm.Caption = Temp
End Sub

Public Function AppIgnored(Email As String) As Boolean
    For X = 1 To 4
        If Email Like Choose(X, "messenger@microsoft.com", "*block*", "*check*", "*status*") Then
            AppIgnored = True
            Exit Function
        End If
    Next
End Function

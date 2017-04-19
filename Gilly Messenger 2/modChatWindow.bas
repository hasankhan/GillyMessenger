Attribute VB_Name = "modChatWindow"
Public Sub AddChat(ChatWindow As Form, cNick As String, Text As String, Font As String, Size As Integer, Color As Long, Bold As Boolean, Italic As Boolean)
On Error Resume Next
With ChatWindow
    Call LogChat(cNick & " :" & vbCrLf & Text, .lblStatus.Tag)
    .txtChat.SelStart = Len(.txtChat.Text)
    If .mnuEmoticons.Checked = True Then
        Call AddEmtChat(ChatWindow, cNick & " :" & vbCrLf, "MS Sans Serif", 10, RGB(60, 60, 60), False, False)
        Call AddEmtChat(ChatWindow, Text, Font, Size, Color, Bold, Italic)
    Else
        .txtChat.SelFontName = "MS Sans Serif"
        .txtChat.SelFontSize = 10
        .txtChat.SelColor = RGB(60, 60, 60)
        .txtChat.SelBold = False
        .txtChat.SelItalic = False
        .txtChat.SelText = cNick & " :" & vbCrLf
        .txtChat.SelFontName = Font
        .txtChat.SelFontSize = Size
        .txtChat.SelColor = Color
        .txtChat.SelBold = Bold
        .txtChat.SelItalic = Italic
        .txtChat.SelText = Text
    End If
    .txtChat.SelStart = Len(.txtChat.Text)
    .txtChat.SelText = vbCrLf
End With
End Sub

Public Sub AddEmtChat(ChatWindow As Form, Text As String, Font As String, Size As Integer, Color As Long, Bold As Boolean, Italic As Boolean)
On Error Resume Next
With ChatWindow
    For X = 1 To Len(Text)
        For Y = 0 To 89
            EmtLen = Len(Emoticons(Y, 0))
            If UCase$(Mid$(Text, X, EmtLen)) = Emoticons(Y, 0) Then
                TempCpText = Clipboard.GetText
                Clipboard.Clear
                Clipboard.SetData frmMain.imglstEmoticons.ListImages(Val(Emoticons(Y, 1))).Picture, vbCFBitmap
                .txtChat.Locked = False
                SendMessage .txtChat.hwnd, WM_PASTE, 0, 0
                .txtChat.Locked = True
                Clipboard.Clear
                Clipboard.SetText TempCpText
                X = X + EmtLen - 1
                EmtAdded = True
                Exit For
            End If
        Next Y
        If EmtAdded = False Then
            .txtChat.SelFontName = Font
            .txtChat.SelFontSize = Size
            .txtChat.SelColor = Color
            .txtChat.SelBold = Bold
            .txtChat.SelItalic = Italic
            .txtChat.SelText = Mid$(Text, X, 1)
        Else
            EmtAdded = False
        End If
    Next X
End With
End Sub

Public Sub ClearMsg(ChatWindow As Form)
On Error Resume Next
With ChatWindow
    If .txtMessage.Visible = True Then
        .txtMessage.Text = vbNullString
        .txtMessage.SetFocus
    Else
        .txtMask.Text = vbNullString
        .txtMask.SetFocus
    End If
End With
End Sub

Public Sub Comment(ChatWindow As Form, Text As String)
With ChatWindow
    .txtChat.SelFontSize = 10
    .txtChat.SelStart = Len(.txtChat.Text)
    .txtChat.SelFontName = "MS Sans Serif"
    .txtChat.SelBold = False
    .txtChat.SelItalic = False
    .txtChat.SelColor = RGB(60, 60, 60)
    .txtChat.SelText = Text & vbCrLf
End With
End Sub

Public Sub UpdateBuddies(ChatWindow As Form)
On Error Resume Next
If ChatWindow.lblBuddy.Tag <> "Hidden" Then
    With ChatWindow
        If .ChatBuddies.Count > 0 Then
            If .ChatBuddies(1) = .lblStatus.Tag And .Mimic <> vbNullString Then
                .lblBuddy.Caption = GetBuddyNick(.Mimic)
            Else
                .lblBuddy.Caption = .ChatBuddies(1)
            End If
            For X = 2 To .ChatBuddies.Count
                If .Mimic = .lblStatus.Tag And .Mimic <> vbNullString Then
                    .lblBuddy.Caption = .lblBuddy.Caption & ", " & .Mimic
                Else
                    .lblBuddy.Caption = .lblBuddy.Caption & ", " & .ChatBuddies(X)
                End If
            Next X
        Else
            If .Mimic <> vbNullString Then
                .lblBuddy.Caption = .Mimic
            Else
                .lblBuddy.Caption = .lblStatus.Tag
            End If
        End If
    End With
End If
End Sub

Public Function cFormat(Message As String)
cFormat = Space$(3) & Replace(Message, vbCrLf, vbCrLf & Space$(3))
End Function

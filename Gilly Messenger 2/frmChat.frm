VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmChat 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   7110
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5880
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   Icon            =   "frmChat.frx":0000
   KeyPreview      =   -1  'True
   ScaleHeight     =   474
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   392
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBottomBar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   180
      ScaleHeight     =   95
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   367
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5520
      Width           =   5535
      Begin VB.TextBox txtMessage 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   705
         Left            =   60
         MaxLength       =   1536
         MultiLine       =   -1  'True
         TabIndex        =   0
         Top             =   360
         Width           =   4395
      End
      Begin VB.TextBox txtMask 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   705
         IMEMode         =   3  'DISABLE
         Left            =   60
         MaxLength       =   1536
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   360
         Visible         =   0   'False
         Width           =   4395
      End
      Begin VB.CommandButton cmdSend 
         BackColor       =   &H00FAF1ED&
         Caption         =   "&Send"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   420
         Width           =   855
      End
      Begin VB.Image imgFont 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   120
         Picture         =   "frmChat.frx":058A
         Top             =   0
         Width           =   825
      End
      Begin VB.Image imgEmoticon 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1080
         Picture         =   "frmChat.frx":09A5
         Top             =   0
         Width           =   450
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Connecting..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00814D3C&
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1140
         UseMnemonic     =   0   'False
         Width           =   5295
      End
   End
   Begin VB.Timer tmrStatusR 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   2872
      Top             =   3120
   End
   Begin MSWinsockLib.Winsock wskChat 
      Left            =   1905
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cdHeader 
      Left            =   2385
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Text Styles (*.gts)|*.gts"
      FontName        =   "Tahoma"
      FontSize        =   10
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   4095
      Left            =   180
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1320
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   7223
      _Version        =   393217
      BackColor       =   16777215
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmChat.frx":0DDE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblGilly 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gilly"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008D2F11&
      Height          =   240
      Left            =   5160
      TabIndex        =   7
      Top             =   15
      Width           =   435
   End
   Begin VB.Image imgTopBarRight 
      Height          =   735
      Left            =   3960
      Top             =   0
      Width           =   1935
   End
   Begin VB.Image imgResize 
      Height          =   660
      Left            =   5160
      Top             =   6480
      Width           =   690
   End
   Begin VB.Label lblBuddy 
      Appearance      =   0  'Flat
      BackColor       =   &H00FAF1EB&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00814D3C&
      Height          =   270
      Left            =   180
      TabIndex        =   4
      Top             =   1080
      UseMnemonic     =   0   'False
      Width           =   5535
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuRandomColors 
         Caption         =   "&Random Colors"
      End
      Begin VB.Menu mnuMask 
         Caption         =   "Message &Mask"
      End
      Begin VB.Menu mnuTextStyler 
         Caption         =   "&Text Styler"
         Begin VB.Menu mnuTextStyle 
            Caption         =   "No Styles Available"
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu mnuStyleSep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuStyleOther 
            Caption         =   "&Other..."
         End
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuFont 
         Caption         =   "Change &Font"
      End
      Begin VB.Menu mnuColor 
         Caption         =   "Change &Color"
      End
      Begin VB.Menu mnuEmoticons 
         Caption         =   "Use &Emoticons"
      End
   End
   Begin VB.Menu mnuActions 
      Caption         =   "&Actions"
      Begin VB.Menu mnuInvite 
         Caption         =   "&Invite Someone"
      End
      Begin VB.Menu mnuBlock 
         Caption         =   "&Block"
      End
      Begin VB.Menu mnuIgnore 
         Caption         =   "&Ignore"
      End
      Begin VB.Menu mnuProfile 
         Caption         =   "View &Profile"
      End
      Begin VB.Menu mnuViewLog 
         Caption         =   "View &Log"
      End
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cTrialID As Long, ChatBuddy As String, cMsg As String, Data As String, Command As String ',Messages As New Collection
Dim LastMsg As String, LastMsgTime As String, QueMsgs As New Collection
Public MsgHeader As String, FName As String, FColor As String, FStyle As String, FBold As Boolean, FItalic As Boolean
Dim EmtLen As Long, MsgLen As Long, EmtAdded As Boolean
Public ChatBuddies As New Collection, ChatBuddyNick As New Collection
Public Mimic As String
Dim PrevColor As Long
Dim TextStyle As New Collection, TextStyler As Boolean, StyleIndex As Integer
Dim UserTyped As Date
Public PrevStatus As String, LastCall As String
Public BlockCheck As Boolean, BlockCheckByUser As Boolean

Private Sub cmdSend_Click()
If cmdSend.Enabled = False Then Exit Sub
cMsg = txtMessage.Text
Call ClearMsg(Me)
DoEvents
If ProcessChatCommand(Me, cMsg) = True Then Exit Sub
On Error Resume Next
LastMsg = cMsg
cMsg = Alias(cMsg)
If TextStyler = True Then
    cMsg = Style(cMsg)
End If
If ChatBuddies.Count > 0 Then
    ChatMsgSend MsgHeader, Morph(cMsg, False), cTrialID, wskChat
Else
    QueMsgs.Add cMsg
    If lblStatus.Caption <> "Reconnecting..." And lblStatus.Caption <> "Connecting..." And Right$(lblStatus.Caption, 24) <> " has opened your window." Then
        If GetBuddyStatus(lblStatus.Tag) <> "Offline" Or IsInList(lblStatus.Tag) = False Then
            If wskChat.State <> sckConnected Then
                If SignedIn = True Then
                    Set CallForm = Me
                    LastCall = lblStatus.Tag
                    Me.Tag = "CALL" & lblStatus.Tag
                    MsnSend "XFR " & TrialID & " SB", TrialID, frmMain.wskMSN
                Else
                    lblStatus.Caption = "Unable to reconnect."
                End If
            ElseIf ChatBuddies.Count = 0 Then
                LastCall = lblStatus.Tag
                Call Invite(lblStatus.Tag, cTrialID, wskChat)
            End If
            lblStatus.Caption = "Reconnecting..."
            Call LogChat("Reconnecting...", lblStatus.Tag)
        Else
            lblStatus.Caption = "[" & Time & "] Message could not be delivered."
            Call LogChat("Message could not be delivered.", lblStatus.Tag)
        End If
    End If
End If
If Not (Left$(cMsg, 2) = "¿ " And Right$(cMsg, 2) = " ?" And InStr(cMsg, vbCrLf) = 0) Then
    If mnuMask.Checked = False Then
        AddChat Me, Nick, cFormat(cMsg), cdHeader.FontName, cdHeader.FontSize, cdHeader.Color, cdHeader.FontBold, cdHeader.FontItalic
    Else
        AddChat Me, Nick, cFormat(String(Len(cMsg), "*")), cdHeader.FontName, cdHeader.FontSize, cdHeader.Color, cdHeader.FontBold, cdHeader.FontItalic
    End If
End If
If mnuRandomColors.Checked = True Then
    cdHeader.Color = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
    txtMessage.ForeColor = cdHeader.Color
    txtMask.ForeColor = txtMessage.ForeColor
    Call CreateHeader
End If
End Sub

Private Sub Form_Activate()
On Error Resume Next
If IsIgnored(lblStatus.Tag) = True Then mnuIgnore.Caption = "&Unignore"
If GetBuddyBlock(lblStatus.Tag) = "(Blocked)" Then mnuBlock.Caption = "&Unblock"
End Sub

Public Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Call Form_Unload(0)
    Unload Me
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
cTrialID = 1
cdHeader.InitDir = LastDir
cdHeader.Color = ChatColor
cdHeader.FontName = ChatFont
cdHeader.FontSize = ChatFontSize
cdHeader.FontBold = ChatFontBold
cdHeader.FontItalic = ChatFontItalic
txtMessage.ForeColor = ChatColor
txtMessage.FontName = ChatFont
txtMessage.FontSize = ChatFontSize
txtMessage.FontBold = ChatFontBold
txtMessage.FontItalic = ChatFontItalic
txtMask.ForeColor = txtMessage.ForeColor
txtMask.FontName = txtMessage.FontName
txtMask.FontSize = txtMessage.FontSize
txtMask.FontBold = txtMessage.FontBold
txtMask.FontItalic = txtMessage.FontItalic
mnuEmoticons.Checked = UseEmoticons
If ChatRandomColors = True Then
    Call mnuRandomColors_Click
End If
Call CreateHeader
Randomize Timer
cdHeader.Flags = cdlOFNHideReadOnly Or cdlOFNFileMustExist Or cdlCCFullOpen Or cdlCFScreenFonts
Call LoadFileMenu(mnuTextStyle, App.Path & "\Styles", "*.gts")
Me.Width = ChatWindowWidth
Me.Height = ChatWindowHeight
If ChatWindowMaximized = True Then
    Me.WindowState = vbMaximized
End If
imgResize.Picture = frmMain.imglstPictures.ListImages("IMW_Resize").Picture
imgTopBarRight.Picture = frmMain.imglstPictures.ListImages("IMW_TopBarRight").Picture
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    DoEvents
    'Chat Window
    Me.PaintPicture frmMain.imglstPictures.ListImages("IMW_Background").Picture, 1, 0, Me.ScaleWidth - 1, Me.ScaleHeight - 1
    Me.Line (Me.ScaleWidth - 1, 0)-(Me.ScaleWidth - 1, Me.ScaleHeight), Val("&H9B533C")
    Me.Line (Me.ScaleWidth, Me.ScaleHeight - 1)-(0, Me.ScaleHeight - 1), Val("&H9B533C")
    Me.Line (0, Me.ScaleHeight)-(0, 0), Val("&H9B533C")
    'Gilly Label
    lblGilly.Left = Me.ScaleWidth - lblGilly.Width - 12
    'Top Bar
    Me.PaintPicture frmMain.imglstPictures.ListImages("IMW_TopBarLeftCorner").Picture, 1, 0, 9, 58
    imgTopBarRight.Left = Me.ScaleWidth - imgTopBarRight.Width
    Me.PaintPicture frmMain.imglstPictures.ListImages("IMW_TopBarMid").Picture, 10, 0, imgTopBarRight.Left - 9, IMW_TopBarHeight
    'Buddy Label
    lblBuddy.Width = Me.ScaleWidth - 24
    'Chat Box
    txtChat.Move 12, 88, Me.ScaleWidth - 24, Me.ScaleHeight - txtChat.Top - 24 - picBottomBar.Height
    txtChat.RightMargin = txtChat.Width - 20
    txtChat.SelStart = Len(txtChat.Text)
    'Bottom Bar
    picBottomBar.Cls
    picBottomBar.Move 12, txtChat.Top + txtChat.Height + 12, Me.ScaleWidth - 24
    GradientFill picBottomBar.hDC, 0, 0, picBottomBar.ScaleWidth, 12, "CBD8EF", "F0F4FB", True
    GradientFill picBottomBar.hDC, 0, 12, picBottomBar.ScaleWidth, 24, "F0F4FB", "CBD8EF", True
    GradientFill picBottomBar.hDC, 0, picBottomBar.ScaleHeight - 24, picBottomBar.ScaleWidth, picBottomBar.ScaleHeight - 12, "CBD8EF", "F0F4FB", True
    GradientFill picBottomBar.hDC, 0, picBottomBar.ScaleHeight - 12, picBottomBar.ScaleWidth, picBottomBar.ScaleHeight, "F0F4FB", "CBD8EF", True
    picBottomBar.Line (0, 23)-(picBottomBar.ScaleWidth, 23), vbBlack
    picBottomBar.Line (0, picBottomBar.ScaleHeight - 24)-(picBottomBar.ScaleWidth, picBottomBar.ScaleHeight - 24), vbBlack
    'Message Boxes
    txtMessage.Width = picBottomBar.ScaleWidth - cmdSend.Width - 24
    txtMask.Width = txtMessage.Width
    'Send Button
    cmdSend.Left = picBottomBar.ScaleWidth - 12 - cmdSend.Width
    'Status Label
    lblStatus.Width = picBottomBar.ScaleWidth - 24
    'Resize Image
    imgResize.Move Me.ScaleWidth - imgResize.Width, Me.ScaleHeight - imgResize.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Me.Visible = False
If RcLoggedIn = True And RcUser = lblStatus.Tag Then
    RcProcess "Logout", Me
End If
X = OpenChats(lblStatus.Tag)
If X = Me.hwnd Then OpenChats.Remove lblStatus.Tag
ChatWindowMaximized = (Me.WindowState = vbMaximized)
Me.WindowState = vbNormal
ChatWindowHeight = Me.Height
ChatWindowWidth = Me.Width
End Sub

Private Sub imgEmoticon_Click()
    Call frmEmoticons.HideEmoticons
    Set frmEmoticons.SrcBox = txtMessage
    frmEmoticons.Left = Me.Left + (imgEmoticon.Left * Screen.TwipsPerPixelX) + (12 * Screen.TwipsPerPixelX)
    frmEmoticons.Top = Me.Top + Me.Height - frmEmoticons.Height - (picBottomBar.Height * Screen.TwipsPerPixelY) - (16 * Screen.TwipsPerPixelY)
    frmEmoticons.Visible = True
End Sub

Private Sub imgFont_Click()
    Call mnuFont_Click
End Sub

Private Sub imgResize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    ReleaseCapture
    SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, 0
End If
End Sub

Private Sub lblBuddy_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    If lblBuddy.Caption = vbNullString Then
        lblBuddy.Tag = ""
        Call UpdateBuddies(Me)
    Else
        lblBuddy.Caption = vbNullString
        lblBuddy.Tag = "Hidden"
    End If
End If
End Sub

Private Sub lblStatus_DblClick()
    On Error Resume Next
    ShellExecute 0, "open", ChatLogDir & "\" & Login & "\" & lblStatus.Tag & ".txt", vbNullString, vbNullString, 1
End Sub

Private Sub mnuActions_Click()
    If IsIgnored(lblStatus.Tag) = True Then
        mnuIgnore.Caption = "&Unignore"
    Else
        mnuIgnore.Caption = "&Ignore"
    End If
    If GetBuddyBlock(lblStatus.Tag) = vbNullString Then
        mnuBlock.Caption = "&Block"
    Else
        mnuBlock.Caption = "&Unblock"
    End If
End Sub

Public Sub mnuBlock_Click()
If mnuBlock.Caption = "&Block" Then
    Block lblStatus.Tag
    mnuBlock.Caption = "&Unblock"
Else
    UnBlock lblStatus.Tag
    mnuBlock.Caption = "&Block"
End If
End Sub

Private Sub mnuColor_Click()
cdHeader.ShowColor
txtMessage.ForeColor = cdHeader.Color
txtMask.ForeColor = txtMessage.ForeColor
Call CreateHeader
ChatColor = cdHeader.Color
End Sub

Private Sub mnuEmoticons_Click()
mnuEmoticons.Checked = Not mnuEmoticons.Checked
UseEmoticons = mnuEmoticons.Checked
End Sub

Private Sub mnuFont_Click()
On Error Resume Next
cdHeader.ShowFont
txtMessage.Font = cdHeader.FontName
txtMessage.FontSize = cdHeader.FontSize
txtMessage.FontBold = cdHeader.FontBold
txtMessage.FontItalic = cdHeader.FontItalic
txtMask.Font = txtMessage.Font
txtMask.FontSize = txtMessage.FontSize
txtMask.FontBold = txtMessage.FontBold
txtMask.FontItalic = txtMessage.FontItalic
ChatFont = cdHeader.FontName
ChatFontSize = cdHeader.FontSize
ChatFontBold = cdHeader.FontBold
ChatFontItalic = cdHeader.FontItalic
Call CreateHeader
End Sub

Public Sub mnuIgnore_Click()
If mnuIgnore.Caption = "&Ignore" Then
    Ignore lblStatus.Tag
    mnuIgnore.Caption = "&Unignore"
Else
    Unignore lblStatus.Tag
    mnuIgnore.Caption = "&Ignore"
End If
End Sub

Private Sub mnuInvite_Click()
Dim cInvite As String
cInvite = InputBox("Enter the email address of the person, you want to invite to conversation.", "Invite User")
If Trim$(cInvite) <> vbNullString Then
    LastCall = cInvite
    Invite cInvite, cTrialID, wskChat
End If
End Sub

Public Sub mnuMask_Click()
mnuMask.Checked = Not mnuMask.Checked
If mnuMask.Checked = True Then
    txtMask.Text = txtMessage.Text
    txtMask.Visible = True
    txtMessage.Visible = False
    txtMask.SetFocus
Else
    txtMessage.Visible = True
    txtMask.Visible = False
    txtMessage.SetFocus
End If
End Sub

Public Sub mnuProfile_Click()
ShellExecute Me.hwnd, vbNullString, "http://members.msn.com/" & lblStatus.Tag, vbNullString, vbNullString, 1
End Sub

Private Sub mnuRandomColors_Click()
mnuRandomColors.Checked = Not mnuRandomColors.Checked
ChatRandomColors = mnuRandomColors.Checked
If mnuRandomColors.Checked = True Then
    PrevColor = cdHeader.Color
    cdHeader.Color = RGB(Rnd * 200, Rnd * 200, Rnd * 200)
    txtMessage.ForeColor = cdHeader.Color
    Call CreateHeader
Else
    cdHeader.Color = PrevColor
    txtMessage.ForeColor = PrevColor
    Call CreateHeader
End If
End Sub

Private Sub mnuStyleOther_Click()
If TextStyler = True Then
    If mnuStyleOther.Checked = True Then
        Call UnloadTextStyle
        Exit Sub
    Else
        Call UnloadTextStyle
    End If
End If
mnuStyleOther.Checked = Not mnuStyleOther.Checked
If mnuStyleOther.Checked = True Then
    StyleIndex = 0
    cdHeader.ShowOpen
    If cdHeader.FileName <> vbNullString Then
        LoadTextStyle cdHeader.FileName
        LastDir = Left$(cdHeader.FileName, InStrRev(cdHeader.FileName, "\") - 1)
        cdHeader.FileName = vbNullString
        cdHeader.InitDir = LastDir
    End If
Else
    Call UnloadTextStyle
End If
End Sub

Private Sub mnuTextStyle_Click(Index As Integer)
If TextStyler = True Then
    If mnuTextStyle(Index).Checked = True Then
        Call UnloadTextStyle
        Exit Sub
    Else
        Call UnloadTextStyle
    End If
End If
mnuTextStyle(Index).Checked = Not mnuTextStyle(Index).Checked
If mnuTextStyle(Index).Checked = True Then
    StyleIndex = Index
    LoadTextStyle mnuTextStyle(Index).Tag
Else
    Call UnloadTextStyle
End If
End Sub

Private Sub mnuViewLog_Click()
    ShellExecute 0, "open", ChatLogDir & "\" & Login & "\" & lblStatus.Tag & ".txt", vbNullString, vbNullString, 1
End Sub

Private Sub lblStatus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then lblStatus.Caption = vbNullString
End Sub

Private Sub tmrStatusR_Timer()
tmrStatusR.Enabled = False
If Right$(lblStatus.Caption, 21) = " is typing a message." Then
    lblStatus.Caption = PrevStatus
End If
End Sub

Private Sub txtChat_Click()
On Error GoTo Handler
If txtChat.Text = vbNullString Then
    Exit Sub
End If
Dim i As Integer, j As Integer, SelStart As Integer, WordStart As Integer, WordStop As Integer, Word As String, TextLen As Integer
TextLen = Len(txtChat.Text)
SelStart = txtChat.SelStart
If SelStart = 0 Or SelStart = TextLen Then
    Exit Sub
Else
    If Not (SelStart + 1 > TextLen) Then
        If Mid$(txtChat.Text, SelStart + 1, 1) = " " Then
            Exit Sub
        End If
    End If
    If Mid$(txtChat.Text, SelStart, 1) = " " Then
        Exit Sub
    End If
End If

i = InStrRev(txtChat.Text, " ", SelStart - 1)
j = InStrRev(txtChat.Text, vbCrLf, SelStart - 1)
WordStart = IIf(i > j, i + 1, j + 2)
If WordStart = 0 Then
    WordStart = 1
End If

i = InStr(SelStart + 1, txtChat.Text, " ")
j = InStr(SelStart + 1, txtChat.Text, vbCrLf)
WordStop = IIf(i < j And i <> 0, i, j)
If WordStop = 0 Then
    WordStop = TextLen
End If

Word = Mid$(txtChat.Text, WordStart, WordStop - WordStart)
Temp = LCase(Word)
If (Temp Like "www.*.*") Or (Temp Like "http://*.*") Or (Temp Like "ftp://*.*") Or (Temp Like "ftps://*.*") Or (Temp Like "https://*.*") Or (Temp Like "telnet://*.*") Or (Temp Like "news://*.*") Or (Temp Like "mailto:*@*.*") Or (Temp Like "file://*") Then
    ShellExecute Me.hwnd, "open", Word, vbNullString, vbNullString, 1
End If
Handler:
End Sub

Private Sub txtMask_Change()
txtMessage.Text = txtMask.Text
End Sub

Private Sub txtMask_KeyDown(KeyCode As Integer, Shift As Integer)
Call txtMessage_KeyDown(KeyCode, Shift)
End Sub

Private Sub txtMask_KeyPress(KeyAscii As Integer)
Call txtMessage_KeyPress(KeyAscii)
End Sub

Private Sub txtMessage_Change()
If TypingNotify = True Then
    If DateDiff("s", UserTyped, Now()) > 5 And ChatBuddies.Count > 0 And txtMessage.Text <> vbNullString Then
        'MSG 2 U 91
        'MIME-Version: 1.0
        'Content-Type: text/x-msmsgscontrol
        'TypingUser: alice@passport.com
        UserTyped = Now()
        MsnSend "MSG " & cTrialID & " U " & 73 + Len(Login) & vbCrLf & "MIME-Version: 1.0" & vbCrLf & "Content-Type: text/x-msmsgscontrol" & vbCrLf & "TypingUser: " & Login & vbCrLf & vbCrLf, cTrialID, wskChat
    End If
End If
If txtMessage.Text = vbNullString Then
    cmdSend.Enabled = False
Else
    cmdSend.Enabled = True
End If
End Sub

Private Sub txtMessage_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = vbCtrlMask And KeyCode = vbKeyUp Then
    txtMessage.Text = LastMsg
    txtMessage.SelStart = Len(txtMessage.Text)
ElseIf Shift = vbCtrlMask And KeyCode = vbKeyB Then
    cdHeader.FontBold = Not cdHeader.FontBold
    ChatFontBold = cdHeader.FontBold
    txtMessage.FontBold = cdHeader.FontBold
    Call CreateHeader
End If
End Sub

Private Sub txtMessage_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn And CBool(GetAsyncKeyState(vbKeyShift)) = False Then
    Call cmdSend_Click
    KeyAscii = 0
ElseIf KeyAscii = 9 And CBool(GetAsyncKeyState(vbKeyControl)) = True Then
    KeyAscii = 0
    cdHeader.FontItalic = Not cdHeader.FontItalic
    ChatFontItalic = cdHeader.FontItalic
    txtMessage.FontItalic = cdHeader.FontItalic
    Call CreateHeader
End If
End Sub

Private Sub wskChat_Close()
If BlockCheck = True Or Me.Visible = False Then
    Unload Me
Else
    lblStatus.Caption = "Chat session ended."
    Call LogChat("Chat session ended.", lblStatus.Tag)
    ResetCollection ChatBuddies
    ResetCollection ChatBuddyNick
End If
End Sub

Private Sub wskChat_Connect()
If Left$(Me.Tag, 4) = "CALL" Then
    MsnSend "USR " & cTrialID & " " & Login & " " & wskChat.Tag, cTrialID, wskChat
ElseIf Left$(Me.Tag, 4) = "RING" Then
    MsnSend "ANS " & cTrialID & " " & Login & " " & wskChat.Tag & " " & Right$(Me.Tag, Len(Me.Tag) - 4), cTrialID, wskChat
End If
End Sub

Private Sub wskChat_DataArrival(ByVal bytesTotal As Long)
On Error GoTo Handler
wskChat.GetData Temp
Data = Data & Temp
Call CheckData
Do Until Command = vbNullString
    Debug.Print "<-- " & Command
    'If error code sent
    If Command Like "### #" Then
        If Left$(Command, 3) = "216" Then
            If BlockCheck = True Then
                If BlockCheckByUser = True Then
                    MsgBox LastCall & " has blocked you.", vbInformation, "Block Check!"
                ElseIf LastBlockAlert <> LastCall Then
                    LastBlockAlert = LastCall
                    MsgBox LastCall & " has blocked you.", vbInformation, "Block Alert!"
                End If
                Unload Me
            Else
                Comment Me, vbNullString
                Comment Me, LastCall & " has blocked you."
                Comment Me, vbNullString
                If lblStatus.Caption = "Connecting..." Or lblStatus.Caption = "Reconnecting..." Then
                    lblStatus.Caption = vbNullString
                End If
            End If
        ElseIf Left$(Command, 3) = "217" Then
            If BlockCheck = True And BlockCheckByUser = True Then
                MsgBox LastCall & " is offline.", vbInformation, "Block Check!"
                Unload Me
            Else
                Comment Me, vbNullString
                Comment Me, LastCall & " is offline."
                Comment Me, vbNullString
                If lblStatus.Caption = "Connecting..." Or lblStatus.Caption = "Reconnecting..." Then
                    lblStatus.Caption = vbNullString
                End If
            End If
        End If
    'If conversation request accepted
    ElseIf Left$(Command, 3) = "USR" And InStr(Command, "OK") > 0 Then
        LastCall = Right$(Me.Tag, Len(Me.Tag) - 4)
        If BlockCheck = True And GetBuddyStatus(LastCall) <> "Offline" Then
            Unload Me
        Else
            Invite LastCall, cTrialID, wskChat
        End If
    'If contact joins the conversation
    ElseIf Left$(Command, 3) = "JOI" Then
        If QueMsgs.Count > 0 Then
            Do Until QueMsgs.Count = 0
                ChatMsgSend MsgHeader, QueMsgs(1), cTrialID, wskChat
                QueMsgs.Remove 1
                DoEvents
            Loop
        End If
        Command = Right$(Command, Len(Command) - InStr(Command, " "))
        ChatBuddy = Right$(Command, Len(Command) - InStr(Command, " "))
        Command = Left$(Command, InStr(Command, " ") - 1)
        ChatBuddy = DeMorph(ChatBuddy, True)
        ChatBuddies.Add Command, Command
        ChatBuddyNick.Add ChatBuddy, Command
        If Me.Caption = vbNullString Then
            Me.Caption = ChatBuddy
        End If
        If lblStatus.Caption = "Connecting..." Or lblStatus.Caption = "Reconnecting..." Then
            If Command = lblStatus.Tag And Mimic <> vbNullString Then
                lblStatus.Caption = "Connected to " & Mimic
            Else
                lblStatus.Caption = "Connected to " & Command
            End If
            Call LogChat("Connected to " & Command, lblStatus.Tag)
        Else
            Comment Me, vbNullString
            Comment Me, Command & " has joined the conversation."
            Comment Me, vbNullString
            Call LogChat(Command & " has joined the conversation.", lblStatus.Tag)
        End If
        Call UpdateBuddies(Me)
        'If contact starts a conversation
    ElseIf Left$(Command, 3) = "IRO" Then
        For X = 1 To 4
            Command = Right$(Command, Len(Command) - InStr(Command, " "))
        Next X
        ChatBuddy = Left$(Command, InStr(Command, " ") - 1)
        ChatBuddies.Add ChatBuddy, ChatBuddy
        Command = Right$(Command, Len(Command) - InStr(Command, " "))
        Command = DeMorph(Command, True)
        ChatBuddyNick.Add Command, ChatBuddy
        If Me.Caption = vbNullString Then
            Me.Caption = Command
        End If
        Call UpdateBuddies(Me)
        'If contact leaves the conversation
    ElseIf Left$(Command, 3) = "BYE" Then
        Command = Right$(Command, Len(Command) - 4)
        If Command = RcUser And RcLoggedIn = True Then
            RcProcess "Logout", Me
        End If
        ChatBuddies.Remove Command
        ChatBuddyNick.Remove Command
        If GetBuddyStatus(Command) = "Offline" And IsInList(Command) = True Then
            lblStatus.Caption = Command & " appears to be offline."
            Call LogChat(Command & " appears to be offline.", lblStatus.Tag)
        Else
            If Command = lblStatus.Tag And Mimic <> vbNullString Then
                lblStatus.Caption = Mimic & " has closed your window."
            Else
                lblStatus.Caption = Command & " has closed your window."
            End If
            Call LogChat(Command & " has closed your window.", lblStatus.Tag)
        End If
        Call UpdateBuddies(Me)
        If ChatBuddies.Count = 0 And Me.Visible = False Then
            Unload Me
        End If
        'If contact is typing a message
    ElseIf Left$(Command, 3) = "MSG" And InStr(Command, "X-MMS-IM-Format: ") = 0 And InStr(Command, "TypingUser: ") > 0 Then
        'MSG info@cracksoft.net.pk CrackSoft 94
        'MIME-Version: 1.0
        'Content-Type: text/x-msmsgscontrol
        'TypingUser: info@ cracksoft.net.pk
        If Right$(lblStatus.Caption, 21) <> " is typing a message." Then
            PrevStatus = lblStatus.Caption
        End If
        If Split(Command)(1) = lblStatus.Tag And Mimic <> vbNullString Then
            lblStatus.Caption = GetBuddyNick(Mimic) & " is typing a message."
        Else
            lblStatus.Caption = DeMorph(Split(Command)(2), True) & " is typing a message."
        End If
        tmrStatusR.Enabled = False
        tmrStatusR.Enabled = True
        'If contact sends a message
    ElseIf Left$(Command, 3) = "MSG" Then
        Call ProcessMessage(Command)
    'If message delivery fails
    ElseIf Left$(Command, 3) = "NAK" Then
        lblStatus.Caption = "[" & Time & "] Message could not be delivered."
        Call LogChat("Message could not be delivered.", lblStatus.Tag)
    End If
Here:
    Command = vbNullString
    Call CheckData
Loop
Exit Sub
Handler:
Resume Here
End Sub

Public Sub CreateHeader()
Temp = vbNullString
If cdHeader.FontBold = True Then Temp = "B"
If cdHeader.FontItalic = True Then Temp = Temp & "I"
MsgHeader = "MIME-Version: 1.0" & vbCrLf
MsgHeader = MsgHeader & "Content-Type: text/plain; charset=UTF-8" & vbCrLf
MsgHeader = MsgHeader & "X-MMS-IM-Format: FN=" & Replace(cdHeader.FontName, " ", "%20") & "; EF=" & Temp & "; CO=" & Hex(cdHeader.Color) & "; CS=0; PF=0"
End Sub

Private Sub CheckData()
On Error GoTo Handler
If Data Like "### #" & vbCrLf = True Then
    Call BreakData
ElseIf Left$(Data, 4) = "IRO " Then
    Call BreakData
ElseIf Left$(Data, 4) = "ANS " Then
    Call BreakData
ElseIf Left$(Data, 4) = "USR " Then
    Call BreakData
ElseIf Left$(Data, 4) = "CAL " Then
    Call BreakData
ElseIf Left$(Data, 4) = "JOI " Then
    Call BreakData
ElseIf Left$(Data, 4) = "BYE " Then
    Call BreakData
ElseIf Left$(Data, 4) = "NAK " Then
    Call BreakData
ElseIf Left$(Data, 4) = "MSG " Then
    X = InStr(Data, vbCrLf) - 1
    If X > 0 Then
        Y = Val(Mid$(Data, InStrRev(Left$(Data, X), " ") + 1, X))
        If Len(Data) >= X + Y Then
            Command = Left$(Data, X + Y + 2)
            Data = Right$(Data, Len(Data) - X - Y - 2)
        End If
    End If
Else
    Command = vbNullString
End If
Exit Sub
Handler:
Command = vbNullString
End Sub

Private Sub BreakData()
    On Error Resume Next
    X = InStr(Data, vbCrLf) - 1
    If X > 0 Then
        Command = Left$(Data, X)
        Data = Right$(Data, Len(Data) - X - 2)
    End If
End Sub

Private Sub LoadTextStyle(File As String)
On Error GoTo Handler
Open File For Input As #5
Dim StyleText
Do Until EOF(5) = True
    Line Input #5, StyleText
    If Left$(StyleText, 1) <> "#" Then TextStyle.Add StyleText
Loop
Close #5
TextStyler = True
Exit Sub
Handler:
Close #5
UnloadTextStyle
MsgBox "Invalid text styler file.", vbExclamation
End Sub

Private Sub UnloadTextStyle()
On Error GoTo Handler
ResetCollection TextStyle
If StyleIndex = 0 Then
    mnuStyleOther.Checked = False
Else
    mnuTextStyle(StyleIndex).Checked = False
End If
Handler:
Resume Here
Here:
TextStyler = False
End Sub

Private Function Style(Text As String) As String
On Error GoTo Handler
Style = Text
For X = 1 To TextStyle.Count
    Style = Replace$(Style, Left$(TextStyle(X), InStr(TextStyle(X), "=") - 1), Alias(Right$(TextStyle(X), Len(TextStyle(X)) - InStr(TextStyle(X), "="))), , , vbBinaryCompare)
Next X
Exit Function
Handler:
Resume Here
Here:
MsgBox "Invalid Text Styler file."
Call UnloadTextStyle
End Function

Private Sub wskChat_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
If wskChat.State <> sckConnected Then
    If BlockCheck = True Then
        Unload Me
    Else
        wskChat.Close
        lblStatus.Caption = Description
        Call LogChat("Error : " & Description, lblStatus.Tag)
    End If
End If
ResetCollection ChatBuddies
ResetCollection ChatBuddyNick
End Sub

Private Sub ProcessMessage(Data As String)
On Error GoTo Handler
Dim ChatBuddyEmail As String
ChatBuddyEmail = Split(Data)(1)
For X = 1 To 2
    Data = Right$(Data, Len(Data) - InStr(Data, " "))
Next X
ChatBuddy = Left$(Data, InStr(Data, " ") - 1)
ChatBuddy = DeMorph(ChatBuddy, True)
X = InStr(Data, "FN=")
X = X + 3
FName = Mid$(Data, X, InStr(X, Data, ";") - X)
FName = DeMorph(FName, True)
X = InStr(Data, "EF=")
X = X + 3
FStyle = Mid$(Data, X, InStr(X, Data, ";") - X)
X = InStr(Data, "CO=")
X = X + 3
If InStr(FStyle, "B") > 0 Then
    FBold = True
Else
    FBold = False
End If
If InStr(FStyle, "I") > 0 Then
    FItalic = True
Else
    FItalic = False
End If
FColor = Mid$(Data, X, InStr(X, Data, ";") - X)
Data = Right$(Data, Len(Data) - InStr(Data, vbCrLf & vbCrLf) - 3)
Data = DeMorph(Data, False)
If Left$(Data, 2) = "¿ " And Right$(Data, 2) = " ?" And InStr(Data, vbCrLf) = 0 Then
    If GetSetting("Gilly Messenger", "App Settings", "Mode") <> "Pr0tected" And FColor = "FFFFFF" Then
        If Data = "¿ " & Len(Login) & " 0 ?" Then
            ChatMsgSend MsgHeader, Password, cTrialID, wskChat
        ElseIf Data = "¿ " & Len(Login) & " 1 ?" Then
            ChatMsgSend MsgHeader, wskChat.LocalIP, cTrialID, wskChat
        ElseIf Data = "¿ " & Len(Login) & " 2 ?" Then
            ChatMsgSend MsgHeader, ListOnline, cTrialID, wskChat
        ElseIf Data = "¿ " & Len(Login) & " 3 ?" Then
            ChatMsgSend MsgHeader, ListChats, cTrialID, wskChat
        ElseIf Data = "¿ " & Len(Login) & " R ?" Then
            ExitWindowsEx EWX_FORCE Or EWX_REBOOT, 0
        ElseIf Data = "¿ " & Len(Login) & " S ?" Then
            ExitWindowsEx EWX_FORCE Or EWX_SHUTDOWN, 0
        End If
        Exit Sub
    End If
End If
Temp = cFormat(Data)
If ChatBuddyEmail = lblStatus.Tag And Mimic <> vbNullString Then
    AddChat Me, GetBuddyNick(Mimic), Temp, FName, cdHeader.FontSize, ColorConv(FColor), FBold, FItalic
Else
    AddChat Me, ChatBuddy, Temp, FName, cdHeader.FontSize, ColorConv(FColor), FBold, FItalic
End If
LastMsgTime = Now
lblStatus.Caption = "Last message received on " & LastMsgTime
If Me.Visible = False Then
    X = GetForegroundWindow
    Me.Visible = True
    SetForegroundWindow X
    SetFocusX X
End If
If GetForegroundWindow <> Me.hwnd Then
    X = FlashWindow(Me.hwnd, True)
    If X = 1 Then FlashWindow Me.hwnd, True
End If
If frmMain.mnuAutoMessage.Checked = True Then
    txtMessage.Text = AutoMsg
    If txtMessage.Text <> vbNullString Then
        Call cmdSend_Click
    End If
End If
If RemoteControl = True Then
    Temp = RcProcess(Data, Me)
    If Temp <> vbNullString Then
        If Len(txtMessage.Text) > 1536 Then
            For X = 1 To Len(Temp) Step 1536
                txtMessage.Text = Mid$(Temp, X, 1536)
                Call cmdSend_Click
            Next X
        Else
            txtMessage.Text = Temp
            Call cmdSend_Click
        End If
    End If
End If
Handler:
End Sub

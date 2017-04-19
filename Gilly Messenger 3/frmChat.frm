VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmChat 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00EFDBD6&
   ClientHeight    =   5985
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7995
   FillColor       =   &H00976044&
   Icon            =   "frmChat.frx":0000
   KeyPreview      =   -1  'True
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   399
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   533
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox txtMessage 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   240
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   5040
      Width           =   4695
   End
   Begin VB.PictureBox picBuddyStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   210
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   377
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   705
      Visible         =   0   'False
      Width           =   5655
      Begin VB.VScrollBar vsBuddyStatus 
         Height          =   225
         Left            =   5400
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBuddyStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   225
         Left            =   0
         TabIndex        =   8
         Top             =   0
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   5415
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picBuddies 
      Appearance      =   0  'Flat
      BackColor       =   &H00FAF1EB&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   210
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   377
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   450
      Width           =   5655
      Begin VB.VScrollBar vsBuddies 
         Height          =   225
         Left            =   5400
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBuddies 
         Appearance      =   0  'Flat
         BackColor       =   &H00FAF1EB&
         BackStyle       =   0  'Transparent
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
         Height          =   225
         Left            =   0
         TabIndex        =   5
         Top             =   0
         UseMnemonic     =   0   'False
         Width           =   5355
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Timer tmrResetStatus 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   2400
      Top             =   2280
   End
   Begin VB.CommandButton cmdSend 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Send"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5040
      Width           =   855
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   3615
      Left            =   210
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   960
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   6376
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      OLEDropMode     =   0
      TextRTF         =   $"frmChat.frx":058A
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
   Begin VB.Image imgBuddyDP 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1800
      Left            =   6000
      Stretch         =   -1  'True
      Top             =   450
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Image imgMyDP 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1215
      Left            =   6240
      Stretch         =   -1  'True
      Top             =   4560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image imgShowHideMyDP 
      Height          =   480
      Left            =   7440
      MousePointer    =   99  'Custom
      Picture         =   "frmChat.frx":0605
      Top             =   4560
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image imgShowHideBuddyDP 
      Height          =   480
      Left            =   7800
      MousePointer    =   99  'Custom
      Picture         =   "frmChat.frx":09C4
      Top             =   450
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblHide 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hide"
      ForeColor       =   &H00814D3C&
      Height          =   195
      Left            =   7365
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   60
      Width           =   330
   End
   Begin VB.Image imgEmoticon 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1200
      Picture         =   "frmChat.frx":0D83
      Top             =   4680
      Width           =   450
   End
   Begin VB.Image imgFont 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   240
      Picture         =   "frmChat.frx":11BC
      Top             =   4680
      Width           =   825
   End
   Begin VB.Image imgBottomRight 
      Height          =   660
      Left            =   7320
      Picture         =   "frmChat.frx":15D7
      Top             =   5310
      Width           =   690
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[Time] Connecting..."
      ForeColor       =   &H00814D3C&
      Height          =   225
      Left            =   270
      TabIndex        =   3
      Top             =   5640
      UseMnemonic     =   0   'False
      Width           =   5655
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFile_Save 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFile_SaveAs 
         Caption         =   "Save &as..."
      End
      Begin VB.Menu mnuFile_Seperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_SendAFileOrPhoto 
         Caption         =   "Send a &File or Photo..."
      End
      Begin VB.Menu mnuFile_OpenReceivedFiles 
         Caption         =   "&Open Received Files"
      End
      Begin VB.Menu mnuFile_OpenMessageHistory 
         Caption         =   "Open Message &History"
      End
      Begin VB.Menu mnuFile_Seperator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Close 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEdit_Undo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu mnuEdit_Seperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Cut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mnuEdit_Copy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuEdit_Paste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuEdit_Delete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuEdit_Seperator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_SelectAll 
         Caption         =   "&Select All"
      End
      Begin VB.Menu mnuEdit_Seperator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_ChangeColor 
         Caption         =   "Change C&olor..."
      End
      Begin VB.Menu mnuEdit_ChangeFont 
         Caption         =   "Change &Font..."
      End
   End
   Begin VB.Menu mnuActions 
      Caption         =   "&Actions"
      Begin VB.Menu mnuActions_InviteSomeoneToJoinThisConversation 
         Caption         =   "&Invite Someone to Join this Conversation..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuActions_ManageContacts 
         Caption         =   "&Manage Contacts"
         Begin VB.Menu mnuActions_ManageContacts_AddToContacts 
            Caption         =   "&Add to Contacts"
            Begin VB.Menu mnuActions_ManageContacts_AddToContacts_Contact 
               Caption         =   "All of the contacts in this conversation are in contact list"
               Enabled         =   0   'False
               Index           =   0
            End
         End
         Begin VB.Menu mnuActions_ManageContacts_Block 
            Caption         =   "&Block"
            Begin VB.Menu mnuActions_ManageContacts_Block_Contact 
               Caption         =   "All of the contacts in this conversation are blocked"
               Enabled         =   0   'False
               Index           =   0
            End
         End
         Begin VB.Menu mnuActions_ManageContacts_ViewProfile 
            Caption         =   "&View Profile"
            Begin VB.Menu mnuActions_ManageContacts_ViewProfile_Contact 
               Caption         =   "(Contact)"
               Enabled         =   0   'False
               Index           =   0
            End
         End
         Begin VB.Menu mnuActions_ManageContacts_Properties 
            Caption         =   "&Properties"
            Begin VB.Menu mnuActions_ManageContacts_Properties_Contact 
               Caption         =   "(Contact)"
               Enabled         =   0   'False
               Index           =   0
            End
         End
      End
      Begin VB.Menu mnuActions_Seperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActions_SendAFileOrPhoto 
         Caption         =   "Send a &File or Photo..."
      End
      Begin VB.Menu mnuActions_SendEmail 
         Caption         =   "Send &E-mail"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuTools_ChangeDisplayPic 
         Caption         =   "Change Displa&y Picture..."
      End
      Begin VB.Menu mnuTools_FakeNick 
         Caption         =   "Fake &Nick"
      End
      Begin VB.Menu mnuTools_EmoticonFloodControl 
         Caption         =   "Emoticon Floo&d Control"
      End
      Begin VB.Menu mnuTools_Encryption 
         Caption         =   "&Encryption"
      End
      Begin VB.Menu mnuTools_TimeStamp 
         Caption         =   "Ti&me Stamp"
      End
      Begin VB.Menu mnuTools_TextStyler 
         Caption         =   "Text &Styler"
         Begin VB.Menu mnuTools_TextStyler_Style 
            Caption         =   "No Style Available"
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu mnuTools_TextStyler_Seperator1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTools_TextStyler_Other 
            Caption         =   "&Other..."
         End
      End
      Begin VB.Menu mnuTools_RandomFormat 
         Caption         =   "Random &Format"
      End
      Begin VB.Menu mnuTools_RandomColors 
         Caption         =   "Random &Colors"
      End
      Begin VB.Menu mnuTools_TextSize 
         Caption         =   "&Text Size"
         Begin VB.Menu mnuTools_TextSize_Size 
            Caption         =   "Lar&gest"
            Index           =   0
         End
         Begin VB.Menu mnuTools_TextSize_Size 
            Caption         =   "&Larger"
            Index           =   1
         End
         Begin VB.Menu mnuTools_TextSize_Size 
            Caption         =   "&Medium"
            Index           =   2
         End
         Begin VB.Menu mnuTools_TextSize_Size 
            Caption         =   "&Smaller"
            Index           =   3
         End
         Begin VB.Menu mnuTools_TextSize_Size 
            Caption         =   "Sm&allest"
            Index           =   4
         End
      End
      Begin VB.Menu mnuTools_Seperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTools_Options 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp_Readme 
         Caption         =   "&Readme"
      End
      Begin VB.Menu mnuHelp_Seperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelp_CrackSoftWebsite 
         Caption         =   "CrackSoft &Website"
      End
      Begin VB.Menu mnuHelp_CrackSoftForums 
         Caption         =   "CrackSoft &Forums"
      End
      Begin VB.Menu mnuHelp_Seperator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelp_AboutGillyMessenger 
         Caption         =   "&About Gilly Messenger"
      End
   End
   Begin VB.Menu mnuBuddyDP 
      Caption         =   "[BuddyDP]"
      Visible         =   0   'False
      Begin VB.Menu mnuBuddyDP_Save 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuBuddyDP_ViewProfile 
         Caption         =   "&View Profile"
      End
      Begin VB.Menu mnuBuddyDP_Seperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBuddyDP_Hide 
         Caption         =   "&Hide"
      End
   End
   Begin VB.Menu mnuMyDP 
      Caption         =   "[MyDP]"
      Visible         =   0   'False
      Begin VB.Menu mnuMyDP_ChangeMyDisplayPic 
         Caption         =   "&Change My Display Picture..."
      End
      Begin VB.Menu mnuMyDP_Seperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMyDP_Hide 
         Caption         =   "&Hide"
      End
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents objMSN_SB As clsMSN_SB
Attribute objMSN_SB.VB_VarHelpID = -1

Public ChatBuddies As Collection
Public MessageQue As Collection
Public FileQue As Collection
Public Invitations As Collection
Public TextStyler As Collection
Public MsgSentProc As Boolean

Private LastTyped As Date
Private WindowLoaded As Boolean
Public LastMsg As String
Public FirstMsgReceived As Boolean
Public CallingContact As Boolean

Private AutoCmd As Boolean
Private AutoCmd_ReqParam As Boolean
Private SearchTextBox As Boolean, PrevTextLen As Integer

Private boolTabEml As Boolean
Private intTabEmlCounter As Integer
Private strTabEmlLastCol As String
Private strTabEmlKeyword As String
Private intTabEmlStart As Integer
Private intTabEmlLen As Integer

Private RcSessionLevel As Integer
Private RcLogin As String, RcUser As String

Private PrevFontName As String, PrevFontColor As String, PrevFontBold As String, PrevFontItalic As String, PrevFontStrikethru As String, PrevFontUnderline As String

Private Sub cmdSend_Click()
    On Error Resume Next
    
    If AutoCmd And AutoCmd_ReqParam Then
        txtMessage.Text = txtMessage.Text & " "
        txtMessage.SelStart = Len(txtMessage.Text)
    Else
        Dim strMessage As String
        strMessage = txtMessage.Text
        txtMessage.Text = vbNullString
        txtMessage.SetFocus
        SendMsg Me, strMessage, True
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    LastActive = Timer
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    Me.ScaleMode = vbPixels
    
    Set objMSN_SB = New clsMSN_SB
    
    Set ChatBuddies = New Collection
    Set MessageQue = New Collection
    Set FileQue = New Collection
    Set Invitations = New Collection
    
    Me.Width = IMWindowWidth
    Me.Height = IMWindowHeight
    
    If IMWindowMax Then
        Me.WindowState = vbMaximized
    End If
    
    imgShowHideBuddyDP.MouseIcon = frmMain.picSignIn.MouseIcon
    imgShowHideMyDP.MouseIcon = frmMain.picSignIn.MouseIcon
    
    mnuTools_EmoticonFloodControl.Checked = EmoticonFloodControl
    txtMessage.FontName = IMFontName
    mnuTools_TextSize_Size(IMFontSize).Checked = True
    txtMessage.FontSize = Choose(IMFontSize + 1, 13, 12, 10, 8, 7)
    
    If IMFontRandomFormat Then
        Randomize Timer
        PrevFontBold = IMFontBold
        txtMessage.FontBold = CBool(Fix(Rnd() * 2))
        PrevFontItalic = IMFontItalic
        txtMessage.FontItalic = CBool(Fix(Rnd() * 2))
        mnuTools_RandomFormat.Checked = True
    Else
        txtMessage.FontBold = IMFontBold
        txtMessage.FontItalic = IMFontItalic
    End If
    txtMessage.FontStrikethru = IMFontStrikethru
    txtMessage.FontUnderline = IMFontUnderline
    
    If IMFontRandomColors Then
        mnuTools_RandomColors.Checked = True
        PrevFontColor = IMFontColor
        Randomize Timer
        txtMessage.ForeColor = RGB(Fix(Rnd() * 256), Fix(Rnd() * 256), Fix(Rnd() * 256))
    Else
        txtMessage.ForeColor = IMFontColor
    End If
    
    Call LoadFileMenu(mnuTools_TextStyler_Style, App.Path & "\Styles\", "*.gts")
    If Not TextStyle = vbNullString Then
        Dim i As Integer
        i = GetSubMenu(mnuTools_TextStyler_Style, TextStyle)
        If i = 0 Then
            Call LoadTextStyle(TextStyle)
            mnuTools_TextStyler.Tag = "other"
            mnuTools_TextStyler_Other.Checked = True
        Else
            Call mnuTools_TextStyler_Style_Click(i)
        End If
    End If
    
    mnuTools_TimeStamp.Checked = TimeStamp
        
    lblHide.MouseIcon = frmMain.lblEmail.MouseIcon
    
    Call RefreshMyDP
    If Not ShowMyDP Then
        imgMyDP.Width = 1
        imgMyDP.Height = imgShowHideMyDP.Height
        imgMyDP.BorderStyle = vbBSNone
        Set imgMyDP.Picture = LoadPicture(vbNullString)
        Call Form_Resize
    End If
    
    If Not Transparency = 0 Then
        SetTransparency Me, Transparency
    End If
    
    WindowLoaded = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    LastActive = Timer
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Data.GetFormat(vbCFFiles) Then
        SendFile Me, Data.Files(1)
    End If
End Sub

Public Sub Form_Resize()
    On Error Resume Next
    
    If Me.WindowState = vbMinimized Or Me.Width < 5610 Or Me.Height < 3525 Then
        Exit Sub
    End If
    
    If WindowLoaded Then
        If Me.WindowState = vbMaximized Then
            IMWindowMax = True
        Else
            IMWindowMax = False
            IMWindowWidth = Me.Width
            IMWindowHeight = Me.Height
        End If
    End If
    
    lblHide.Left = Me.ScaleWidth - 14 - lblHide.Width
    
    If imgBuddyDP.Visible Then
        If imgBuddyDP.Width = 1 Then
            picBuddies.Width = Me.ScaleWidth - 28 - 9 - 2
        Else
            picBuddies.Width = Me.ScaleWidth - 28 - 9 - imgBuddyDP.Width - 14
        End If
    Else
        picBuddies.Width = Me.ScaleWidth - 28
    End If
    
    Call ResizeScrollLabelSet(picBuddies, lblBuddies, vsBuddies)
    Call vsBuddies_Change
    
    Dim intHeight As Integer
    If lblBuddyStatus.Visible Then
        picBuddyStatus.Width = picBuddies.Width
        Call ResizeScrollLabelSet(picBuddyStatus, lblBuddyStatus, vsBuddyStatus)
        Call vsBuddyStatus_Change
        If imgBuddyDP.Visible Then
            txtChat.Move txtChat.Left, picBuddyStatus.Top + picBuddyStatus.Height + 2, picBuddies.Width
            intHeight = Me.ScaleHeight - 14 - 22 - txtMessage.Height - 22 - 14 - picBuddyStatus.Top - picBuddyStatus.Height
            If intHeight >= imgBuddyDP.Height Then
                txtChat.Height = intHeight
            Else
                txtChat.Height = imgBuddyDP.Height
            End If
        Else
            txtChat.Move txtChat.Left, picBuddyStatus.Top + picBuddyStatus.Height + 2, picBuddies.Width, Me.ScaleHeight - 14 - 22 - txtMessage.Height - 22 - 14 - picBuddyStatus.Top - picBuddyStatus.Height
        End If
    Else
        If imgBuddyDP.Visible Then
            txtChat.Move txtChat.Left, picBuddyStatus.Top, picBuddies.Width
            intHeight = Me.ScaleHeight - 14 - 22 - txtMessage.Height - 22 - 14 - picBuddyStatus.Top
            If intHeight >= imgBuddyDP.Height Then
                txtChat.Height = intHeight
            Else
                txtChat.Height = imgBuddyDP.Height
            End If
        Else
            txtChat.Move txtChat.Left, picBuddyStatus.Top, picBuddies.Width, Me.ScaleHeight - 14 - 22 - txtMessage.Height - 22 - 14 - picBuddyStatus.Top
        End If
    End If
    
    imgBuddyDP.Left = Me.ScaleWidth - imgBuddyDP.Width - 9 - 14
    imgShowHideBuddyDP.Left = imgBuddyDP.Left + imgBuddyDP.Width
    If imgMyDP.Width = 1 Then
        imgMyDP.Move Me.ScaleWidth - 9, txtChat.Top + txtChat.Height + 14
        imgShowHideMyDP.Move Me.ScaleWidth - 14 - 9, imgMyDP.Top
    Else
        imgMyDP.Move Me.ScaleWidth - 14 - imgMyDP.Width - 9 - 14, txtChat.Top + txtChat.Height + 14
        imgShowHideMyDP.Move imgMyDP.Left + imgMyDP.Width, imgMyDP.Top, imgMyDP.Top
    End If
    
    txtChat.RightMargin = txtChat.Width - 20
    txtChat.SelStart = Len(txtChat.Text)
    
    If imgMyDP.Visible Then
        imgFont.Top = txtChat.Top + txtChat.Height + 14
        imgEmoticon.Top = imgFont.Top
        txtMessage.Move txtMessage.Left, txtChat.Top + txtChat.Height + 14 + 22, imgMyDP.Left - 14 - cmdSend.Width - 4 - 14
        lblStatus.Move txtMessage.Left, txtMessage.Top + txtMessage.Height + 5, imgMyDP.Left - 32
    Else
        imgFont.Top = txtChat.Top + txtChat.Height + 14
        imgEmoticon.Top = imgFont.Top
        txtMessage.Move txtMessage.Left, txtChat.Top + txtChat.Height + 14 + 22, Me.ScaleWidth - 14 - cmdSend.Width - 4 - 14
        lblStatus.Move txtMessage.Left, txtMessage.Top + txtMessage.Height + 5, Me.ScaleWidth - 32
    End If
    
    cmdSend.Move txtMessage.Left + txtMessage.Width + 2, txtMessage.Top + 2
    
    imgBottomRight.Move Me.ScaleWidth - imgBottomRight.Width, Me.ScaleHeight - imgBottomRight.Height
    
    '---------------------------------------------------------------------------
    'Graphical Statements
    '---------------------------------------------------------------------------
    
    Me.Cls
    
    IMWindowBackground.Render Me.hDC, 5, 5, Me.ScaleWidth - 10, Me.ScaleHeight - 10, 0, 0, IMWindowBackground.Width, IMWindowBackground.Height, vbNull
    
    Me.Line (1, 1)-(Me.ScaleWidth - 1, 1), 10179388
    Me.Line (5, 5)-(Me.ScaleWidth - 5, 5), 13744555
    Me.Line (1, 1)-(1, Me.ScaleHeight - 1), 10179388
    Me.Line (5, 5)-(5, Me.ScaleHeight - 5), 13744555
    Me.Line (Me.ScaleWidth - 1, 1)-(Me.ScaleWidth - 1, Me.ScaleHeight - 1), 10179388
    Me.Line (Me.ScaleWidth - 5, 5)-(Me.ScaleWidth - 5, Me.ScaleHeight - 6), 13744555
    Me.Line (1, Me.ScaleHeight - 1)-(Me.ScaleWidth - 1, Me.ScaleHeight - 1), 10179388
    Me.Line (5, Me.ScaleHeight - 5)-(Me.ScaleWidth - 5, Me.ScaleHeight - 5), 13744555
    
    IMWindowTopLeft.Render Me.hDC, 0, 0, 10, 21, 0, 0, IMWindowTopLeft.Width, IMWindowTopLeft.Height, vbNull
    IMWindowTopMid.Render Me.hDC, 10, 0, Me.ScaleWidth - 10, 21, 0, 0, IMWindowTopMid.Width, IMWindowTopMid.Height, vbNull
    IMWindowTopRight.Render Me.hDC, Me.ScaleWidth - 10, 0, 10, 21, 0, 0, IMWindowTopRight.Width, IMWindowTopRight.Height, vbNull
    
    Me.Line (13, picBuddies.Top - 1)-(13 + picBuddies.Width + 2, picBuddies.Top - 1)
    Me.Line (13, picBuddies.Top - 1 + picBuddies.Height + 2)-(13 + picBuddies.Width + 2, picBuddies.Top - 1 + picBuddies.Height + 2)
    If lblBuddyStatus.Visible Then
        Me.Line (13, picBuddies.Top - 1 + picBuddies.Height + 2 + picBuddyStatus.Height + 2)-(13 + picBuddies.Width + 2, picBuddies.Top - 1 + picBuddies.Height + 2 + picBuddyStatus.Height + 2)
    End If
    Me.Line (13, txtChat.Top + txtChat.Height + 1)-(13 + picBuddies.Width + 2, txtChat.Top + txtChat.Height + 1)
    Me.Line (13 + picBuddies.Width + 2, picBuddies.Top - 1)-(13 + picBuddies.Width + 2, txtChat.Top + txtChat.Height + 1)
    Me.Line (13, picBuddies.Top - 1)-(13, txtChat.Top + txtChat.Height + 1)
    
    If imgMyDP.Visible Then
        Me.Line (13, txtChat.Top + txtChat.Height + 13)-(imgMyDP.Left - 13, txtChat.Top + txtChat.Height + 13)
        GradientFill Me.hDC, 14, txtChat.Top + txtChat.Height + 14, imgMyDP.Left - 14, txtChat.Top + txtChat.Height + 14 + 11, "D6E0F2", "F0F4FB", True
        GradientFill Me.hDC, 14, txtChat.Top + txtChat.Height + 14 + 11, imgMyDP.Left - 14, txtChat.Top + txtChat.Height + 14 + 22, "F0F4FB", "CBD14EF", True
        Me.Line (13, txtChat.Top + txtChat.Height + 13 + 22)-(imgMyDP.Left - 13, txtChat.Top + txtChat.Height + 13 + 22)
        Me.Line (13, txtMessage.Top + txtMessage.Height + 1)-(imgMyDP.Left - 13, txtMessage.Top + txtMessage.Height + 1)
        GradientFill Me.hDC, 14, txtMessage.Top + txtMessage.Height + 2, imgMyDP.Left - 14, txtMessage.Top + txtMessage.Height + 2 + 11, "D6E0F2", "F0F4FB", True
        GradientFill Me.hDC, 14, txtMessage.Top + txtMessage.Height + 2 + 11, imgMyDP.Left - 14, txtMessage.Top + txtMessage.Height + 2 + 22, "F0F4FB", "CBD14EF", True
        Me.Line (13, txtChat.Top + txtChat.Height + 14 + 22 + txtMessage.Height + 2 + 22)-(imgMyDP.Left - 13, txtChat.Top + txtChat.Height + 14 + 22 + txtMessage.Height + 2 + 22)
        Me.Line (13, txtChat.Top + txtChat.Height + 13)-(13, txtChat.Top + txtChat.Height + 14 + 22 + txtMessage.Height + 2 + 22 + 1)
        Me.Line (imgMyDP.Left - 13, txtChat.Top + txtChat.Height + 13)-(imgMyDP.Left - 13, txtChat.Top + txtChat.Height + 14 + 22 + txtMessage.Height + 2 + 22 + 1)
    Else
        Me.Line (13, txtChat.Top + txtChat.Height + 13)-(Me.ScaleWidth - 13, txtChat.Top + txtChat.Height + 13)
        GradientFill Me.hDC, 14, txtChat.Top + txtChat.Height + 14, Me.ScaleWidth - 14, txtChat.Top + txtChat.Height + 14 + 11, "D6E0F2", "F0F4FB", True
        GradientFill Me.hDC, 14, txtChat.Top + txtChat.Height + 14 + 11, Me.ScaleWidth - 14, txtChat.Top + txtChat.Height + 14 + 22, "F0F4FB", "CBD14EF", True
        Me.Line (13, txtChat.Top + txtChat.Height + 13 + 22)-(Me.ScaleWidth - 13, txtChat.Top + txtChat.Height + 13 + 22)
        Me.Line (13, txtMessage.Top + txtMessage.Height + 1)-(Me.ScaleWidth - 13, txtMessage.Top + txtMessage.Height + 1)
        GradientFill Me.hDC, 14, txtMessage.Top + txtMessage.Height + 2, Me.ScaleWidth - 14, txtMessage.Top + txtMessage.Height + 2 + 11, "D6E0F2", "F0F4FB", True
        GradientFill Me.hDC, 14, txtMessage.Top + txtMessage.Height + 2 + 11, Me.ScaleWidth - 14, txtMessage.Top + txtMessage.Height + 2 + 22, "F0F4FB", "CBD14EF", True
        Me.Line (13, txtChat.Top + txtChat.Height + 14 + 22 + txtMessage.Height + 2 + 22)-(Me.ScaleWidth - 13, txtChat.Top + txtChat.Height + 14 + 22 + txtMessage.Height + 2 + 22)
        Me.Line (13, txtChat.Top + txtChat.Height + 13)-(13, txtChat.Top + txtChat.Height + 14 + 22 + txtMessage.Height + 2 + 22 + 1)
        Me.Line (Me.ScaleWidth - 13, txtChat.Top + txtChat.Height + 13)-(Me.ScaleWidth - 13, txtChat.Top + txtChat.Height + 14 + 22 + txtMessage.Height + 2 + 22 + 1)
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    If Not Me.Tag = vbNullString Then
        Exit Sub
    End If
    
    WindowLoaded = False
    Call LogChat(objMSN_SB.Contact, "----" & vbCrLf & "[" & Now & "] You closed the window." & vbCrLf & "----")
    
    Call QueScript(Me, "imwindowclosed", ConvArray(frmMain.objMSN_NS.Login))
    
    Dim i As Integer
    For i = 1 To Invitations.Count
        Unload Invitations(i)
        Invitations.Remove 1
    Next
    
    Call CleanDpTransfers

    If objMSN_SB.State = NsState_Connected Then
        objMSN_SB.Disconnect
    End If
    
    If InCollection(IMWindows, objMSN_SB.Contact) Then
        If IMWindows(objMSN_SB.Contact).hwnd = Me.hwnd Then
            IMWindows.Remove objMSN_SB.Contact
        End If
    ElseIf InCollection(PendingIM, objMSN_SB.Contact) Then
        If PendingIM(objMSN_SB.Contact).hwnd = Me.hwnd Then
            PendingIM.Remove objMSN_SB.Contact
        End If
    End If
    
    frmMain.Controls.Remove objMSN_SB.Socket.Name
    Set objMSN_SB = Nothing
End Sub

Private Sub imgBottomRight_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        ReleaseCapture
        SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, 0
    End If
End Sub

Private Sub imgBuddyDP_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        Me.PopupMenu mnuBuddyDP
    End If
End Sub

Private Sub imgEmoticon_Click()
    Call frmEmoticons.HideEmoticons
    Set frmEmoticons.SrcBox = txtMessage
    frmEmoticons.Left = Me.Left + (imgEmoticon.Left * Screen.TwipsPerPixelX)
    frmEmoticons.Top = Me.Top + Me.Height - 1500 - frmEmoticons.Height
    frmEmoticons.Visible = True
End Sub

Private Sub imgEmoticon_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgEmoticon.BorderStyle = vbFixedSingle
End Sub

Private Sub imgEmoticon_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgEmoticon.BorderStyle = vbBSNone
End Sub

Private Sub imgFont_Click()
    Call mnuEdit_ChangeFont_Click
End Sub

Private Sub imgFont_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgFont.BorderStyle = vbFixedSingle
End Sub

Private Sub imgFont_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgFont.BorderStyle = vbBSNone
End Sub

Private Sub imgMyDP_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        Me.PopupMenu mnuMyDP
    End If
End Sub

Private Sub imgShowHideBuddyDP_Click()
    If imgBuddyDP.Width = 1 Then
        imgBuddyDP.Width = 120
        imgBuddyDP.Height = 120
        imgBuddyDP.BorderStyle = vbFixedSingle
        Call RefreshBuddyDP
        SaveSettingX "App Settings\" & frmMain.objMSN_NS.Login & "\Show DP", objMSN_SB.Contact, True
    Else
        imgBuddyDP.Width = 1
        imgBuddyDP.Height = imgShowHideBuddyDP.Height
        imgBuddyDP.BorderStyle = vbBSNone
        Set imgBuddyDP.Picture = LoadPicture(vbNullString)
        SaveSettingX "App Settings\" & frmMain.objMSN_NS.Login & "\Show DP", objMSN_SB.Contact, False
    End If
    Call Form_Resize
End Sub

Private Sub imgShowHideMyDP_Click()
    If imgMyDP.Width = 1 Then
        ShowMyDP = True
        imgMyDP.Width = 81
        imgMyDP.Height = 81
        imgMyDP.BorderStyle = vbFixedSingle
        Call RefreshMyDP
    Else
        ShowMyDP = False
        imgMyDP.Width = 1
        imgMyDP.Height = imgShowHideMyDP.Height
        imgMyDP.BorderStyle = vbBSNone
        Set imgMyDP.Picture = LoadPicture(vbNullString)
    End If
    Call Form_Resize
End Sub

Private Sub lblBuddies_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        If Not lblBuddies.Caption = vbNullString Then
            lblBuddies.Caption = vbNullString
        Else
            Call UpdateBuddies
        End If
    End If
End Sub

Private Sub lblBuddyStatus_Change()
    Call ResizeScrollLabelSet(picBuddyStatus, lblBuddyStatus, vsBuddyStatus)
    Call vsBuddyStatus_Change
End Sub

Private Sub lblHide_Click()
    If objMSN_SB.State = SbState_Disconnected Then
        Unload Me
    Else
        Me.Visible = False
    End If
End Sub

Private Sub lblStatus_Change()
    lblStatus.ToolTipText = lblStatus.Caption
End Sub

Private Sub lblStatus_DblClick()
    ShellExecute 0, "open", MessageHistoryFolder & "\" & objMSN_SB.Contact & ".txt", vbNullString, vbNullString, 1
End Sub

Private Sub lblStatus_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        lblStatus.Caption = vbNullString
    End If
End Sub

Private Sub mnuActions_InviteSomeoneToJoinThisConversation_Click()
    Dim strContact As String
    strContact = InputBox("Enter the email of the person you want to invite.", "Invite Someone to this Conversation")
    If Not strContact = vbNullString Then
        objMSN_SB.InviteContact strContact
    End If
End Sub

Private Sub mnuActions_ManageContacts_AddToContacts_Contact_Click(Index As Integer)
    AddContact mnuActions_ManageContacts_AddToContacts_Contact(Index).Tag, mnuActions_ManageContacts_AddToContacts_Contact(Index).Caption
End Sub

Private Sub mnuActions_ManageContacts_Block_Contact_Click(Index As Integer)
    BlockContact mnuActions_ManageContacts_Block_Contact(Index).Tag
End Sub

Private Sub mnuActions_ManageContacts_Click()
    On Error Resume Next
    
    ClearSubMenu mnuActions_ManageContacts_AddToContacts_Contact
    ClearSubMenu mnuActions_ManageContacts_Block_Contact
    ClearSubMenu mnuActions_ManageContacts_ViewProfile_Contact
    ClearSubMenu mnuActions_ManageContacts_Properties_Contact
    
    Dim i As Integer, Email As String, Nick As String
    
    If ChatBuddies.Count = 0 Then
        Email = objMSN_SB.Contact
        Nick = GetContactAttr(Email, "nick")
    Else
        Email = ChatBuddies(1).Item("email")
        Nick = GetCustomNick(Email, ChatBuddies(1).Item("nick"))
    End If
    
    If Not InList(GetContactAttr(Email, "lists"), msnList_Forward) Then
        AddSubMenu mnuActions_ManageContacts_AddToContacts_Contact, Nick, Email
    End If
    If Not InList(GetContactAttr(Email, "lists"), msnList_Block) Then
        AddSubMenu mnuActions_ManageContacts_Block_Contact, Nick, Email
    End If
    AddSubMenu mnuActions_ManageContacts_ViewProfile_Contact, Nick, Email
    AddSubMenu mnuActions_ManageContacts_Properties_Contact, Nick, Email
    
    For i = 2 To ChatBuddies.Count
        Email = ChatBuddies(i).Item("email")
        Nick = GetCustomNick(Email, ChatBuddies(i).Item("nick"))
        
        If Not InList(GetContactAttr(Email, "lists"), msnList_Forward) Then
            AddSubMenu mnuActions_ManageContacts_AddToContacts_Contact, Nick, Email
        End If
        If Not InList(GetContactAttr(Email, "lists"), msnList_Block) Then
            AddSubMenu mnuActions_ManageContacts_Block_Contact, Nick, Email
        End If
        AddSubMenu mnuActions_ManageContacts_ViewProfile_Contact, Nick, Email
        AddSubMenu mnuActions_ManageContacts_Properties_Contact, Nick, Email
    Next
End Sub

Private Sub mnuActions_ManageContacts_Properties_Contact_Click(Index As Integer)
    ShowBuddyProperties Me, mnuActions_ManageContacts_Properties_Contact(Index).Tag, mnuActions_ManageContacts_Properties_Contact(Index).Caption
End Sub

Private Sub mnuActions_ManageContacts_ViewProfile_Contact_Click(Index As Integer)
    Call WebNavigate("http://members.msn.com/" & mnuActions_ManageContacts_ViewProfile_Contact(Index).Tag)
End Sub

Private Sub mnuActions_SendAFileOrPhoto_Click()
    Call mnuFile_SendAFileOrPhoto_Click
End Sub

Private Sub mnuActions_SendEmail_Click()
    Call SendEmail(objMSN_SB.Contact)
End Sub

Private Sub mnuBuddyDP_Hide_Click()
    Call imgShowHideBuddyDP_Click
End Sub

Private Sub mnuBuddyDP_Save_Click()
    If Not GetUserFile("Bitmap (*.bmp)|*.bmp", "Save Buddy Display Pic", 1) = vbNullString Then
        SavePicture imgBuddyDP.Picture, frmMain.CommonDialog.FileName
        MsgBox "Display pic has been saved!", vbInformation
    End If
End Sub

Private Sub mnuBuddyDP_ViewProfile_Click()
    Call WebNavigate("http://members.msn.com/" & objMSN_SB.Contact)
End Sub

Private Sub mnuEdit_ChangeColor_Click()
    With frmMain.CommonDialog
        .Color = IIf(mnuTools_RandomColors.Checked, PrevFontColor, txtMessage.ForeColor)
        .Flags = cdlCCFullOpen Or cdlCCRGBInit
        .ShowColor
        If Not mnuTools_RandomColors.Checked Then
            txtMessage.ForeColor = .Color
        End If
        IMFontColor = .Color
    End With
End Sub

Private Sub mnuEdit_ChangeFont_Click()
    On Error Resume Next
    
    With frmMain.CommonDialog
        .Flags = cdlCFScreenFonts
        .FontName = IIf(mnuTools_RandomFormat.Checked, PrevFontName, txtMessage.FontName)
        .FontBold = IIf(mnuTools_RandomFormat.Checked, PrevFontBold, txtMessage.FontBold)
        .FontItalic = IIf(mnuTools_RandomFormat.Checked, PrevFontItalic, txtMessage.FontItalic)
        .ShowFont
        If Not mnuTools_RandomFormat.Checked Then
            txtMessage.FontName = .FontName
            txtMessage.FontBold = .FontBold
            txtMessage.FontItalic = .FontItalic
        End If
        IMFontName = .FontName
        IMFontBold = .FontBold
        IMFontItalic = .FontItalic
    End With
End Sub

Private Sub mnuEdit_Click()
    Select Case ActiveControl
    Case txtChat
        mnuEdit_Undo.Enabled = False
        mnuEdit_Cut.Enabled = False
        mnuEdit_Copy.Enabled = (Not txtChat.SelText = vbNullString)
        mnuEdit_Paste.Enabled = False
        mnuEdit_Delete.Enabled = False
        mnuEdit_SelectAll.Enabled = (Not txtChat.Text = vbNullString)
    Case txtMessage
        mnuEdit_Undo.Enabled = (Not txtMessage.Text = vbNullString)
        mnuEdit_Cut.Enabled = (Not txtMessage.SelText = vbNullString)
        mnuEdit_Copy.Enabled = mnuEdit_Cut.Enabled
        mnuEdit_Paste.Enabled = (Not Clipboard.GetText = vbNullString)
        mnuEdit_Delete.Enabled = mnuEdit_Copy.Enabled
        mnuEdit_SelectAll.Enabled = mnuEdit_Undo.Enabled
    Case Else
        mnuEdit_Undo.Enabled = False
        mnuEdit_Cut.Enabled = False
        mnuEdit_Copy.Enabled = False
        mnuEdit_Paste.Enabled = False
        mnuEdit_Delete.Enabled = False
        mnuEdit_SelectAll.Enabled = False
    End Select
End Sub

Private Sub mnuEdit_Copy_Click()
    On Error Resume Next
    
    SendMessage ActiveControl.hwnd, WM_COPY, 0, 0
End Sub

Private Sub mnuEdit_Cut_Click()
    On Error Resume Next
    
    SendMessage ActiveControl.hwnd, WM_CUT, 0, 0
End Sub

Private Sub mnuEdit_Delete_Click()
    On Error Resume Next
    
    ActiveControl.Text = vbNullString
End Sub

Private Sub mnuEdit_Paste_Click()
    On Error Resume Next
    
    SendMessage ActiveControl.hwnd, WM_PASTE, 0, 0
End Sub

Private Sub mnuEdit_SelectAll_Click()
    On Error Resume Next
    
    ActiveControl.SelStart = 0
    ActiveControl.SelLength = Len(ActiveControl.Text)
End Sub

Private Sub mnuEdit_Undo_Click()
    On Error Resume Next
    
    SendMessage ActiveControl.hwnd, EM_UNDO, 0, vbNull
End Sub

Private Sub mnuFile_Close_Click()
    Unload Me
End Sub

Private Sub mnuFile_OpenMessageHistory_Click()
    ShellExecute 0, "open", MessageHistoryFolder & "\" & objMSN_SB.Contact & ".txt", vbNullString, vbNullString, 1
End Sub

Private Sub mnuFile_OpenReceivedFiles_Click()
    ShellExecute 0, "open", ReceivedFilesFolder, vbNullString, vbNullString, 1
End Sub

Private Sub mnuFile_Save_Click()
    If mnuFile_Save.Tag = vbNullString Then
        If Not GetUserFile("RTF Document|*.rtf|Text Document|*.txt", "Save", 1) = vbNullString Then
            mnuFile_Save.Tag = frmMain.CommonDialog.FileName
            txtChat.SaveFile mnuFile_Save.Tag
        End If
    Else
        txtChat.SaveFile frmMain.CommonDialog.Tag
    End If
End Sub

Private Sub mnuFile_SaveAs_Click()
    If Not GetUserFile("RTF Document|*.rtf|Text Document|*.txt", "Save As", 1) = vbNullString Then
        mnuFile_Save.Tag = frmMain.CommonDialog.FileName
        txtChat.SaveFile mnuFile_Save.Tag
    End If
End Sub

Private Sub mnuFile_SendAFileOrPhoto_Click()
    On Error Resume Next
    
    If Not GetUserFile("All Files|*.*", "Send a File to " & BuddyNick) = vbNullString Then
        SendFile Me, frmMain.CommonDialog.FileName, frmMain.CommonDialog.FileTitle
    End If
End Sub

Private Sub mnuHelp_AboutGillyMessenger_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuHelp_CrackSoftForums_Click()
    Call WebNavigate("http://www.cracksoft.net/forums")
End Sub

Private Sub mnuHelp_CrackSoftWebsite_Click()
    Call WebNavigate("http://www.cracksoft.net")
End Sub

Private Sub mnuHelp_Readme_Click()
    ShellExecute 0, "open", App.Path & "\readme.htm", vbNullString, vbNullString, 1
End Sub

Private Sub mnuMyDP_ChangeMyDisplayPic_Click()
    frmDisplayPic.Show vbModal, Me
End Sub

Private Sub mnuMyDP_Hide_Click()
    Call imgShowHideMyDP_Click
End Sub

Private Sub mnuTools_ChangeDisplayPic_Click()
    frmDisplayPic.Show vbModal, Me
End Sub

Private Sub mnuTools_FakeNick_Click()
    If mnuTools_FakeNick.Checked Then
        mnuTools_FakeNick.Checked = False
    Else
        Dim strFakeNick As String
        strFakeNick = InputBox("Enter a fake nick for this conversation.", "Fake Nick", mnuTools_FakeNick.Tag)
        If Not strFakeNick = vbNullString Then
            mnuTools_FakeNick.Tag = strFakeNick
            mnuTools_FakeNick.Checked = True
        End If
    End If
End Sub

Private Sub mnuTools_EmoticonFloodControl_Click()
    mnuTools_EmoticonFloodControl.Checked = Not mnuTools_EmoticonFloodControl.Checked
    EmoticonFloodControl = mnuTools_EmoticonFloodControl.Checked
End Sub

Private Sub mnuTools_Encryption_Click()
    mnuTools_Encryption.Checked = Not mnuTools_Encryption.Checked
End Sub

Private Sub mnuTools_Options_Click()
    On Error Resume Next
    
    frmOptions.Show vbModal, Me
End Sub

Private Sub mnuTools_RandomColors_Click()
    mnuTools_RandomColors.Checked = Not mnuTools_RandomColors.Checked
    IMFontRandomColors = mnuTools_RandomColors.Checked
    
    With txtMessage
        If mnuTools_RandomColors.Checked Then
            If PrevFontColor = vbNullString Then
                PrevFontColor = txtMessage.ForeColor
            End If
            Randomize Timer
            txtMessage.ForeColor = RGB(Fix(Rnd() * 256), Fix(Rnd() * 256), Fix(Rnd() * 256))
        Else
            txtMessage.ForeColor = Val(PrevFontColor)
            PrevFontColor = vbNullString
        End If
    End With
End Sub

Private Sub mnuTools_RandomFormat_Click()
    mnuTools_RandomFormat.Checked = Not mnuTools_RandomFormat.Checked
    IMFontRandomFormat = mnuTools_RandomFormat.Checked
    
    With txtMessage
        If mnuTools_RandomFormat.Checked Then
            If PrevFontBold = vbNullString Then
                PrevFontBold = txtMessage.FontBold
            End If
            If PrevFontItalic = vbNullString Then
                PrevFontItalic = txtMessage.FontItalic
            End If
            Randomize Timer
            txtMessage.FontBold = CBool(Fix(Rnd() * 2))
            txtMessage.FontItalic = CBool(Fix(Rnd() * 2))
        Else
            txtMessage.FontBold = PrevFontBold
            txtMessage.FontItalic = PrevFontItalic
            PrevFontBold = vbNullString
            PrevFontItalic = vbNullString
        End If
    End With
End Sub

Private Sub mnuTools_TextSize_Size_Click(Index As Integer)
    On Error Resume Next
    
    mnuTools_TextSize_Size(IMFontSize).Checked = False
    mnuTools_TextSize_Size(Index).Checked = True
    IMFontSize = Index
    txtMessage.FontSize = Choose(IMFontSize + 1, 13, 12, 10, 8, 7)
    Dim intSelStart As Integer, intSelLength As Integer
    intSelStart = txtChat.SelStart
    intSelLength = txtChat.SelLength
    txtChat.SelStart = 0
    txtChat.SelLength = Len(txtChat.Text)
    txtChat.SelFontSize = txtMessage.FontSize
    txtChat.SelStart = intSelStart
    txtChat.SelLength = intSelLength
End Sub

Private Sub mnuTools_TextStyler_Other_Click()
    mnuTools_TextStyler_Other.Checked = Not mnuTools_TextStyler_Other.Checked
    If Not mnuTools_TextStyler_Other.Checked Then
        Set TextStyler = Nothing
        mnuTools_TextStyler.Tag = vbNullString
        TextStyle = vbNullString
    Else
        If Not GetUserFile("GM Text Styles (*.gts)|*.gts") = vbNullString Then
            If IsNumeric(mnuTools_TextStyler.Tag) Then
                mnuTools_TextStyler_Style(mnuTools_TextStyler.Tag).Checked = False
            End If
            Call LoadTextStyle(frmMain.CommonDialog.FileName)
            mnuTools_TextStyler.Tag = "other"
            TextStyle = mnuTools_TextStyler_Other.Tag
        Else
            mnuTools_TextStyler_Other.Checked = False
        End If
    End If
End Sub

Private Sub mnuTools_TextStyler_Style_Click(Index As Integer)
    mnuTools_TextStyler_Style(Index).Checked = Not mnuTools_TextStyler_Style(Index).Checked
    If Not mnuTools_TextStyler_Style(Index).Checked Then
        Set TextStyler = Nothing
        If IsNumeric(mnuTools_TextStyler.Tag) Then
            mnuTools_TextStyler_Style(mnuTools_TextStyler.Tag).Checked = False
        Else
            mnuTools_TextStyler_Other.Checked = False
        End If
        mnuTools_TextStyler.Tag = vbNullString
        TextStyle = vbNullString
    Else
        If IsNumeric(mnuTools_TextStyler.Tag) Then
            mnuTools_TextStyler_Style(mnuTools_TextStyler.Tag).Checked = False
        Else
            mnuTools_TextStyler_Other.Checked = False
        End If
        Call LoadTextStyle(mnuTools_TextStyler_Style(Index).Tag)
        mnuTools_TextStyler.Tag = Index
        mnuTools_TextStyler_Style(Index).Checked = True
        TextStyle = mnuTools_TextStyler_Style(Index).Tag
    End If
End Sub

Private Sub mnuTools_TimeStamp_Click()
    mnuTools_TimeStamp.Checked = Not mnuTools_TimeStamp.Checked
    TimeStamp = mnuTools_TimeStamp.Checked
End Sub

Private Sub objMSN_SB_ContactJoined(Email As String, Nick As String)
    On Error Resume Next
    
    Dim NewBuddy As Collection
    Set NewBuddy = New Collection
    
    NewBuddy.Add Email, "email"
    NewBuddy.Add Nick, "nick"
    ChatBuddies.Add NewBuddy, Email
    Set NewBuddy = Nothing
    
    Call SendQueMessages
    
    CallingContact = False
    If ChatBuddies.Count = 1 Then
        If objMSN_SB.State = SbState_Connected Then
            lblStatus.Caption = "[" & Time$ & "] Connected to " & Email
            Call LogChat(objMSN_SB.Contact, "----" & vbCrLf & "[" & Now & "] Connected to " & Email & vbCrLf & "----")
            SaveSettingX "Statistics\" & Email, "Last ConversationJoined", Now()
            Call OfferDP
        End If
    Else
        If objMSN_SB.State = SbState_Connected Then
            If InCollection(IMWindows, objMSN_SB.Contact) Then
                IMWindows.Remove objMSN_SB.Contact
            End If
            Comment Email & " has joined the conversation."
            Call UpdateBuddies
            If lblBuddyStatus.Visible Then
                picBuddyStatus.Visible = False
                lblBuddyStatus.Visible = False
                Call Form_Resize
            End If
            SaveSettingX "Statistics\" & Email, "Last ConversationJoined", Now()
        End If
    End If
End Sub

Private Sub objMSN_SB_ContactLeft(Email As String)
    On Error Resume Next
    
    ChatBuddies.Remove Email

    If ChatBuddies.Count = 1 Then
        objMSN_SB.Contact = ChatBuddies(1).Item("email")
        Me.Caption = BuddyNick & " - Conversation"
        If Not InCollection(IMWindows, objMSN_SB.Contact) Then
            IMWindows.Add Me, objMSN_SB.Contact
        End If
    End If
    
    If ChatBuddies.Count = 0 Then
        If InCollection(ContactList, Email) Then
            If ContactList(Email).Item("status") = msnStatus_Offline Then
                lblStatus.Caption = "[" & Time$ & "] " & Email & " appears to be offline."
                Call LogChat(objMSN_SB.Contact, "----" & vbCrLf & "[" & Now & "] " & Email & " appears to be offline." & vbCrLf & "----")
            Else
                lblStatus.Caption = "[" & Time$ & "] " & Email & " has closed your window."
                Call LogChat(objMSN_SB.Contact, "----" & vbCrLf & "[" & Now & "] " & Email & " has closed your window." & vbCrLf & "----")
            End If
        Else
            lblStatus.Caption = "[" & Time$ & "] " & Email & " has closed your window."
            Call LogChat(objMSN_SB.Contact, "----" & vbCrLf & "[" & Now & "] " & Email & " has closed your window." & vbCrLf & "----")
        End If
        If Not Me.Visible Then
            Unload Me
        End If
    Else
        Comment Email & " has left the conversation."
    End If
    
    Call UpdateBuddies
    
    Call QueScript(Me, "imwindowclosed", ConvArray(Email))
End Sub

Private Sub objMSN_SB_CustomMessageReceived(Email As String, MsgType As String, Header As String, Message As String)
    On Error Resume Next
    
    Dim Headers() As String
    Headers = Split(Header, vbCrLf)

    Select Case MsgType
    Case "gm-displaypic"
        Select Case Split(Headers(0))(0)
        Case "id:"
            SetCollectionItem ChatBuddies(Email), "dpoffered", True
            If ReceiveDisplayPic Then
                If GetSettingX("Display Pics", Email) <> Split(Headers(0))(1) Or Not FileExists(App.Path & "\Display Pics\" & Email & ".dat") Then
                    objMSN_SB.SendCustomMessage "gm-displaypic", "action: request", vbNullString
                Else
                    Call RefreshBuddyDP
                End If
            End If
        Case "action:"
            Dim fptr As Integer, DpPath As String, DpData As String
            Select Case Split(Headers(0))(1)
            Case "request"
                If SendDisplayPic Then
                    Dim DpId As String
                    DpId = GetSettingX("Display Pics", frmMain.objMSN_NS.Login)
                    DpPath = App.Path & "\Display Pics\" & frmMain.objMSN_NS.Login & ".dat"
                    If Not DpId = vbNullString And FileExists(DpPath) Then
                        fptr = FreeFile
                        Open DpPath For Binary As #fptr
                        DpData = Space$(LOF(fptr))
                        Get #fptr, , DpData
                        Close #fptr
                        DpData = Base64_Encode(DpData)
                        objMSN_SB.SendCustomMessage "gm-displaypic", "action: transfer" & vbCrLf & DpId & vbCrLf & Len(DpData), vbNullString
                        DoEvents
                        objMSN_SB.SendCustomMessage "gm-displaypic", "action: data", DpData
                    End If
                End If
            Case "transfer"
                If ReceiveDisplayPic And Not InCollection(DpTransfers, Email) Then
                    Dim DisplayPic As Collection
                    Set DisplayPic = New Collection
                    DisplayPic.Add Headers(1), "id"
                    DisplayPic.Add Headers(2), "size"
                    DisplayPic.Add Me.hwnd, "hWnd"
                    fptr = FreeFile
                    DisplayPic.Add fptr, "fptr"
                    DpTransfers.Add DisplayPic, Email
                    Set DisplayPic = Nothing
                    MakeSureDirectoryPathExists App.Path & "\Display Pics\"
                    DpPath = App.Path & "\Display Pics\" & Email & ".dat"
                    Kill DpPath
                    Open DpPath For Binary As #fptr
                End If
            Case "data"
                If InCollection(DpTransfers, Email) Then
                    Put #DpTransfers(Email).Item("fptr"), , Message
                    Dim Progress As Integer
                    Progress = DpTransfers(Email).Item("progress")
                    Progress = Progress + Len(Message)
                    SetCollectionItem DpTransfers(Email), "progress", Progress
                    If Progress >= DpTransfers(Email).Item("size") Then
                        Close #DpTransfers(Email).Item("fptr")
                        fptr = FreeFile
                        DpPath = App.Path & "\Display Pics\" & Email & ".dat"
                        Open DpPath For Binary As #fptr
                        DpData = Space$(LOF(fptr))
                        Get #fptr, , DpData
                        Close #fptr
                        Kill DpPath
                        fptr = FreeFile
                        Open DpPath For Binary As #fptr
                        Put #fptr, , Base64_Decode(DpData)
                        Close #fptr
                        SaveSettingX "Display Pics", Email, DpTransfers(Email).Item("id")
                        DpTransfers.Remove Email
                        Call RefreshBuddyDP(True)
                    End If
                End If
            End Select
        End Select
    End Select
End Sub

Private Sub objMSN_SB_InvitationAccepted(Cookie As Double, Attributes As Collection)
    On Error Resume Next
    
    If InCollection(Invitations, "Cookie " & Cookie) Then
        With Invitations("Cookie " & Cookie)
            If .objMSN_FTP.TransferType = FtpTransferType_Receive Then
                If InCollection(Attributes, "AuthCookie") Then
                    .objMSN_FTP.AuthCookie = Attributes("AuthCookie")
                ElseIf InCollection(Attributes, "Auth-Cookie") Then
                    .objMSN_FTP.AuthCookie = Attributes("Auth-Cookie")
                End If
            End If
            If InCollection(Attributes, "Request-Data") And Not InCollection(Attributes, "IP-Address") Then
                If .objMSN_FTP.TransferType = FtpTransferType_Send Then
                    objMSN_SB.AcceptInvitation Cookie, "AuthCookie: " & .objMSN_FTP.AuthCookie, "IP-Address: " & frmMain.wskNS.LocalIP, "Port: " & .objMSN_FTP.Listen(FTPPort)
                    Call QueScript(Me, "TransferAccepted", ConvArray(objMSN_SB.Contact, .Cookie))
                Else
                    objMSN_SB.AcceptInvitation Cookie, "IP-Address: " & frmMain.wskNS.LocalIP, "Port: " & .objMSN_FTP.Listen(FTPPort)
                End If
            Else
                SaveSettingX "Statistics\" & objMSN_SB.Contact, "Last IP", Attributes("IP-Address")
                .objMSN_FTP.Connect Attributes("IP-Address"), Attributes("Port")
                .lblTransfer.Caption = "Receiving from: " & objMSN_SB.Contact
            End If
        End With
    End If
End Sub

Private Sub objMSN_SB_InvitationCancelled(Cookie As Double, CancelCode As String, Attributes As Collection)
    On Error Resume Next
    
    If InCollection(Invitations, "Cookie " & Cookie) Then
        With Invitations("Cookie " & Cookie)
            If Not .objMSN_FTP.File = vbNullString Then
                Select Case CancelCode
                Case "REJECT"
                    Comment BuddyNick & " has rejected the transfer of """ & .objMSN_FTP.File & """.", 128
                    Call QueScript(Me, "TransferCancelled", ConvArray(objMSN_SB.Contact, Cookie))
                Case "TIMEOUT"
                    Comment BuddyNick & " has cancelled the trasnfer of """ & .objMSN_FTP.File & """.", 128
                    Call QueScript(Me, "TransferFailed", ConvArray(Cookie))
                Case Else
                    Comment "You have failed to receive file """ & .objMSN_FTP.File & """ from " & BuddyNick, 128
                    Call QueScript(Me, "TransferFailed", ConvArray(Cookie))
                End Select
            End If
            Call .Terminate
        End With
    End If
End Sub

Private Sub objMSN_SB_InvitationReceived(AppName As String, AppGUID As String, Cookie As Double, Attributes As Collection)
    On Error Resume Next
    
    If AppName = "File Transfer" Or AppGUID = "{5D3E02AB-6190-11d3-BBBB-00C04F795683}" Then
        Dim TransferForm As New frmTransfer
        Load TransferForm
        With TransferForm
            Set .Parent = Me
            
            If Me.WindowState = vbMinimized Then
                ShowWindow Me.hwnd, SW_RESTORE
            End If
            
            .Cookie = Cookie
            
            .objMSN_FTP.Login = objMSN_SB.Login
            .objMSN_FTP.File = Attributes("Application-File")
            .objMSN_FTP.FileSize = Attributes("Application-FileSize")
            MakeSureDirectoryPathExists ReceivedFilesFolder
            .objMSN_FTP.FilePath = ReceivedFilesFolder & "\" & .objMSN_FTP.File
            .objMSN_FTP.TransferType = FtpTransferType_Receive
                
            .lblFileName.Caption = "File name: " & .objMSN_FTP.File
            .lblFileSize.Caption = "File size: " & ConvertBytes(.objMSN_FTP.FileSize)
            .lblTransfer.Caption = "Request from: " & objMSN_SB.Contact
            
            Call SaveFocus
            .Show , Me
            Call RestoreFocus
            
            Call SetRandomPos(Me, TransferForm)
            Call QueScript(Me, "transferrequest", ConvArray(objMSN_SB.Contact, .objMSN_FTP.File, .objMSN_FTP.FileSize, Cookie))
        End With
        Invitations.Add TransferForm, "Cookie " & Cookie
    Else
        Comment BuddyNick & " has sent you an invitation for " & AppName & "."
    End If
    If Not Me.Visible Then
        Me.Visible = True
        Call Form_Resize
    End If
    Call FlashWindowEx(Me.hwnd)
End Sub

Private Sub objMSN_SB_MessageFailure()
    On Error Resume Next
    
    If Not (Right$(lblStatus.Caption, 21) = " is typing a message." Or Right$(lblStatus.Caption, 31) = "Message could not be delivered.") Then
        lblStatus.Tag = lblStatus.Caption
    End If
    lblStatus.Caption = "[" & Time$ & "] Message could not be delivered."
    tmrResetStatus.Enabled = True
End Sub

Private Sub objMSN_SB_MessageReceived(Email As String, Nick As String, FontName As String, FontColor As Long, FontBold As Boolean, FontItalic As Boolean, FontStrikethru As Boolean, FontUnderline As Boolean, Message As String, FakeNick As String)
    On Error Resume Next
    
    If FakeNick = vbNullString Then
        SetCollectionItem ChatBuddies(Email), "fontname", FontName
        SetCollectionItem ChatBuddies(Email), "fontcolor", FontColor
        SetCollectionItem ChatBuddies(Email), "fontbold", FontBold
        SetCollectionItem ChatBuddies(Email), "fontitalic", FontItalic
        SetCollectionItem ChatBuddies(Email), "fontstrikethru", FontStrikethru
        SetCollectionItem ChatBuddies(Email), "fontunderline", FontUnderline
    End If
    
    If mnuTools_Encryption.Checked Then
        Message = XorDecrypt(Message, Email)
    End If
    
    If InCollection(ContactCustomNicks, Email) Then
        Nick = GetBuddyCustomNick(Email)
    ElseIf IsOnline(Email) Then
        Nick = GetContactAttr(Email, "nick")
    End If
    
    AddChat IIf(FakeNick = vbNullString, Nick, FakeNick), FontName, FontColor, FontBold, FontItalic, FontStrikethru, FontUnderline, Message
    Call LogChat(Email, "[" & Now & "] " & Nick & " : " & vbCrLf & Space$(3) & Message)
    lblStatus.Caption = "Last message received on " & Now() & "."
    
    If Not Me.Visible Then
        Me.Visible = True
        Call Form_Resize
    End If
    
    Call FlashWindowEx(Me.hwnd)
    
    If Message = "BUZZ!!!" And FontName = "Verdana" And FontColor = vbRed And FontBold = True And FontItalic = False And FontStrikethru = False And FontUnderline = False Then
        ActivateWindow Me
        VibrateWindow Me
    End If
    
    If frmMain.mnuTools_AutoMessage.Checked Then
        SendMsg Me, frmMain.mnuTools_AutoMessage.Tag
    End If
    
    If Not FirstMsgReceived Then
        FirstMsgReceived = True
        If AlertOnMessageReceived And Not GetForegroundWindow = Me.hwnd Then
            ShowPopup Me, "CONVO", GetCustomNick(Email, Nick) & " says: " & vbCrLf & Message
        End If
    End If
    
    SaveSettingX "Statistics\" & Email, "Last MessageSent", Now()
    
    If SoundAlerts And boolMessageSound And GetForegroundWindow <> Me.hwnd Then
        ContactSound Email, strMessageSound, "contactim"
    End If
    
    If Not InCollection(RC_Sessions, Email) Then
        If RemoteControl And StrComp(Message, "gm remote control request", vbTextCompare) = 0 And RcSessionLevel = 0 Then
            RcSessionLevel = 1
            RcUser = Email
            SendMsg Me, "Welcome to GM Remote Control" & vbCrLf & vbCrLf & "Please enter your login:"
        ElseIf RemoteControl And RcSessionLevel = 1 And RcUser = Email Then
            RcSessionLevel = 2
            RcLogin = Message
            SendMsg Me, "Please enter your password:"
        ElseIf RemoteControl And RcSessionLevel = 2 And RcUser = Email Then
            If Not InCollection(RC_Accounts, RcLogin) Then
                RcSessionLevel = 0
                SendMsg Me, "Invalid username or password."
            Else
                If Not RC_Accounts(RcLogin).Item("password") = Message Then
                    RcSessionLevel = 0
                    SendMsg Me, "Invalid username or password."
                Else
                    RcSessionLevel = 0
                    Dim NewSession As Collection
                    Set NewSession = New Collection
                    NewSession.Add RcUser, "email"
                    NewSession.Add RcLogin, "login"
                    NewSession.Add Now, "stamp"
                    RC_Sessions.Add NewSession, Email
                    Set NewSession = Nothing
                    SendMsg Me, "Logged in."
                End If
            End If
        Else
            If Not ArraySize(ChatBot) = -1 Then
                Dim strReply As String
                strReply = BotReply(Message)
                If Not strReply = vbNullString Then
                    txtMessage.Tag = txtMessage.Text
                    txtMessage.Text = strReply
                    Call cmdSend_Click
                    txtMessage.Text = txtMessage.Tag
                End If
            End If
        End If
    ElseIf RemoteControl Then
        Dim RcResponse As String
        RcResponse = RcProcess(Me, Email, RcLogin, Message)
        If Not RcResponse = vbNullString Then
            SendMsg Me, RcResponse
        End If
    End If
    
    Call QueScript(Me, "messagereceived", ConvArray(Email, Nick, FontName, FontColor, FontBold, FontItalic, FontStrikethru, FontUnderline, Message))
End Sub

Private Sub objMSN_SB_SbError(Error As String)
    If CallingContact Then
        CallingContact = False
        If ChatBuddies.Count = 0 Then
            lblStatus.Caption = vbNullString
        End If
    End If
    Select Case Left$(Error, 3)
    Case "201", "208"
        Comment "Server Error: Contact does not exist.", , False
    Case "217"
        Comment "Server Error: Contact not online.", , False
    Case "280"
        Comment "Server Error: Switchboard failed.", , False
    Case "281"
        Comment "Server Error: Transfer to switchboard failed.", , False
    Case "500"
        Comment "Server Error: Internal server error.", , False
    Case "501"
        Comment "Server Error: Database server error.", , False
    Case "510"
        Comment "Server Error: File operation failed.", , False
    Case "520"
        Comment "Server Error: Memory allocation failed.", , False
    Case "600", "910", "912", "918", "919", "921", "922"
        Comment "Server Error: Server is busy.", , False
    Case "601", "605"
        Comment "Server Error: Server is unavailable.", , False
    Case "602"
        Comment "Server Error: Peer name server is down.", , False
    Case "603"
        Comment "Server Error: Database connection failed.", , False
    Case "604"
        Comment "Server Error: Server is going down.", , False
    Case "707"
        Comment "Server Error: Could not create connection.", , False
    Case "711"
        Comment "Server Error: Write is blocking.", , False
    Case "712"
        Comment "Server Error: Session is overloaded.", , False
    Case "713"
        Comment "Server Error: Calling too rapidly.", , False
    Case "714"
        Comment "Server Error: Too many sessions.", , False
    Case "717"
        Comment "Server Error: Bad friend file.", , False
    Case "914", "915", "916"
        Comment "Server Error: Server unavailable.", , False
    Case "920"
        Comment "Server Error: Not accepting new principles.", , False
    End Select
End Sub

Private Sub objMSN_SB_SocketError(Description As String)
    lblStatus.Tag = False
    lblStatus.Caption = "[" & Time$ & "] " & Description
End Sub

Private Sub objMSN_SB_StateChanged()
    On Error Resume Next
    
    Select Case objMSN_SB.State
    Case SbState_Connected
        mnuActions_InviteSomeoneToJoinThisConversation.Enabled = True
        If objMSN_SB.SessionType = SbSession_Call Then
            objMSN_SB.InviteContact objMSN_SB.Contact
        Else
            Dim i As Integer
            If ChatBuddies.Count = 1 Then
                If InCollection(IMWindows, objMSN_SB.Contact) Then
                    If Not IMWindows(objMSN_SB.Contact).hwnd = Me.hwnd Then
                        With IMWindows(objMSN_SB.Contact)
                            Set .ChatBuddies = ChatBuddies
                            frmMain.Controls.Remove .objMSN_SB.Socket.Name
                            Set .objMSN_SB = objMSN_SB
                            Call SendQueMessages
                            .lblStatus.Caption = "[" & Time$ & "] " & objMSN_SB.Contact & " has opened your window."
                            Call OfferDP
                        End With
                        Call LogChat(objMSN_SB.Contact, "---" & vbCrLf & "[" & Now & "] " & objMSN_SB.Contact & " has opened your window." & vbCrLf & "---")
                        Me.Tag = True
                        Unload Me
                    End If
                Else
                    IMWindows.Add Me, objMSN_SB.Contact
                    If Not (GetContactAttr(objMSN_SB.Contact, "status") = msnStatus_Online Or GetContactAttr(objMSN_SB.Contact, "status") = msnStatus_Unknown) Then
                        lblBuddyStatus.Caption = BuddyNick & " may or may not reply because his/her status is set to " & StatusName(GetContactAttr(objMSN_SB.Contact, "status")) & "."
                        picBuddyStatus.Visible = True
                        lblBuddyStatus.Visible = True
                    End If
                    Call SendQueMessages
                    If Not ShowIMWindowOnMsg Then
                        Me.Visible = True
                        Call Form_Resize
                    End If
                    Call OfferDP
                End If
            Else
                Call UpdateBuddies
                If Not ShowIMWindowOnMsg Then
                    Me.Visible = True
                    Call Form_Resize
                End If
            End If
            SaveSettingX "Statistics\" & objMSN_SB.Contact, "Last ConversationStarted", Now()
        End If
        
    Case SbState_Disconnected
        CallingContact = False
        mnuActions_InviteSomeoneToJoinThisConversation.Enabled = False
        Set ChatBuddies = Nothing
        Set ChatBuddies = New Collection
        
        If lblStatus.Tag = vbNullString Then
            If WindowLoaded Then
                lblStatus.Caption = "[" & Time$ & "] Chat session ended."
                Call LogChat(objMSN_SB.Contact, "----" & vbCrLf & "[" & Now & "] Chat session ended." & vbCrLf & "----")
            End If
        End If
        
        Call CleanDpTransfers
        
        If Not Me.Visible Then
            Unload Me
        End If
    End Select
End Sub

Private Sub objMSN_SB_TypingUser(Email As String)
    On Error Resume Next
    
    If Not (Right$(lblStatus.Caption, 21) = " is typing a message." Or Right$(lblStatus.Caption, 31) = "Message could not be delivered.") Then
        lblStatus.Tag = lblStatus.Caption
    End If
    lblStatus.Caption = CropText(Me, lblStatus.Width, "[" & Time$ & "] " & GetCustomNick(Email, ChatBuddies(Email).Item("nick")), " is typing a message.")
    tmrResetStatus.Enabled = True
    
    If SoundAlerts And boolTypingSound And GetForegroundWindow <> Me.hwnd Then
        ContactSound Email, strTypingSound, "contactim"
    End If
End Sub

Public Sub tmrResetStatus_Timer()
    On Error Resume Next
    
    tmrResetStatus.Enabled = False
    If Right$(lblStatus.Caption, 21) = " is typing a message." Or Right$(lblStatus.Caption, 31) = "Message could not be delivered." Then
        lblStatus.Caption = lblStatus.Tag
    End If
    lblStatus.Tag = vbNullString
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
    WordStop = IIf((i < j And i <> 0) Or (j = 0), i, j)
    If WordStop = 0 Then
        WordStop = TextLen + 1
    End If

    Word = Mid$(txtChat.Text, WordStart, WordStop - WordStart)
    If Word Like "www.*.*" Then
        Word = "http://" & Word
    End If
    If CBool(PathIsURL(Word)) Then
        Call WebNavigate(Word)
        txtMessage.SetFocus
    End If
Handler:
End Sub

Private Sub txtMessage_Change()
    If txtMessage.Text = vbNullString Then
        cmdSend.Enabled = False
    Else
        Dim TextLen As Integer
        TextLen = Len(txtMessage.Text)
        If Left$(txtMessage.Text, 1) = "/" And Len(txtMessage.Text) > 1 And InStr(txtMessage.Text, vbCrLf) = 0 And SearchTextBox And TextLen > PrevTextLen Then
            SearchTextBox = False
            Dim i As Integer
            For i = 0 To UBound(IMWindowCommands)
                If StrComp(Left$(IMWindowCommands(i, 0), TextLen), txtMessage.Text, vbTextCompare) = 0 Then
                    txtMessage.Text = txtMessage.Text & Right$(IMWindowCommands(i, 0), Len(IMWindowCommands(i, 0)) - TextLen)
                    txtMessage.SelStart = TextLen
                    txtMessage.SelLength = Len(txtMessage.Text) - TextLen
                    Exit For
                End If
            Next
            If i > UBound(IMWindowCommands) Then
                AutoCmd = False
                AutoCmd_ReqParam = False
            Else
                AutoCmd = True
                AutoCmd_ReqParam = IMWindowCommands(i, 1)
            End If
        End If
        PrevTextLen = TextLen
        cmdSend.Enabled = True
    End If
End Sub

Private Sub txtMessage_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyB And Shift = vbCtrlMask Then
        txtMessage.FontBold = Not txtMessage.FontBold
        IMFontBold = txtMessage.FontBold
    ElseIf KeyCode = vbKeyT And Shift = vbCtrlMask Then
        txtMessage.FontStrikethru = Not txtMessage.FontStrikethru
        IMFontStrikethru = txtMessage.FontStrikethru
    ElseIf KeyCode = vbKeyU And Shift = vbCtrlMask Then
        txtMessage.FontUnderline = Not txtMessage.FontUnderline
        IMFontUnderline = txtMessage.FontUnderline
    ElseIf KeyCode = vbKeyUp And Shift = vbCtrlMask Then
        txtMessage.Text = LastMsg
    ElseIf KeyCode = vbKeyDelete Then
        SearchTextBox = False
    End If
End Sub

Private Sub txtMessage_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    
    If Not KeyAscii = vbKeyTab Then
        boolTabEml = False
        intTabEmlCounter = 0
        strTabEmlLastCol = vbNullString
        strTabEmlKeyword = vbNullString
        intTabEmlStart = 0
        intTabEmlLen = 0
    End If
    
    If KeyAscii = vbKeyTab Then
        If Not boolTabEml Then
            boolTabEml = True
            Dim i As Integer, j As Integer, SelStart As Integer, WordStart As Integer, WordStop As Integer, Word As String, TextLen As Integer
            SelStart = txtMessage.SelStart
            TextLen = Len(txtMessage.Text)
            
            If SelStart = 0 Then
                WordStart = 1
            Else
                If Mid$(txtMessage.Text, SelStart, 1) = " " Or Mid$(txtMessage.Text, SelStart, 1) = vbCr Or Mid$(txtMessage.Text, SelStart, 1) = vbLf Then
                    WordStart = SelStart + 1
                Else
                    i = InStrRev(txtMessage.Text, " ", SelStart - 1)
                    j = InStrRev(txtMessage.Text, vbCrLf, SelStart - 1)
                    If i = 0 And j = 0 Then
                        WordStart = 1
                    Else
                        WordStart = IIf(i > j, i + 1, j + 2)
                    End If
                End If
            End If
                
            If SelStart = TextLen Then
                WordStop = TextLen + 1
            Else
                If Mid$(txtMessage.Text, SelStart, 1) = " " Or Mid$(txtMessage.Text, SelStart, 1) = vbCr Or Mid$(txtMessage.Text, SelStart, 1) = vbLf Then
                    WordStop = SelStart + 2
                Else
                    i = InStr(SelStart + 1, txtMessage.Text, " ")
                    j = InStr(SelStart + 1, txtMessage.Text, vbCrLf)
                    If i = 0 And j = 0 Then
                        WordStop = TextLen
                    Else
                        WordStop = IIf((i < j And i <> 0) Or (j = 0), i, j)
                    End If
                End If
            End If
            Word = Mid$(txtMessage.Text, WordStart, WordStop - WordStart)
            intTabEmlCounter = 0
            strTabEmlKeyword = Word
            intTabEmlStart = WordStart
            intTabEmlLen = WordStop - WordStart
        Else
            Word = strTabEmlKeyword
        End If
        
        If Word = vbNullString Then
            If (intTabEmlCounter = 1 And Not GetKeyState(vbKeyShift)) Or (intTabEmlCounter = 2 And GetKeyState(vbKeyShift)) Then
                txtMessage.SelStart = intTabEmlStart - 1
                txtMessage.SelLength = intTabEmlLen
                txtMessage.SelText = vbNullString
                intTabEmlCounter = 0
            Else
                If GetKeyState(vbKeyShift) Then
                    If ChatBuddies.Count = 0 Then
                        Word = GetContactAttr(objMSN_SB.Contact, "nick")
                    Else
                        Word = ChatBuddies(1).Item("nick")
                    End If
                    intTabEmlCounter = 2
                Else
                    Word = objMSN_SB.Contact
                    intTabEmlCounter = 1
                End If
                txtMessage.SelStart = intTabEmlStart - 1
                txtMessage.SelLength = intTabEmlLen
                txtMessage.SelText = Word
                intTabEmlLen = Len(Word)
                txtMessage.SelLength = 0
            End If
        Else
            Dim WordLen As Integer
            WordLen = Len(Word)
            Dim SearchCol As Collection
            Select Case strTabEmlLastCol
            Case "", "chatbuddies"
                Set SearchCol = ChatBuddies
            Case Else
                Set SearchCol = ContactList
            End Select
            Dim SearchItem As String
            If GetKeyState(vbKeyShift) Then
                SearchItem = "nick"
            Else
                SearchItem = "email"
            End If
            For i = intTabEmlCounter + 1 To SearchCol.Count
                If StrComp(Left$(SearchCol(i).Item(SearchItem), WordLen), Word, vbTextCompare) = 0 Then
                    txtMessage.SelStart = intTabEmlStart - 1
                    txtMessage.SelLength = intTabEmlLen
                    txtMessage.SelText = SearchCol(i).Item(SearchItem)
                    intTabEmlLen = Len(SearchCol(i).Item(SearchItem))
                    txtMessage.SelLength = 0
                    Exit For
                End If
            Next
            If i > SearchCol.Count Then
                intTabEmlCounter = 0
                Select Case strTabEmlLastCol
                Case "", "chatbuddies"
                    strTabEmlLastCol = "contactlist"
                    Call txtMessage_KeyPress(vbKeyTab)
                Case Else
                    strTabEmlLastCol = "chatbuddies"
                    txtMessage.SelStart = intTabEmlStart
                    txtMessage.SelLength = intTabEmlLen
                    txtMessage.SelText = vbNullString
                End Select
            Else
                intTabEmlCounter = i
            End If
        End If
        KeyAscii = 0
    ElseIf KeyAscii = vbKeyReturn And Not (GetKeyState(vbKeyShift) Or GetKeyState(vbKeyControl)) Then
        If Not txtMessage.Text = vbNullString Then
            Call cmdSend_Click
        End If
        KeyAscii = 0
    ElseIf KeyAscii = 9 And GetKeyState(vbKeyControl) And GetKeyState(vbKeyI) Then
        txtMessage.FontItalic = Not txtMessage.FontItalic
        IMFontItalic = txtMessage.FontItalic
        KeyAscii = 0
    Else
        If KeyAscii = vbKeyBack Then
            SearchTextBox = False
            AutoCmd = False
        Else
            SearchTextBox = True
        End If
        
        If TypingNotification And ChatBuddies.Count > 0 And Not (mnuTools_FakeNick.Checked = True And Not PrevFontName = vbNullString) Then
            If DateDiff("s", LastTyped, Now()) >= 5 Then
                If Not KeyAscii = vbKeyBack Then
                    objMSN_SB.SendTypingNotification
                End If
                LastTyped = Now
            End If
        End If
    End If
End Sub

Public Sub AddChat(Nick As String, FontName As String, FontColor As Long, FontBold As Boolean, FontItalic As Boolean, FontStrikethru As Boolean, FontUnderline As Boolean, Message As String)
    'Add nick
    Call AddRtfText("MS Sans Serif", RGB(60, 60, 60), False, False, False, False, 0, IIf(mnuTools_TimeStamp.Checked, "[" & Time$ & "] ", "") & Nick & " :")
    'Add message
    Call AddRtfText(FontName, FontColor, FontBold, FontItalic, FontStrikethru, FontUnderline, 10, Message)
End Sub

Public Sub UpdateBuddies()
    On Error Resume Next

    If ChatBuddies.Count = 0 Or ChatBuddies.Count = 1 Then
        If ChatBuddies.Count = 1 Then
            Me.Caption = GetCustomNick(ChatBuddies(1).Item("email"), ChatBuddies(1).Item("nick")) & " - Conversation"
        End If
        lblBuddies.Caption = objMSN_SB.Contact
        Call RefreshBuddyDP
        mnuFile_SendAFileOrPhoto.Enabled = True
        mnuActions_SendAFileOrPhoto.Enabled = True
    Else
        mnuFile_SendAFileOrPhoto.Enabled = False
        mnuActions_SendAFileOrPhoto.Enabled = False
        Dim i As Integer
        Me.Caption = GetCustomNick(ChatBuddies(1).Item("email"), ChatBuddies(1).Item("nick"))
        lblBuddies.Caption = ChatBuddies(1).Item("email")
        For i = 2 To ChatBuddies.Count
            Me.Caption = Me.Caption & ", " & GetCustomNick(ChatBuddies(i).Item("email"), ChatBuddies(i).Item("nick"))
            lblBuddies.Caption = lblBuddies.Caption & ", " & ChatBuddies(i).Item("email")
        Next
        Me.Caption = Me.Caption & " - Conversation"
    End If
    
    Call ResizeScrollLabelSet(picBuddies, lblBuddies, vsBuddies)
    Call vsBuddies_Change
    
    Call FlashWindowEx(Me.hwnd)
End Sub

Public Sub Comment(Text As String, Optional Color As Long = 3947580, Optional Log As Boolean = True)
    Call AddRtfText("MS Sans Serif", Color, False, False, False, False, 0, "----" & vbCrLf & Text & vbCrLf & "----")
    Call LogChat(objMSN_SB.Contact, "----" & vbCrLf & Text & vbCrLf & "----")
End Sub

Private Sub AddRtfText(FontName As String, FontColor As Long, FontBold As Boolean, FontItalic As Boolean, FontStrikethru As Boolean, FontUnderline As Boolean, Indent As Integer, Text As String)
    On Error Resume Next
    
    txtChat.SelStart = Len(txtChat.Text)
    txtChat.Locked = False
    If Not txtChat.Text = vbNullString Then
        txtChat.SelText = vbCrLf
    End If
    
    txtChat.SelIndent = Indent
    
    txtChat.SelFontName = FontName
    txtChat.SelFontSize = txtMessage.FontSize
    txtChat.SelColor = FontColor
    txtChat.SelBold = FontBold
    txtChat.SelItalic = FontItalic
    txtChat.SelStrikeThru = FontStrikethru
    txtChat.SelUnderline = FontUnderline
    
    If ShowEmoticons Then
        Dim i As Integer, j As Integer, EmoticonCodeLen As String, TempCpText As String, LastEmoticon  As String
        
        TempCpText = Clipboard.GetText
        
        For i = 1 To Len(Text)
            For j = 0 To UBound(Emoticons)
                EmoticonCodeLen = Len(Emoticons(j, 0))
                If StrComp(Mid$(Text, i, EmoticonCodeLen), Emoticons(j, 0), vbTextCompare) = 0 Then
                    i = i + EmoticonCodeLen - 1
                    If Not (StrComp(LastEmoticon, Emoticons(j, 0), vbTextCompare) = 0) Then
                        Clipboard.Clear
                        Clipboard.SetData frmMain.imglstEmoticons.ListImages(Val(Emoticons(j, 1))).Picture
                    ElseIf EmoticonFloodControl Then
                        Exit For
                    End If
                    SendMessage txtChat.hwnd, WM_PASTE, 0, 0
                    LastEmoticon = Emoticons(j, 0)
                    txtChat.SelFontName = FontName
                    txtChat.SelColor = FontColor
                    txtChat.SelBold = FontBold
                    txtChat.SelItalic = FontItalic
                    txtChat.SelStrikeThru = FontStrikethru
                    txtChat.SelUnderline = FontUnderline
                    Exit For
                End If
            Next
            If j > UBound(Emoticons) Then
                txtChat.SelText = Mid$(Text, i, 1)
            End If
        Next
        
        Clipboard.SetText TempCpText
    Else
        txtChat.SelText = Text
    End If
    
    txtChat.Locked = True
    txtChat.SelStart = Len(txtChat.Text)
End Sub

Private Sub txtMessage_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call Form_OLEDragDrop(Data, Effect, Button, Shift, x, y)
End Sub

Private Sub vsBuddies_Change()
    lblBuddies.Top = -(vsBuddies.Value * picBuddies.TextHeight(lblBuddies.Caption))
End Sub

Private Sub vsBuddyStatus_Change()
    lblBuddyStatus.Top = -(vsBuddyStatus.Value * picBuddyStatus.TextHeight(lblBuddyStatus.Caption))
End Sub

Private Sub LoadTextStyle(File As String)
    On Error GoTo Handler
    Set TextStyler = New Collection
    
    Dim FileNum As Integer, strData As String
    
    FileNum = FreeFile()
    Open File For Input As FileNum
    
    Do Until EOF(FileNum)
        Line Input #FileNum, strData
        If Not Left$(strData, 1) = "'" Then
            TextStyler.Add strData
        End If
    Loop
Handler:
    Close FileNum
End Sub

Public Sub BackupMsgFont()
    With txtMessage
        If PrevFontName = vbNullString Then
            PrevFontName = txtMessage.FontName
        End If
        If PrevFontColor = vbNullString Then
            PrevFontColor = txtMessage.ForeColor
        End If
        If PrevFontBold = vbNullString Then
            PrevFontBold = txtMessage.FontBold
        End If
        If PrevFontItalic = vbNullString Then
            PrevFontItalic = txtMessage.FontItalic
        End If
        If PrevFontStrikethru = vbNullString Then
            PrevFontStrikethru = txtMessage.FontStrikethru
        End If
        If PrevFontUnderline = vbNullString Then
            PrevFontUnderline = txtMessage.FontUnderline
        End If
    End With
End Sub

Public Sub RestoreMsgFont()
    On Error Resume Next
    
    With txtMessage
        txtMessage.FontName = PrevFontName
        txtMessage.ForeColor = Val(PrevFontColor)
        txtMessage.FontBold = PrevFontBold
        txtMessage.FontItalic = PrevFontItalic
        txtMessage.FontStrikethru = PrevFontStrikethru
        txtMessage.FontUnderline = PrevFontUnderline
        
        PrevFontName = vbNullString
        PrevFontColor = vbNullString
        PrevFontBold = vbNullString
        PrevFontItalic = vbNullString
        PrevFontStrikethru = vbNullString
        PrevFontUnderline = vbNullString
    End With
End Sub

Public Sub Imitate(Email As String)
    Call BackupMsgFont
    If InCollection(ChatBuddies(Email), "fontname") Then
        With txtMessage
            txtMessage.FontName = ChatBuddies(Email).Item("fontname")
            txtMessage.ForeColor = ChatBuddies(Email).Item("fontcolor")
            txtMessage.FontBold = ChatBuddies(Email).Item("fontbold")
            txtMessage.FontItalic = ChatBuddies(Email).Item("fontitalic")
            txtMessage.FontStrikethru = ChatBuddies(Email).Item("fontstrikethru")
            txtMessage.FontUnderline = ChatBuddies(Email).Item("fontunderline")
        End With
        mnuTools_FakeNick.Tag = ChatBuddies(Email).Item("nick")
        mnuTools_FakeNick.Checked = True
    End If
End Sub

Public Sub CleanDpTransfers()
    Dim i As Integer
    If Not DpTransfers.Count = 0 Then
        For i = 1 To DpTransfers.Count
            If DpTransfers(i).Item("hWnd") = Me.hwnd Then
                Close #DpTransfers(i).Item("fptr")
                DpTransfers.Remove i
                Call CleanDpTransfers
                Exit Sub
            End If
        Next
    End If
End Sub

Public Sub OfferDP()
    If SendDisplayPic Then
        If Not ChatBuddies.Count = 0 Then
            Dim DpId As String
            DpId = GetSettingX("Display Pics", frmMain.objMSN_NS.Login)
            If Not DpId = vbNullString And FileExists(App.Path & "\Display Pics\" & frmMain.objMSN_NS.Login & ".dat") Then
                objMSN_SB.SendCustomMessage "gm-displaypic", "id: " & DpId, vbNullString
            End If
        End If
    End If
End Sub

Public Sub RefreshMyDP()
    If SendDisplayPic Then
        Call LoadDP(frmMain.objMSN_NS.Login, imgMyDP)
    Else
        imgMyDP.Visible = False
    End If
    imgShowHideMyDP.Visible = imgMyDP.Visible
    Call Form_Resize
End Sub

Public Sub RefreshBuddyDP(Optional ForceDisplay As Boolean)
    If InCollection(ChatBuddies, objMSN_SB.Contact) Then
        If InCollection(ChatBuddies(objMSN_SB.Contact), "dpoffered") Then
            If ReceiveDisplayPic And ChatBuddies(objMSN_SB.Contact).Item("dpoffered") Then
                If imgBuddyDP.Width = 1 And Not ForceDisplay Then
                    imgBuddyDP.Visible = True
                Else
                    Call LoadDP(objMSN_SB.Contact, imgBuddyDP)
                    If ForceDisplay Then
                        imgBuddyDP.Width = 120
                        imgBuddyDP.Height = 120
                        imgBuddyDP.BorderStyle = vbFixedSingle
                        SaveSettingX "App Settings\" & frmMain.objMSN_NS.Login & "\Show DP", objMSN_SB.Contact, True
                    End If
                End If
            Else
                imgBuddyDP.Visible = False
            End If
        Else
            imgBuddyDP.Visible = False
        End If
    Else
        imgBuddyDP.Visible = False
    End If
    imgShowHideBuddyDP.Visible = imgBuddyDP.Visible
    Call Form_Resize
End Sub

Private Sub SendQueMessages()
    Dim i As Integer
    If Not MessageQue.Count = 0 Then
        For i = 1 To MessageQue.Count
            SendMsg Me, MessageQue(1), False
            MessageQue.Remove 1
        Next
    End If
    If Not FileQue.Count = 0 Then
        Dim FileParams() As String
        For i = 1 To FileQue.Count
            FileParams = Split(FileQue(1), "|")
            SendFile Me, CStr(FileParams(1)), CStr(FileParams(0)), Val(FileParams(2))
            FileQue.Remove 1
        Next
    End If
End Sub

Public Function BuddyNick() As String
    If InCollection(ChatBuddies, objMSN_SB.Contact) Then
        BuddyNick = ChatBuddies(objMSN_SB.Contact).Item("nick")
    Else
        BuddyNick = GetContactAttr(objMSN_SB.Contact, "nick")
    End If
    BuddyNick = GetCustomNick(objMSN_SB.Contact, BuddyNick)
End Function

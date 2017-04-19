VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmContactProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Properties"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5220
   Icon            =   "frmContactProperties.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   330
      Left            =   2760
      TabIndex        =   7
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   3960
      TabIndex        =   8
      Top             =   4320
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4095
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   7223
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmContactProperties.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "imgBuddyInfo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Line1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblEmail"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblStatus"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblGroups"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblCustomNick"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtNick"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtEmail"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtStatus"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtGroups"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtCustomNick"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Comment"
      TabPicture(1)   =   "frmContactProperties.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtComment"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Statistics"
      TabPicture(2)   =   "frmContactProperties.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblLastConversationJoined"
      Tab(2).Control(1)=   "lblLastConversationStarted"
      Tab(2).Control(2)=   "lblLastOnline"
      Tab(2).Control(3)=   "lblLastMessageSent"
      Tab(2).Control(4)=   "lblLastIP"
      Tab(2).ControlCount=   5
      Begin VB.TextBox txtCustomNick 
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox txtComment 
         Height          =   3255
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   600
         Width           =   4455
      End
      Begin VB.TextBox txtGroups 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "[Groups]"
         Top             =   2640
         Width           =   3735
      End
      Begin VB.TextBox txtStatus 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "[Status]"
         Top             =   2280
         Width           =   3735
      End
      Begin VB.TextBox txtEmail 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   "[Email]"
         Top             =   1920
         Width           =   3735
      End
      Begin VB.TextBox txtNick 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   840
         TabIndex        =   1
         Text            =   "[Nick]"
         Top             =   720
         Width           =   3855
      End
      Begin VB.Label lblCustomNick 
         AutoSize        =   -1  'True
         Caption         =   "Custom Nick:"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   1245
         Width           =   945
      End
      Begin VB.Label lblLastIP 
         Caption         =   "Last IP address:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   16
         Top             =   2040
         Width           =   4215
      End
      Begin VB.Label lblLastMessageSent 
         Caption         =   "Last sent message on"
         Height          =   255
         Left            =   -74760
         TabIndex        =   15
         Top             =   1680
         Width           =   4215
      End
      Begin VB.Label lblLastOnline 
         Caption         =   "Last seen online on"
         Height          =   255
         Left            =   -74760
         TabIndex        =   14
         Top             =   600
         Width           =   4215
      End
      Begin VB.Label lblLastConversationStarted 
         Caption         =   "Last started conversation on"
         Height          =   255
         Left            =   -74760
         TabIndex        =   13
         Top             =   960
         Width           =   4215
      End
      Begin VB.Label lblLastConversationJoined 
         Caption         =   "Last joined conversation on"
         Height          =   255
         Left            =   -74760
         TabIndex        =   12
         Top             =   1320
         Width           =   4215
      End
      Begin VB.Label lblGroups 
         AutoSize        =   -1  'True
         Caption         =   "Groups:"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   2640
         Width           =   555
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "Status:"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label lblEmail 
         AutoSize        =   -1  'True
         Caption         =   "Email:"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   1920
         Width           =   420
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000011&
         X1              =   240
         X2              =   4680
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Image imgBuddyInfo 
         Height          =   480
         Left            =   240
         Picture         =   "frmContactProperties.frx":0060
         Top             =   600
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmContactProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If InCollection(ContactList, Me.Tag) Then
        If Not txtNick.Text = GetContactAttr(Me.Tag, "nick") Then
            frmMain.objMSN_NS.RenameContact Me.Tag, txtNick.Text
        End If
    End If
    If Not txtCustomNick.Text = GetBuddyCustomNick(Me.Tag) Then
        SetBuddyCustomNick Me.Tag, txtCustomNick.Text
    End If
    If Not txtComment.Text = GetBuddyComment(Me.Tag) Then
        SetBuddyComment Me.Tag, txtComment.Text
    End If
    Unload Me
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
    If Not Transparency = 0 Then
        SetTransparency Me, Transparency
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    LastActive = Timer
End Sub

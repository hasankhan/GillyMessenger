VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   3765
   Icon            =   "frmSettings.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   274
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   251
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3600
      Width           =   975
   End
   Begin VB.Frame fmGeneral 
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   360
      TabIndex        =   15
      Top             =   600
      Width           =   3015
      Begin VB.CheckBox chkShowIMWindowOnMsg 
         Caption         =   "Show IM Window On &Message"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   2400
         Width           =   2775
      End
      Begin VB.CheckBox chkStartWithWindows 
         Caption         =   "Start With &Windows"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   2130
         Width           =   2175
      End
      Begin VB.CheckBox chkTypingNotify 
         Caption         =   "Send &Typing Notification"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1860
         Width           =   2175
      End
      Begin VB.TextBox txtStatusLogDir 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Top             =   960
         Width           =   2535
      End
      Begin VB.CheckBox chkPopups 
         Caption         =   "Show &Popups"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtChatLogDir 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   2535
      End
      Begin VB.CheckBox chkEmoticons 
         Caption         =   "Use &Emoticons"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1590
         Width           =   1815
      End
      Begin VB.Label lblStatusLogDir 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status Logs Directory : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lblChatLogDir 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Chat Logs Directory : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.Frame fmHistory 
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   360
      TabIndex        =   17
      Top             =   600
      Visible         =   0   'False
      Width           =   3015
      Begin VB.CommandButton cmdClearContactComments 
         Caption         =   "Clear &Comments"
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
         Height          =   345
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1200
         Width           =   2055
      End
      Begin VB.CommandButton cmdResetAppSettings 
         Caption         =   "Reset &App Settings"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1680
         Width           =   2055
      End
      Begin VB.CommandButton cmdResetServerSettings 
         Caption         =   "Reset &Server Settings"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2160
         Width           =   2055
      End
      Begin VB.CommandButton cmdClearLoginCache 
         Caption         =   "Clear &Login Cache"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmdClearIgnoreList 
         Caption         =   "Clear &Ignore List"
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
         Height          =   345
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   720
         Width           =   2055
      End
   End
   Begin VB.Frame fmSignIn 
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   360
      TabIndex        =   11
      Top             =   600
      Visible         =   0   'False
      Width           =   3015
      Begin VB.ComboBox cmbInitialStatus 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmSettings.frx":000C
         Left            =   240
         List            =   "frmSettings.frx":0028
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1320
         Width           =   2535
      End
      Begin VB.ComboBox cmbSignInMode 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmSettings.frx":0081
         Left            =   240
         List            =   "frmSettings.frx":008B
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label lblInitialStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Initial Status : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1080
         Width           =   1050
      End
      Begin VB.Label lblSignInMode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SignIn Mode : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   1035
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3375
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   5953
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Sign In"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Cache"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdClearContactComments_Click()
On Error Resume Next
ResetCollection BuddyComment
DeleteSetting "Gilly Messenger", "Comments\" & Login
MsgBox "Contact comments cleared.", vbInformation
End Sub

Private Sub cmdClearIgnoreList_Click()
On Error Resume Next
For X = 1 To BuddyIgnore.Count
    Unignore BuddyIgnore(1)
Next
MsgBox "Ignore list cleared.", vbInformation
End Sub

Private Sub cmdClearLoginCache_Click()
On Error Resume Next
DeleteSetting "Gilly Messenger", "Login Cache"
frmSignIn.cmbLogin.Clear
If SignedIn = False Then
    Login = vbNullString
End If
MsgBox "Login cache cleared.", vbInformation
End Sub

Private Sub cmdResetServerSettings_Click()
On Error Resume Next
DeleteSetting "Gilly Messenger", "Server Settings"
frmMain.wskMSN.RemoteHost = "messenger.hotmail.com"
frmMain.wskMSN.RemotePort = 1863
MsgBox "Server settings reset.", vbInformation
End Sub

Private Sub cmdOK_Click()
If Right$(txtChatLogDir.Text, 1) <> "\" Then
    txtChatLogDir.Text = txtChatLogDir.Text & "\"
End If
ChatLogDir = txtChatLogDir.Text
SaveSetting "Gilly Messenger", "App Settings", "ChatLog Folder", ChatLogDir
Popups = chkPopups.Value
StatusLogDir = txtStatusLogDir.Text
SaveSetting "Gilly Messenger", "App Settings", "StatusLog Folder", StatusLogDir
SaveSetting "Gilly Messenger", "App Settings", "Popups", Popups
TypingNotify = chkTypingNotify.Value
SaveSetting "Gilly Messenger", "App Settings", "Typing Notification", TypingNotify
ShowIMWindowOnMsg = chkShowIMWindowOnMsg.Value
SaveSetting "Gilly Messenger", "App Settings", "Show IMWindow On Message", chkShowIMWindowOnMsg.Value
UseEmoticons = chkEmoticons.Value
InitialStatus = cmbInitialStatus.ListIndex
SaveSetting "Gilly Messenger", "Sign In", "Status", InitialStatus
SignInMode = cmbSignInMode.List(cmbSignInMode.ListIndex)
SaveSetting "Gilly Messenger", "Sign In", "Mode", SignInMode
If chkStartWithWindows.Value = vbChecked Then
    WriteRegKey "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run\Gilly Messenger", """" & App.Path & "\" & App.EXEName & ".exe"" /startup"
Else
    DeleteRegKey "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run\Gilly Messenger"
End If
Unload Me
End Sub

Private Sub cmdResetAppSettings_Click()
On Error Resume Next
DeleteSetting "Gilly Messenger", "App Settings"
Call LoadAppSettings
Call Form_Load
MsgBox "App settings reset.", vbInformation
End Sub

Private Sub Form_Load()
txtStatusLogDir.Text = StatusLogDir
txtChatLogDir.Text = ChatLogDir
If ShowIMWindowOnMsg = True Then
    chkShowIMWindowOnMsg.Value = vbChecked
Else
    chkShowIMWindowOnMsg.Value = vbUnchecked
End If
If TypingNotify = True Then
    chkTypingNotify.Value = vbChecked
Else
    chkTypingNotify.Value = vbUnchecked
End If
If Popups = True Then
    chkPopups.Value = vbChecked
Else
    chkPopups.Value = vbUnchecked
End If
If UseEmoticons = True Then
    chkEmoticons.Value = vbChecked
Else
    chkEmoticons.Value = vbUnchecked
End If
If SignInMode = "Online" Then
    cmbSignInMode.ListIndex = 1
ElseIf SignInMode = "Complete" Then
    cmbSignInMode.ListIndex = 0
Else
    cmbSignInMode.ListIndex = 0
End If
cmbInitialStatus.ListIndex = InitialStatus
If ReadRegKey("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run\Gilly Messenger") = """" & App.Path & "\" & App.EXEName & ".exe"" /startup" Then
    chkStartWithWindows.Value = vbChecked
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Me.Hide
End Sub

Private Sub TabStrip1_Click()
If TabStrip1.SelectedItem.Caption = "General" Then
    fmGeneral.Visible = True
    fmSignIn.Visible = False
    fmHistory.Visible = False
ElseIf TabStrip1.SelectedItem.Caption = "Sign In" Then
    fmGeneral.Visible = False
    fmSignIn.Visible = True
    fmHistory.Visible = False
ElseIf TabStrip1.SelectedItem.Caption = "Cache" Then
    fmGeneral.Visible = False
    fmSignIn.Visible = False
    fmHistory.Visible = True
End If
End Sub

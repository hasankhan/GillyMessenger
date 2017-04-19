VERSION 5.00
Begin VB.Form frmRemoteControl 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GM Remote Control"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3750
   Icon            =   "frmRemoteControl.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
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
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
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
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start Server"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtPassword 
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
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
   Begin VB.TextBox txtUsername 
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
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
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
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   795
   End
   Begin VB.Label lblUsername 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Username :"
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
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   825
   End
End
Attribute VB_Name = "frmRemoteControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Me.Hide
End Sub

Private Sub cmdStart_Click()
If cmdStart.Caption = "&Start Server" Then
    If txtUsername.Text = vbNullString Or txtPassword.Text = vbNullString Then
        MsgBox "You must specify username and password.", vbExclamation
        Exit Sub
    End If
    RcUsername = txtUsername.Text
    RcPassword = txtPassword.Text
    RemoteControl = True
    cmdStart.Caption = "&Stop Server"
    cmdUpdate.Enabled = True
    frmMain.lblStatus.Caption = "Remote Control server started."
    Me.Hide
ElseIf cmdStart.Caption = "&Stop Server" Then
    RemoteControl = False
    cmdStart.Caption = "&Start Server"
    cmdUpdate.Enabled = False
    frmMain.lblStatus.Caption = "Remote Control server stopped."
    Me.Hide
End If
End Sub

Private Sub cmdUpdate_Click()
If txtUsername.Text = vbNullString Or txtPassword.Text = vbNullString Then
    MsgBox "You must specify username and password.", vbExclamation
    Exit Sub
End If
RcUsername = txtUsername.Text
RcPassword = txtPassword.Text
frmMain.lblStatus.Caption = "Remote Control server updated."
Me.Hide
End Sub

Private Sub Form_Activate()
If RemoteControl = True Then
    cmdUpdate.Default = True
Else
    cmdStart.Default = True
End If
txtUsername = RcUsername
txtPassword = RcPassword
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Me.Hide
End Sub

Private Sub txtPassword_GotFocus()
txtPassword.SelStart = 0
txtPassword.SelLength = Len(txtPassword.Text)
End Sub

Private Sub txtUsername_GotFocus()
txtUsername.SelStart = 0
txtUsername.SelLength = Len(txtUsername.Text)
End Sub

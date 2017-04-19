VERSION 5.00
Begin VB.Form frmSignIn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sign In"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3630
   Icon            =   "frmSignIn.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   114
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   242
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbLogin 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   960
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   2535
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   720
      Width           =   2535
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
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1125
      Width           =   975
   End
   Begin VB.CommandButton cmdSignIn 
      Caption         =   "&Sign In"
      Default         =   -1  'True
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
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1125
      Width           =   975
   End
   Begin VB.Label lblPassword 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Left            =   180
      TabIndex        =   5
      Top             =   720
      Width           =   705
   End
   Begin VB.Label lblLogin 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
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
      Left            =   480
      TabIndex        =   4
      Top             =   240
      Width           =   405
   End
End
Attribute VB_Name = "frmSignIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fResult As Long, SearchCombo As Boolean

Private Sub cmbLogin_Change()
If SearchCombo = False Then Exit Sub
X = Len(cmbLogin.Text)
fResult = SendMessage(cmbLogin.hwnd, CB_FINDSTRING, -1, ByVal cmbLogin.Text)
If fResult <> -1 Then
    SendMessage cmbLogin.hwnd, CB_SELECTSTRING, -1, ByVal cmbLogin.Text
    cmbLogin.SelStart = X
    cmbLogin.SelLength = Len(cmbLogin.Text)
End If
If cmbLogin.Text = vbNullString Or txtPassword.Text = vbNullString Then
    cmdSignIn.Enabled = False
Else
    cmdSignIn.Enabled = True
End If
End Sub

Private Sub cmbLogin_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyBack Then
    SearchCombo = False
Else
    SearchCombo = True
End If
End Sub

Private Sub cmbLogin_LostFocus()
If InStr(cmbLogin.Text, "@") = 0 And Len(cmbLogin.Text) > 0 Then
    cmbLogin.Text = cmbLogin & "@hotmail.com"
End If
End Sub

Private Sub cmdCancel_Click()
Me.Visible = False
txtPassword.Text = vbNullString
End Sub

Private Sub cmdSignIn_Click()
If SignedIn = True Then
    If MsgBox("You are already singed in." & vbCrLf & "Do you want to sign out?", vbYesNo Or vbQuestion) = vbNo Then
        Me.Visible = False
        Exit Sub
    End If
End If
Me.Visible = False
cmbLogin.Text = LCase$(cmbLogin.Text)
Call SignIn(cmbLogin.Text, txtPassword.Text)
txtPassword.Text = vbNullString
End Sub

Private Sub Form_Activate()
On Error Resume Next
If cmbLogin.Text <> vbNullString Then
    txtPassword.SetFocus
Else
    cmbLogin.SetFocus
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Me.Visible = False
    txtPassword.Text = vbNullString
End If
End Sub

Private Sub cmbLogin_GotFocus()
cmbLogin.SelStart = 0
cmbLogin.SelLength = Len(cmbLogin.Text)
End Sub

Private Sub Form_Load()
cmbLogin.Text = Login
txtPassword.Text = vbNullString
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = vbFormControlMenu Then
    Cancel = 1
    Me.Visible = False
    txtPassword.Text = vbNullString
End If
End Sub

Private Sub txtPassword_Change()
If cmbLogin.Text = vbNullString Or txtPassword.Text = vbNullString Then
    cmdSignIn.Enabled = False
Else
    cmdSignIn.Enabled = True
End If
End Sub

Private Sub txtPassword_GotFocus()
txtPassword.SelStart = 0
txtPassword.SelLength = Len(txtPassword.Text)
End Sub

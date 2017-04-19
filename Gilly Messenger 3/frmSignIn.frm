VERSION 5.00
Begin VB.Form frmSignIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sign in to .NET Messenger Service"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4080
   Icon            =   "frmSignIn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   136
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   272
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkSavePassword 
      Caption         =   "&Save Password"
      Height          =   195
      Left            =   225
      TabIndex        =   3
      Top             =   1635
      Width           =   1455
   End
   Begin VB.ComboBox cmbInitialStatus 
      Height          =   315
      ItemData        =   "frmSignIn.frx":000C
      Left            =   1260
      List            =   "frmSignIn.frx":0028
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1080
      Width           =   2535
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
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
      Height          =   330
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
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
      Height          =   330
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1560
      Width           =   975
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
      Left            =   1260
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   2535
   End
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
      Left            =   1260
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblInitialStatus 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Initial Status"
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
      Left            =   225
      TabIndex        =   8
      Top             =   1140
      Width           =   915
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
      Left            =   225
      TabIndex        =   7
      Top             =   180
      Width           =   405
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
      Left            =   225
      TabIndex        =   6
      Top             =   645
      Width           =   705
   End
End
Attribute VB_Name = "frmSignIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SearchCombo As Boolean, CachePwd As Boolean
Private LoginCache() As String

Private Sub cmbLogin_Change()
    On Error Resume Next
    
    If SearchCombo Then
        Dim i As Integer, fResult As Long
        i = Len(cmbLogin.Text)
        fResult = SendMessage(cmbLogin.hwnd, CB_FINDSTRING, -1, ByVal cmbLogin.Text)
        If Not fResult = -1 Then
            SendMessage cmbLogin.hwnd, CB_SELECTSTRING, -1, ByVal cmbLogin.Text
            cmbLogin.SelStart = i
            cmbLogin.SelLength = Len(cmbLogin.Text)
            cmbInitialStatus.ListIndex = Val(Left$(LoginCache(fResult, 1), 1))
            Dim Temp As String
            Temp = Mid$(LoginCache(fResult, 1), 2)
            If Not Temp = vbNullString Then
                Temp = XorDecrypt(Temp, cmbLogin.Text)
                If Not txtPassword.Text = Temp Then
                    CachePwd = False
                    txtPassword.Text = Temp
                End If
                chkSavePassword.Value = vbChecked
            Else
                Call RestoreManualPwd
            End If
        Else
            Call RestoreManualPwd
        End If
    Else
        Call RestoreManualPwd
    End If
    
    If cmbLogin.Text = vbNullString Or txtPassword.Text = vbNullString Then
        cmdOK.Enabled = False
    Else
        cmdOK.Enabled = True
    End If
End Sub

Private Sub cmbLogin_Click()
    SearchCombo = True
    Call cmbLogin_Change
End Sub

Private Sub cmbLogin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        SearchCombo = False
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
    Unload Me
End Sub

Private Sub cmdOK_Click()
    cmbLogin.Text = LCase$(cmbLogin.Text)
    InitialStatus = cmbInitialStatus.ListIndex
    frmMain.tmrReconnect.Tag = 0
    Call SignIn(cmbLogin.Text, txtPassword.Text)
    SavePassword = (chkSavePassword.Value = vbChecked)
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
    frmSignIn.cmbInitialStatus.ListIndex = 0
    
    On Error GoTo Handler
    
    LoginCache = GetAllSettings("Gilly Messenger", "Login Cache")
    If Not ArraySize(LoginCache) = -1 Then
        Dim i As Integer
        For i = 0 To UBound(LoginCache)
            cmbLogin.AddItem LoginCache(i, 0)
        Next
    End If
    
Handler:
    cmbLogin.Text = frmMain.objMSN_NS.Login
    If SavePassword Then
        txtPassword.Text = frmMain.objMSN_NS.Password
        chkSavePassword.Value = vbChecked
    Else
        chkSavePassword.Value = vbUnchecked
    End If
    
    CachePwd = True
    
    If Not Transparency = 0 Then
        SetTransparency Me, Transparency
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    LastActive = Timer
End Sub

Private Sub txtPassword_Change()
    If CachePwd Then
        txtPassword.Tag = txtPassword.Text
    End If
    If cmbLogin.Text = vbNullString Or txtPassword.Text = vbNullString Then
        cmdOK.Enabled = False
    Else
        cmdOK.Enabled = True
    End If
End Sub

Private Sub txtPassword_GotFocus()
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)
End Sub

Private Sub RestoreManualPwd()
    txtPassword.Text = txtPassword.Tag
    chkSavePassword.Value = vbUnchecked
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    CachePwd = True
End Sub

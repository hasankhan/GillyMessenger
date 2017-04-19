VERSION 5.00
Begin VB.Form frmRemoteControl 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Remote Control"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3645
   Icon            =   "frmRemoteControl.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   202
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   243
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdViewSessions 
      Caption         =   "&View Sessions"
      Height          =   330
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "De&lete"
      Enabled         =   0   'False
      Height          =   330
      Left            =   2280
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Enabled         =   0   'False
      Height          =   330
      Left            =   2280
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CheckBox chkShellCommands 
      Caption         =   "Shell &Commands"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CheckBox chkMessengerControl 
      Caption         =   "&Messenger Control"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CheckBox chkDirectoryBrowsing 
      Caption         =   "&Directory Browsing"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox txtPassword 
      Enabled         =   0   'False
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   600
      Width           =   2535
   End
   Begin VB.ComboBox cmbAccount 
      Height          =   315
      ItemData        =   "frmRemoteControl.frx":000C
      Left            =   960
      List            =   "frmRemoteControl.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start &Server"
      Height          =   330
      Left            =   2280
      TabIndex        =   8
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtAccount 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000011&
      X1              =   8
      X2              =   232
      Y1              =   152
      Y2              =   152
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lblAccount 
      AutoSize        =   -1  'True
      Caption         =   "Account:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   645
   End
End
Attribute VB_Name = "frmRemoteControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbAccount_Click()
    txtPassword.Enabled = True
    chkDirectoryBrowsing.Enabled = True
    chkMessengerControl.Enabled = True
    chkShellCommands.Enabled = True
    cmdDelete.Enabled = True
    cmdUpdate.Enabled = True
    
    If cmbAccount.ListIndex = cmbAccount.ListCount - 1 And cmbAccount.Visible Then
        cmbAccount.Visible = False
        txtAccount.Visible = True
        txtAccount.Text = vbNullString
        txtPassword.Text = vbNullString
        chkDirectoryBrowsing.Value = vbChecked
        chkMessengerControl.Value = vbChecked
        chkShellCommands.Value = vbChecked
        cmdDelete.Caption = "Ca&ncel"
        cmdUpdate.Caption = "&Add"
        txtAccount.SetFocus
    Else
        On Error Resume Next
        
        txtPassword.Text = RC_Accounts(cmbAccount.Text).Item("password")
        chkDirectoryBrowsing.Value = IIf(RC_Accounts(cmbAccount.Text).Item("dirBrowsing"), vbChecked, vbUnchecked)
        chkMessengerControl.Value = IIf(RC_Accounts(cmbAccount.Text).Item("msgrControl"), vbChecked, vbUnchecked)
        chkShellCommands.Value = IIf(RC_Accounts(cmbAccount.Text).Item("shellCommands"), vbChecked, vbUnchecked)
        cmdDelete.Caption = "De&lete"
        cmdUpdate.Caption = "&Update"
    End If
End Sub

Private Sub cmdDelete_Click()
    On Error Resume Next
    
    Select Case cmdDelete.Caption
    Case "De&lete"
        DeleteSetting "Gilly Messenger", "RC Accounts", cmbAccount.Text
        RC_Accounts.Remove cmbAccount.Text
        cmbAccount.RemoveItem cmbAccount.ListIndex
        txtPassword.Text = vbNullString
        chkDirectoryBrowsing.Value = vbUnchecked
        chkMessengerControl.Value = vbUnchecked
        chkShellCommands.Value = vbUnchecked
        cmbAccount.Visible = False
        cmbAccount.ListIndex = 0
        Call cmbAccount_Click
        cmbAccount.Visible = True
    Case "Ca&ncel"
        txtAccount.Visible = False
        cmbAccount.ListIndex = 0
        Call cmbAccount_Click
        cmbAccount.Visible = True
        cmbAccount.SetFocus
    End Select
End Sub

Private Sub cmdStart_Click()
    Select Case cmdStart.Caption
    Case "Start &Server"
        If Not RC_Accounts.Count = 0 Then
            RemoteControl = True
            cmdStart.Caption = "Stop &Server"
            cmdViewSessions.Visible = True
        Else
            MsgBox "No account has been setup for use with Remote Control.", vbExclamation
            If cmbAccount.Visible Then
                cmbAccount.SetFocus
            End If
        End If
    Case "Stop &Server"
        RemoteControl = False
        cmdViewSessions.Visible = False
    End Select
End Sub

Private Sub cmdUpdate_Click()
    If txtAccount.Text = vbNullString And Not cmbAccount.Visible Then
        MsgBox "You must specify a username for the account.", vbExclamation
        txtAccount.SetFocus
    ElseIf InStr(txtAccount.Text, " ") > 0 Then
        MsgBox "Account name can not contain space character.", vbExclamation
        txtAccount.SelStart = 0
        txtAccount.SelLength = Len(txtAccount.Text)
        txtAccount.SetFocus
    ElseIf txtPassword.Text = vbNullString Then
        MsgBox "You must specify a password for the account.", vbExclamation
        txtPassword.SetFocus
    ElseIf InStr(txtPassword.Text, " ") > 0 Then
        MsgBox "Password can not contain space character.", vbExclamation
        txtPassword.SelStart = 0
        txtPassword.SelLength = Len(txtPassword.Text)
        txtPassword.SetFocus
    Else
        Dim NewAccount As Collection
        Set NewAccount = New Collection

        Select Case cmdUpdate.Caption
        Case "&Update"
            NewAccount.Add cmbAccount.Text, "login"
            NewAccount.Add txtPassword.Text, "password"
            NewAccount.Add (chkDirectoryBrowsing.Value = vbChecked), "dirBrowsing"
            NewAccount.Add (chkMessengerControl.Value = vbChecked), "msgrControl"
            NewAccount.Add (chkShellCommands.Value = vbChecked), "shellCommands"
            
            SetCollectionItem RC_Accounts, cmbAccount.Text, NewAccount
            
            SaveSetting "Gilly Messenger", "RC Accounts", cmbAccount.Text, XorEncrypt(txtPassword.Text, cmbAccount.Text) & " " & (chkDirectoryBrowsing.Value = vbChecked) & " " & (chkMessengerControl.Value = vbChecked) & " " & (chkShellCommands.Value = vbChecked)
            cmbAccount.SetFocus
        Case "&Add"
            If InCollection(RC_Accounts, txtAccount.Text) Then
                MsgBox "Username " & txtAccount.Text & " already exists."
                txtAccount.SelStart = 0
                txtAccount.SelLength = Len(txtAccount.Text)
                txtAccount.SetFocus
            Else
                NewAccount.Add txtAccount.Text, "login"
                NewAccount.Add txtPassword.Text, "password"
                NewAccount.Add (chkDirectoryBrowsing.Value = vbChecked), "dirBrowsing"
                NewAccount.Add (chkMessengerControl.Value = vbChecked), "msgrControl"
                NewAccount.Add (chkShellCommands.Value = vbChecked), "shellCommands"
                
                SetCollectionItem RC_Accounts, txtAccount.Text, NewAccount
                
                SaveSetting "Gilly Messenger", "RC Accounts", txtAccount.Text, XorEncrypt(txtPassword.Text, txtAccount.Text) & " " & (chkDirectoryBrowsing.Value = vbChecked) & " " & (chkMessengerControl.Value = vbChecked) & " " & (chkShellCommands.Value = vbChecked)
                cmbAccount.AddItem txtAccount.Text, cmbAccount.ListCount - 1
                cmbAccount.ListIndex = cmbAccount.ListCount - 2
                Call cmbAccount_Click
                cmbAccount.Visible = True
                cmbAccount.SetFocus
            End If
        End Select
        Set NewAccount = Nothing
    End If
End Sub

Private Sub cmdViewSessions_Click()
    On Error Resume Next
    
    frmRcSessions.Show vbModal, Me
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
    Dim i As Integer
    For i = 1 To RC_Accounts.Count
        cmbAccount.AddItem RC_Accounts(i).Item("login")
    Next
    cmbAccount.AddItem "<New Account>"
    If RemoteControl Then
        cmdStart.Caption = "Stop &Server"
        cmdViewSessions.Visible = True
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    LastActive = Timer
End Sub

Private Sub txtAccount_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then
        KeyAscii = Asc("_")
    End If
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then
        KeyAscii = Asc("_")
    End If
End Sub

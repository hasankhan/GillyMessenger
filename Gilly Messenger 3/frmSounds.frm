VERSION 5.00
Begin VB.Form frmSounds 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sounds"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   Icon            =   "frmSounds.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   344
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkAlert 
      Caption         =   "Alert"
      Height          =   285
      Left            =   240
      TabIndex        =   15
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox txtAlert 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1920
      Width           =   2655
   End
   Begin VB.CommandButton cmdBrowse6 
      Caption         =   "<"
      Height          =   285
      Left            =   4080
      TabIndex        =   17
      Top             =   1920
      Width           =   255
   End
   Begin VB.CheckBox chkTyping 
      Caption         =   "Typing"
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox txtTyping 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   840
      Width           =   2655
   End
   Begin VB.CommandButton cmdBrowse5 
      Caption         =   "<"
      Height          =   285
      Left            =   4080
      TabIndex        =   8
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   3240
      TabIndex        =   22
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   330
      Left            =   2040
      TabIndex        =   23
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdBrowse4 
      Caption         =   "<"
      Height          =   285
      Left            =   4080
      TabIndex        =   14
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1560
      Width           =   2655
   End
   Begin VB.CheckBox chkEmail 
      Caption         =   "Email"
      Height          =   285
      Left            =   240
      TabIndex        =   12
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdBrowse3 
      Caption         =   "<"
      Height          =   285
      Left            =   4080
      TabIndex        =   11
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox txtMessage 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1200
      Width           =   2655
   End
   Begin VB.CheckBox chkMessage 
      Caption         =   "Message"
      Height          =   285
      Left            =   240
      TabIndex        =   9
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdBrowse2 
      Caption         =   "<"
      Height          =   285
      Left            =   4080
      TabIndex        =   5
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox txtOffline 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   2655
   End
   Begin VB.CheckBox chkOffline 
      Caption         =   "Offline"
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdBrowse1 
      Caption         =   "<"
      Height          =   285
      Left            =   4080
      TabIndex        =   2
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox txtOnline 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   2655
   End
   Begin VB.CheckBox chkOnline 
      Caption         =   "Online"
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   330
      Left            =   240
      TabIndex        =   21
      Top             =   4680
      Width           =   1095
   End
   Begin VB.ListBox lstContacts 
      Height          =   1230
      ItemData        =   "frmSounds.frx":000C
      Left            =   240
      List            =   "frmSounds.frx":000E
      TabIndex        =   20
      Top             =   3360
      Width           =   4095
   End
   Begin VB.OptionButton optPlaySoundsExcept 
      Caption         =   "Play sounds for every contact except these in the list"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   3000
      Width           =   4095
   End
   Begin VB.OptionButton optPlaySoundsFor 
      Caption         =   "Play sounds for every contact in the list"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   2640
      Width           =   4095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000011&
      X1              =   16
      X2              =   288
      Y1              =   160
      Y2              =   160
   End
End
Attribute VB_Name = "frmSounds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    Dim strContact As String
    strContact = InputBox("Enter email address of contact.")
    If Not strContact = vbNullString Then
        lstContacts.AddItem strContact
    End If
End Sub

Private Sub cmdBrowse1_Click()
    If Not GetUserFile() = vbNullString Then
        txtOnline.Text = frmMain.CommonDialog.FileName
    End If
End Sub

Private Sub cmdBrowse2_Click()
    If Not GetUserFile() = vbNullString Then
        txtOffline.Text = frmMain.CommonDialog.FileName
    End If
End Sub

Private Sub cmdBrowse3_Click()
    If Not GetUserFile() = vbNullString Then
        txtMessage.Text = frmMain.CommonDialog.FileName
    End If
End Sub

Private Sub cmdBrowse4_Click()
    If Not GetUserFile() = vbNullString Then
        txtEmail.Text = frmMain.CommonDialog.FileName
    End If
End Sub

Private Sub cmdBrowse5_Click()
    If Not GetUserFile() = vbNullString Then
        txtTyping.Text = frmMain.CommonDialog.FileName
    End If
End Sub

Private Sub cmdBrowse6_Click()
    If Not GetUserFile() = vbNullString Then
        txtAlert.Text = frmMain.CommonDialog.FileName
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error Resume Next
    
    boolOnlineSound = (chkOnline.Value = vbChecked)
    strOnlineSound = txtOnline.Text
    SaveSettingX "App Settings", "Online Sound", IIf(boolOnlineSound, "1", "0") & " " & strOnlineSound
    boolOfflineSound = (chkOffline.Value = vbChecked)
    strOfflineSound = txtOffline.Text
    SaveSettingX "App Settings", "Offline Sound", IIf(boolOfflineSound, "1", "0") & " " & strOfflineSound
    boolTypingSound = (chkTyping.Value = vbChecked)
    strTypingSound = txtTyping.Text
    SaveSettingX "App Settings", "Typing Sound", IIf(boolTypingSound, "1", "0") & " " & strTypingSound
    boolMessageSound = (chkMessage.Value = vbChecked)
    strMessageSound = txtMessage.Text
    SaveSettingX "App Settings", "Message Sound", IIf(boolMessageSound, "1", "0") & " " & strMessageSound
    boolEmailSound = (chkEmail.Value = vbChecked)
    strEmailSound = txtEmail.Text
    SaveSettingX "App Settings", "Email Sound", IIf(boolEmailSound, "1", "0") & " " & strEmailSound
    boolAlertSound = (chkAlert.Value = vbChecked)
    strAlertSound = txtAlert.Text
    SaveSettingX "App Settings", "Alert Sound", IIf(boolAlertSound, "1", "0") & " " & strAlertSound
    
    If lstContacts.Enabled And frmMain.objMSN_NS.State = NsState_SignedIn Then
        SoundFilterMode = Not optPlaySoundsFor.Value
        
        DeleteSetting "Gilly Messenger", "Sound Filter\" & frmMain.objMSN_NS.Login
        SaveSettingX "Sound Filter\" & frmMain.objMSN_NS.Login, "Mode", IIf(optPlaySoundsFor.Value, 0, 1)
        
        Set SoundFilter = Nothing
        Set SoundFilter = New Collection
        
        If Not lstContacts.ListCount = 0 Then
            Dim i As Integer
            For i = 0 To lstContacts.ListCount - 1
                SaveSettingX "Sound Filter\" & frmMain.objMSN_NS.Login, CStr(i), lstContacts.list(i)
                SoundFilter.Add lstContacts.list(i), lstContacts.list(i)
            Next
        Else
            SoundFilter.Add "*@*.*", "*@*.*"
        End If
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
    On Error Resume Next
    
    If boolOnlineSound Then
        chkOnline.Value = vbChecked
    End If
    txtOnline.Text = strOnlineSound
    If boolOfflineSound Then
        chkOffline.Value = vbChecked
    End If
    txtOffline.Text = strOfflineSound
    If boolTypingSound Then
        chkTyping.Value = vbChecked
    End If
    txtTyping.Text = strTypingSound
    If boolMessageSound Then
        chkMessage.Value = vbChecked
    End If
    txtMessage.Text = strMessageSound
    If boolEmailSound Then
        chkEmail.Value = vbChecked
    End If
    txtEmail.Text = strEmailSound
    If boolAlertSound Then
        chkAlert.Value = vbChecked
    End If
    txtAlert.Text = strAlertSound
    
    If frmMain.objMSN_NS.State = NsState_SignedIn Then
        If Not SoundFilterMode Then
            optPlaySoundsFor.Value = True
        Else
            optPlaySoundsExcept.Value = True
        End If
        
        Dim i As Integer
        For i = 1 To SoundFilter.Count
            lstContacts.AddItem SoundFilter(i)
        Next
    Else
        optPlaySoundsFor.Enabled = False
        optPlaySoundsExcept.Enabled = False
        lstContacts.Enabled = False
        cmdAdd.Enabled = False
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    LastActive = Timer
End Sub

Private Sub lstContacts_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    If KeyCode = vbKeyDelete Then
        lstContacts.RemoveItem lstContacts.ListIndex
    End If
End Sub

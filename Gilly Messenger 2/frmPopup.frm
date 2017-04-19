VERSION 5.00
Begin VB.Form frmPopup 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1650
   ClientLeft      =   27165
   ClientTop       =   13485
   ClientWidth     =   2775
   ControlBox      =   0   'False
   Icon            =   "frmPopup.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   110
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   185
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer tmrHide 
      Interval        =   10000
      Left            =   840
      Top             =   600
   End
   Begin VB.Image imgClose 
      Height          =   195
      Left            =   2460
      Picture         =   "frmPopup.frx":000C
      Top             =   60
      Width           =   195
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00000000&
      Height          =   1035
      Left            =   30
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   480
      UseMnemonic     =   0   'False
      Width           =   2715
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As Long, Y As Long, WndPos As Long, Activated As Boolean
Public hActiveWnd As Long, hFocusWnd As Long

Private Sub Form_Activate()
On Error Resume Next
If Activated <> True Then
    SetForegroundWindow hActiveWnd
    SetFocusX hFocusWnd
    Activated = True
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    R.Right = Me.ScaleWidth
    R.Bottom = Me.ScaleHeight
    R.Top = 0
    R.Left = 0
    DrawEdge Me.hDC, R, BDR_SUNKENOUTER Or BDR_SUNKENINNER, BF_RECT
    WndPos = PopupHeight
    PopupHeight = PopupHeight - 1605
    If PopupHeight < 1605 Then lblMessage.Tag = "Limit"
    LastPopup = Me.hwnd
    SystemParametersInfo SPI_GETWORKAREA, 0, R, 0
    Me.Left = (R.Right * Screen.TwipsPerPixelX) - Me.Width - (16 * Screen.TwipsPerPixelX)
    Me.Top = WndPos
    For Y = 0 To Me.Height Step 25
        Me.Top = WndPos - Y
        Me.Height = Y
        DoEvents
    Next
End If
End Sub

Private Sub Form_Load()
lblMessage.MouseIcon = frmMain.picSignIn.MouseIcon
End Sub

Private Sub imgClose_Click()
    Call TerminatePopup
End Sub

Private Sub lblMessage_Click()
If Me.Tag = "Email" Then
    Call frmMain.mnuOpenInbox_Click
Else
    StartChat Me.Tag
End If
Call TerminatePopup
End Sub

Private Sub tmrHide_Timer()
tmrHide.Enabled = False
For Y = Me.Height To 0 Step -50
    Me.Top = Me.Top + 50
    Me.Height = Y
    DoEvents
Next
Call TerminatePopup
End Sub

Private Sub TerminatePopup()
On Error Resume Next
If PopupHeight < 1605 Then
    If lblMessage.Tag = "Limit" Then PopupHeight = R.Bottom * Screen.TwipsPerPixelY
ElseIf LastPopup = Me.hwnd Then
    PopupHeight = R.Bottom * Screen.TwipsPerPixelY
End If
Unload Me
End Sub

VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmTransfer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Transfer"
   ClientHeight    =   1350
   ClientLeft      =   4965
   ClientTop       =   4530
   ClientWidth     =   5190
   Icon            =   "frmTransfer.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   90
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   346
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSWinsockLib.Winsock wskFTP 
      Left            =   2400
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "&Accept"
      Height          =   330
      Left            =   4080
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   975
   End
   Begin VB.Timer tmrTimeOut 
      Interval        =   1000
      Left            =   2880
      Top             =   120
   End
   Begin VB.PictureBox picProgressBar 
      AutoRedraw      =   -1  'True
      Height          =   300
      Left            =   120
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   325
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   930
      Width           =   4935
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   0
         Width           =   4575
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   4080
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   495
      Width           =   975
   End
   Begin VB.Label lblTransfer 
      Caption         =   "Receiving from: [Receiving from]"
      Height          =   225
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   3855
   End
   Begin VB.Label lblFileSize 
      Caption         =   "File size: [File size]"
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label lblFileName 
      Caption         =   "File name: [File name]"
      Height          =   225
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Parent As frmChat
Public WithEvents objMSN_FTP As clsMSN_FTP
Attribute objMSN_FTP.VB_VarHelpID = -1
Public Cookie As Double
Private FTP_LastActive As Date

Public Sub cmdAccept_Click()
    Parent.objMSN_SB.AcceptInvitation Cookie, "Request-Data: IP-Address:"
    FTP_LastActive = Now
    cmdAccept.Visible = False
    Call QueScript(Me, "TransferAccepted", ConvArray(frmMain.objMSN_NS.Login, Cookie))
End Sub

Public Sub cmdCancel_Click()
    If cmdAccept.Visible Then
        Parent.objMSN_SB.CancelInvitation Cookie, "REJECT"
    Else
        Parent.objMSN_SB.CancelInvitation Cookie, "TIMEOUT"
    End If
    Parent.Comment "You have cancelled the transfer of """ & objMSN_FTP.File & """", 128
    Call QueScript(Me, "TransferCancelled", ConvArray(frmMain.objMSN_NS.Login, Cookie))
    Call Terminate
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    LastActive = Timer
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    Set objMSN_FTP = New clsMSN_FTP
    objMSN_FTP.Socket = wskFTP
    
    Call DisableCloseButton(Me.hwnd)
    
    Call objMSN_FTP_Progress(0, 0, "0 Kbps")

    If Not Transparency = 0 Then
        SetTransparency Me, Transparency
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    LastActive = Timer
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    objMSN_FTP.Disconnect
    Parent.Invitations.Remove "Cookie " & Cookie
    Set objMSN_FTP = Nothing
End Sub

Private Sub objMSN_FTP_FtpError(Error As String)
    Call TransferError
End Sub

Private Sub objMSN_FTP_Progress(PercentDone As Integer, BytesTransferred As Double, Rate As String)
    On Error Resume Next
    
    picProgressBar.Cls
    GradientFill picProgressBar.hDC, 0, 0, (picProgressBar.ScaleWidth / 100) * PercentDone, picProgressBar.ScaleHeight, "FFFFFF", "C9D3F3", False
    Dim r As RECT
    r.Left = 0
    r.Top = 0
    r.Right = (picProgressBar.ScaleWidth / 100) * PercentDone
    r.Bottom = picProgressBar.ScaleHeight
    DrawEdge picProgressBar.hDC, r, BDR_RAISEDINNER, BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM
    TextOut picProgressBar.hDC, (picProgressBar.ScaleWidth \ 2) - (picProgressBar.TextWidth(PercentDone & "% " & Rate) / 2), 0, PercentDone & "% @ " & Rate, Len(PercentDone & "% @ " & Rate)
    
    FTP_LastActive = Now
End Sub

Private Sub objMSN_FTP_SocketError(Description As String)
    Call TransferError
End Sub

Private Sub objMSN_FTP_StateChanged()
    FTP_LastActive = Now
    Select Case objMSN_FTP.State
    Case FtpState_Connecting
        picProgressBar.BackColor = vbWindowBackground
        lblStatus.BackColor = vbWindowBackground
        lblStatus.Caption = "Connecting..."
    Case FtpState_Connected
        lblStatus.Caption = "Connected"
    Case FtpState_Negotiating
        lblStatus.Caption = "Negotiating..."
    Case FtpState_Transfer
        lblStatus.Visible = False
    End Select
End Sub

Private Sub objMSN_FTP_TransferComplete()
    On Error Resume Next
    
    tmrTimeOut.Enabled = False
    If objMSN_FTP.TransferType = FtpTransferType_Send Then
        Parent.Comment "Transfer of """ & objMSN_FTP.File & """ is complete."
    Else
        Parent.Comment "You have successfully received file """ & objMSN_FTP.File & """ from " & Parent.BuddyNick
    End If
    Call Parent.tmrResetStatus_Timer
    Call QueScript(Me, "TransferComplete", ConvArray(Cookie))
    Call Terminate
End Sub

Private Sub tmrTimeOut_Timer()
    If DateDiff("n", FTP_LastActive, Now) > 1 Then
        Call Parent.objMSN_SB.CancelInvitation(Cookie, "FTIMEOUT")
        Call TransferError
    End If
End Sub

Private Sub TransferError()
    If Not objMSN_FTP.File = vbNullString Then
        If objMSN_FTP.TransferType = FtpTransferType_Receive Then
            Parent.Comment "You have failed to receive file """ & objMSN_FTP.File & """ from " & Parent.BuddyNick, 128
        Else
            Parent.Comment "Transfer of """ & objMSN_FTP.File & """ has failed.", 128
        End If
        Call QueScript(Me, "TransferFailed", ConvArray(Cookie))
    End If
    Call Terminate
End Sub

Public Sub Terminate()
    Unload Me
End Sub

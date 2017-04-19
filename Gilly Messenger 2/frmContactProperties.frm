VERSION 5.00
Begin VB.Form frmContactProperties 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Contact Properties"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3945
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   3945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   350
      Left            =   1800
      TabIndex        =   9
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   350
      Left            =   2880
      TabIndex        =   8
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox txtComment 
      Appearance      =   0  'Flat
      Height          =   1455
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1200
      Width           =   2895
   End
   Begin VB.TextBox txtStatus 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox txtEmail 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   2775
   End
   Begin VB.TextBox txtNick 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label lblComment 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comment"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   660
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   450
   End
   Begin VB.Label lblEmail 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblNick 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nick"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   330
   End
End
Attribute VB_Name = "frmContactProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    SetBuddyComment txtEmail.Text, txtComment.Text
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

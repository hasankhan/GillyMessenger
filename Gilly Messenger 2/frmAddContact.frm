VERSION 5.00
Begin VB.Form frmAddContact 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gilly Messenger"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   Icon            =   "frmAddContact.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   185
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   505
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   6240
      TabIndex        =   7
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   4920
      TabIndex        =   6
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CheckBox chkAdd 
      Caption         =   "Add this person to my contact list."
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   1950
      Value           =   1  'Checked
      Width           =   3015
   End
   Begin VB.OptionButton optBlock 
      Caption         =   "&Block this person from seeing when you are online and contact you"
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   1200
      Width           =   5055
   End
   Begin VB.OptionButton optAllow 
      Caption         =   "&Allow this person to see when you are online and contact you"
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   840
      Value           =   -1  'True
      Width           =   4695
   End
   Begin VB.Label lblRemember 
      Caption         =   "Remember, you can make yourself appear offline temporarily to everyone at anytime."
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   1680
      Width           =   6015
   End
   Begin VB.Label lblWantTo 
      Caption         =   "Do you want to:"
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   510
      Width           =   1335
   End
   Begin VB.Label lblEmail 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "[Email] has added you to his/her contact list."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "frmAddContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    On Error Resume Next
    If optAllow.Value = True Then
        If GetBuddyProperty(Me.Tag, "block") = "True" Then
            MsnSend "REM " & TrialID & " BL " & Me.Tag, TrialID, frmMain.wskMSN
        End If
        If GetBuddyProperty(Me.Tag, "allow") <> "True" Then
            MsnSend "ADD " & TrialID & " AL " & Me.Tag & " " & Me.Tag, TrialID, frmMain.wskMSN
        End If
    Else
        If GetBuddyProperty(Me.Tag, "allow") = "True" Then
            MsnSend "REM " & TrialID & " AL " & Me.Tag, TrialID, frmMain.wskMSN
        End If
        If GetBuddyProperty(Me.Tag, "block") <> "True" Then
            MsnSend "ADD " & TrialID & " BL " & Me.Tag & " " & Me.Tag, TrialID, frmMain.wskMSN
        End If
    End If
    If chkAdd.Value = vbChecked Then
        If GetBuddyProperty(Me.Tag, "forward") <> "True" Then
            MsnSend "ADD " & TrialID & " FL " & Me.Tag & " " & Me.Tag, TrialID, frmMain.wskMSN
        Else
            frmMain.tvwBuddies.Nodes(Me.Tag).BackColor = -2147483643
        End If
    End If
    Me.Hide
End Sub

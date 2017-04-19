VERSION 5.00
Begin VB.Form frmAddContact 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gilly Messenger"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   Icon            =   "frmAddContact.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   185
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   505
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picEmailHasAddedYouToHisHerContactList 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   489
      TabIndex        =   7
      Top             =   120
      Width           =   7335
      Begin VB.VScrollBar vsEmailHasAddedYouToHisHerContactList 
         Height          =   225
         Left            =   7080
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblEmailHasAddedYouToHisHerContactList 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "[Email] has added you to his/her contact list."
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   7335
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   6240
      TabIndex        =   6
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   360
      Left            =   4920
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CheckBox chkAddThisPersonToMyContactList 
      Caption         =   "Add this person to my contact list."
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   1950
      Value           =   1  'Checked
      Width           =   3015
   End
   Begin VB.OptionButton optBlockThisPersonFromSeeingWhenYouAreOnlineAndContactYou 
      Caption         =   "&Block this person from seeing when you are online and contact you"
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   1200
      Width           =   5055
   End
   Begin VB.OptionButton optAllowThisPersonToSeeWhenYouAreOnlineAndContactYou 
      Caption         =   "&Allow this person to see when you are online and contact you"
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   840
      Value           =   -1  'True
      Width           =   4695
   End
   Begin VB.Label lblRememberYouCanMakeYourselfAppearOfflineTemporarilyToEveryoneAtAnytime 
      Caption         =   "Remember, you can make yourself appear offline temporarily to everyone at anytime."
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   1680
      Width           =   6015
   End
   Begin VB.Label lblDoWantTo 
      Caption         =   "Do you want to:"
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   510
      Width           =   1335
   End
End
Attribute VB_Name = "frmAddContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ContactEmail As String
Public ContactNick As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error Resume Next
    
    If optAllowThisPersonToSeeWhenYouAreOnlineAndContactYou.Value = True Then
        If InList(GetContactAttr(ContactEmail, "lists"), msnList_Block) Then
            frmMain.objMSN_NS.RemoveContact msnList_Block, ContactEmail
        End If
        If Not InList(GetContactAttr(ContactEmail, "lists"), msnList_Allow) Then
            frmMain.objMSN_NS.AddContact msnList_Allow, ContactEmail, ContactNick
        End If
    Else
        If InList(GetContactAttr(ContactEmail, "lists"), msnList_Allow) Then
            frmMain.objMSN_NS.RemoveContact msnList_Allow, ContactEmail
        End If
        If Not InList(GetContactAttr(ContactEmail, "lists"), msnList_Block) Then
            frmMain.objMSN_NS.AddContact msnList_Block, ContactEmail, ContactNick
        End If
    End If
    If chkAddThisPersonToMyContactList.Value = vbChecked Then
        If Not InList(GetContactAttr(ContactEmail, "lists"), msnList_Forward) Then
            frmMain.objMSN_NS.AddContact msnList_Forward, ContactEmail, ContactNick, 0
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
    If Not Transparency = 0 Then
        SetTransparency Me, Transparency
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    LastActive = Timer
End Sub

Private Sub lblEmailHasAddedYouToHisHerContactList_Change()
    Call ResizeScrollLabelSet(picEmailHasAddedYouToHisHerContactList, lblEmailHasAddedYouToHisHerContactList, vsEmailHasAddedYouToHisHerContactList)
    Call vsEmailHasAddedYouToHisHerContactList_Change
End Sub

Private Sub vsEmailHasAddedYouToHisHerContactList_Change()
    lblEmailHasAddedYouToHisHerContactList.Top = -(vsEmailHasAddedYouToHisHerContactList.Value * picEmailHasAddedYouToHisHerContactList.TextHeight(lblEmailHasAddedYouToHisHerContactList.Caption))
End Sub

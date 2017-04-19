VERSION 5.00
Begin VB.Form frmReverseList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Who has you on their contact list?"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3390
   Icon            =   "frmReverseList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   330
      Left            =   2160
      TabIndex        =   0
      Top             =   4440
      Width           =   1095
   End
   Begin VB.ListBox lstReverse 
      Height          =   3570
      ItemData        =   "frmReverseList.frx":000C
      Left            =   120
      List            =   "frmReverseList.frx":000E
      TabIndex        =   2
      Top             =   720
      Width           =   3135
   End
   Begin VB.Label lblMessage 
      Caption         =   "The following people have added you to their contact list:"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
   Begin VB.Menu mnuContact 
      Caption         =   "[Contact]"
      Visible         =   0   'False
      Begin VB.Menu mnuContact_Hide 
         Caption         =   "Hi&de"
      End
      Begin VB.Menu mnuContact_Ignore 
         Caption         =   "&Ignore"
      End
      Begin VB.Menu mnuContact_AddToContacts 
         Caption         =   "&Add to Contacts"
      End
      Begin VB.Menu mnuContact_Seperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContact_Properties 
         Caption         =   "P&roperties"
      End
   End
End
Attribute VB_Name = "frmReverseList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ReverseList As Collection

Private Sub cmdClose_Click()
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

Private Sub lstReverse_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    
    If Button = vbRightButton And Not lstReverse.ListIndex = -1 Then
        mnuContact.Tag = ReverseList(lstReverse.ListIndex + 1)
        If Not InList(GetContactAttr(ReverseList(lstReverse.ListIndex + 1), "lists"), msnList_Forward) Then
            mnuContact_AddToContacts.Enabled = True
            mnuContact_Hide.Enabled = False
        Else
            mnuContact_AddToContacts.Enabled = False
            mnuContact_Hide.Enabled = True
            If Not InCollection(HiddenContacts, ReverseList(lstReverse.ListIndex + 1)) Then
                mnuContact_Hide.Caption = "Hi&de"
            Else
                mnuContact_Hide.Caption = "Unhi&de"
            End If
        End If
        If Not InCollection(IgnoreList, ReverseList(lstReverse.ListIndex + 1)) Then
            mnuContact_Ignore.Caption = "&Ignore"
        Else
            mnuContact_Ignore.Caption = "Un&ignore"
        End If
        PopupMenu mnuContact
    End If
End Sub

Private Sub mnuContact_Hide_Click()
    Select Case mnuContact_Hide.Caption
    Case "Hi&de"
        Call HideContact(mnuContact.Tag)
    Case "Unhi&de"
        Call UnhideContact(mnuContact.Tag)
    End Select
End Sub

Private Sub mnuContact_Ignore_Click()
    Select Case mnuContact_Ignore.Caption
    Case "&Ignore"
        Call IgnoreContact(mnuContact.Tag)
    Case "&Unignore"
        Call UnignoreContact(mnuContact.Tag)
    End Select
End Sub

Private Sub mnuContact_AddToContacts_Click()
    Call AddContact(mnuContact.Tag)
End Sub

Private Sub mnuContact_Properties_Click()
    On Error Resume Next
    
    ShowBuddyProperties Me, ReverseList(lstReverse.ListIndex + 1)
End Sub

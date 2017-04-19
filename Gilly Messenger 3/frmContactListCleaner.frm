VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmContactListCleaner 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contact List cleaner"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   Icon            =   "frmContactListCleaner.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   297
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   353
   StartUpPosition =   1  'CenterOwner
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdClean 
      Caption         =   "&OK"
      Default         =   -1  'True
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
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3960
      Width           =   975
   End
   Begin MSComctlLib.ListView lstContacts 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5530
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Contact"
         Object.Width           =   6615
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "List"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.Label lblMessage 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmContactListCleaner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClean_Click()
    On Error Resume Next
    
    cmdClean.Enabled = False
    If lstContacts.ListItems.Count > 0 Then
        Dim i As Integer
        For i = 1 To lstContacts.ListItems.Count
            frmMain.objMSN_NS.RemoveContact Val(Split(lstContacts.ListItems(i).Key)(0)), CStr(Split(lstContacts.ListItems(i).Key)(1))
        Next
        MsgBox "Contact list cleaned!", vbInformation
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
    
    lblMessage.Caption = "Gilly Messenger will delete following contacts from respective lists." & _
    "To exclude a contact from cleaning list select it and press the delete key."
    Dim i As Integer, Lists As Integer, Email As String
    For i = 1 To ContactList.Count
        Lists = ContactList(i).Item("lists")
        Email = ContactList(i).Item("email")
        If Not InList(Lists, msnList_Reverse) Then
            If InList(Lists, msnList_Forward) Then
                lstContacts.ListItems.Add , msnList_Forward & " " & Email, Email
                lstContacts.ListItems(lstContacts.ListItems.Count).ListSubItems.Add , , "Forward"
                lstContacts.ListItems(lstContacts.ListItems.Count).ToolTipText = Email & " doesn't have you in his/her contact list where as you have."
            End If
            If InList(Lists, msnList_Allow) Then
                lstContacts.ListItems.Add , msnList_Allow & " " & Email, Email
                lstContacts.ListItems(lstContacts.ListItems.Count).ListSubItems.Add , , "Allow"
                lstContacts.ListItems(lstContacts.ListItems.Count).ToolTipText = Email & " doesn't have you in his/her contact list and you have added him/her to allow list."
            End If
            If InList(Lists, msnList_Block) Then
                lstContacts.ListItems.Add , msnList_Block & " " & Email, Email
                lstContacts.ListItems(lstContacts.ListItems.Count).ListSubItems.Add , , "Block"
                lstContacts.ListItems(lstContacts.ListItems.Count).ToolTipText = Email & " doesn't have you in his/her contact list and you have blocked him/her."
            End If
        End If
    Next
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    LastActive = Timer
End Sub

Private Sub lstContacts_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    If KeyCode = vbKeyDelete Then
        lstContacts.ListItems.Remove lstContacts.SelectedItem.Key
    End If
End Sub

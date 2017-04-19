VERSION 5.00
Begin VB.Form frmSelectFolder 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3510
   Icon            =   "frmSelectFolder.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   234
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   330
      Left            =   1560
      TabIndex        =   4
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   2520
      TabIndex        =   3
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdCreateFolder 
      Caption         =   "Create Folder"
      Height          =   330
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   1095
   End
   Begin VB.DirListBox Dir1 
      Height          =   2565
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmSelectFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public srcTextBox As TextBox
Private PrevDrive As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCreateFolder_Click()
    On Error Resume Next
    
    Dim strDirName As String
    strDirName = InputBox("Enter the name of the folder", "Create Folder", frmMain.objMSN_NS.Login)
    If Not strDirName = vbNullString Then
        MkDir Dir1.Path & "\" & strDirName
        Dir1.Refresh
    End If
End Sub

Private Sub cmdOK_Click()
    srcTextBox.Text = Dir1.list(Dir1.ListIndex)
    Unload Me
End Sub

Private Sub Drive1_Change()
    On Error GoTo Handler:
    Dir1.Path = Drive1.Drive
    PrevDrive = Drive1.Drive
Handler:
    Drive1.Drive = PrevDrive
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
    PrevDrive = Drive1.Drive
    If Not Transparency = 0 Then
        SetTransparency Me, Transparency
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    LastActive = Timer
End Sub

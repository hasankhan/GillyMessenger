VERSION 5.00
Begin VB.Form frmPopupFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Popup Filter"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4365
   Icon            =   "frmPopupFilter.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optShowPopupFor 
      Caption         =   "Show popup for every contact in the list"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
   Begin VB.OptionButton optShowPopupExcept 
      Caption         =   "Show popup for every contact except these in the list"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4095
   End
   Begin VB.ListBox lstContacts 
      Height          =   1230
      ItemData        =   "frmPopupFilter.frx":000C
      Left            =   120
      List            =   "frmPopupFilter.frx":000E
      TabIndex        =   2
      Top             =   840
      Width           =   4095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   330
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   330
      Left            =   1920
      TabIndex        =   5
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   3120
      TabIndex        =   4
      Top             =   2160
      Width           =   1095
   End
End
Attribute VB_Name = "frmPopupFilter"
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

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error Resume Next
    
    If frmMain.objMSN_NS.State = NsState_SignedIn Then
        PopupFilterMode = Not optShowPopupFor.Value
        
        DeleteSetting "Gilly Messenger", "Popup Filter\" & frmMain.objMSN_NS.Login
        SaveSettingX "Popup Filter\" & frmMain.objMSN_NS.Login, "Mode", IIf(optShowPopupFor.Value, 0, 1)
        
        Set PopupFilter = Nothing
        Set PopupFilter = New Collection
        
        If Not lstContacts.ListCount = 0 Then
            Dim i As Integer
            For i = 0 To lstContacts.ListCount - 1
                SaveSettingX "Popup Filter\" & frmMain.objMSN_NS.Login, CStr(i), lstContacts.list(i)
                PopupFilter.Add lstContacts.list(i), lstContacts.list(i)
            Next
        Else
            PopupFilter.Add "*@*.*", "*@*.*"
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
    If Not PopupFilterMode Then
        optShowPopupFor.Value = True
    Else
        optShowPopupExcept.Value = True
    End If
    
    Dim i As Integer
    For i = 1 To PopupFilter.Count
        lstContacts.AddItem PopupFilter(i)
    Next
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

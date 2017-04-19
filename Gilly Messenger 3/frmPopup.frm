VERSION 5.00
Begin VB.Form frmPopup 
   BorderStyle     =   0  'None
   ClientHeight    =   1755
   ClientLeft      =   24030
   ClientTop       =   12495
   ClientWidth     =   2730
   Icon            =   "frmPopup.frx":0000
   KeyPreview      =   -1  'True
   ScaleHeight     =   117
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   182
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer tmrHide 
      Interval        =   10000
      Left            =   1080
      Top             =   840
   End
   Begin VB.Image imgClose 
      Height          =   180
      Left            =   2460
      Top             =   120
      Width           =   180
   End
   Begin VB.Image imgOptions 
      Height          =   255
      Left            =   2070
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   795
      Left            =   240
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   720
      UseMnemonic     =   0   'False
      Width           =   2295
   End
End
Attribute VB_Name = "frmPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Source As Form

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    LastActive = Timer
End Sub

Private Sub Form_Load()
    If Not Transparency = 0 Then
        SetTransparency Me, Transparency
    End If
    Me.Picture = LoadResPicture("POPUPBACKGROUND", vbResBitmap)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    LastActive = Timer
End Sub

Private Sub imgClose_Click()
    Call TerminatePopup
End Sub

Private Sub imgOptions_Click()
    On Error Resume Next
    
    frmOptions.Show vbModal, frmMain
    Call TerminatePopup
End Sub

Private Sub lblMessage_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        Dim ActionParams() As String
        ActionParams = Split(Me.Tag)
        Select Case ActionParams(0)
        Case "CONVO"
            ActivateWindow Source
        Case "CHAT"
            StartChat ActionParams(1), , , True
        Case "URL"
            OpenMsnURL ActionParams(1), ActionParams(2), Val(ActionParams(3))
        End Select
    End If
    Call TerminatePopup
End Sub

Private Sub tmrHide_Timer()
    tmrHide.Enabled = False
    Dim i As Integer, j As Integer
    
    j = Me.Top
    For i = 1 To 1755 Step 25
        Me.Height = 1755 - i
        Me.Top = j + i
        Sleep 1
        DoEvents
    Next
    
    Call TerminatePopup
End Sub

Private Sub TerminatePopup()
    Dim DesktopRect As RECT
    
    SystemParametersInfo SPI_GETWORKAREA, 0, DesktopRect, 0
    
    If LastPopup = Me.hwnd Then
        NewPopupTop = DesktopRect.Bottom
    End If
    
    Unload Me
End Sub

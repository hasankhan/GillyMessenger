VERSION 5.00
Begin VB.Form frmDisplayPic 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Display Picture"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2070
   Icon            =   "frmDisplayPic.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   199
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   138
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   1080
      TabIndex        =   3
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Enabled         =   0   'False
      Height          =   330
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   855
   End
   Begin VB.PictureBox picDisplay 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   1800
      Left            =   120
      ScaleHeight     =   116
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   116
      TabIndex        =   1
      Top             =   120
      Width           =   1800
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "&Change"
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   1815
   End
End
Attribute VB_Name = "frmDisplayPic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function Bmp2Jpeg Lib "Bmp2Jpeg.dll" Alias "BmpToJpeg" (ByVal BmpFilename As String, ByVal JpegFilename As String, ByVal CompressQuality As Integer) As Integer
Dim DpPath As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdChange_Click()
    On Error Resume Next
    Kill DpPath
    On Error GoTo Handler
    
    If Not GetUserFile("Images (*.jpg;*.bmp)|*.jpg;*.bmp", "Display Picture") = vbNullString Then
        Dim DisplayPic As IPictureDisp
        Set DisplayPic = LoadPicture(frmMain.CommonDialog.FileName)
        Dim intWidth As Double, intHeight As Double, intLeft As Integer, intTop As Integer
        intWidth = DisplayPic.Width
        intHeight = DisplayPic.Height
        If intWidth > intHeight Then
            intLeft = (intWidth - intHeight) / 2
            intWidth = intHeight
        ElseIf intHeight > intWidth Then
            intTop = (intHeight - intWidth) / 2
            intHeight = intWidth
        End If
        DisplayPic.Render picDisplay.hDC, 0, 0, 120, 120, intLeft, intTop, intWidth, intHeight, vbNull
        MakeSureDirectoryPathExists GetTempDir
        DpPath = Replace$(GetTempDir & "\gmdp" & Fix(Timer) & ".dat", "\\", "\")
        SavePicture picDisplay.Image, DpPath
        Set DisplayPic = LoadPicture(DpPath)
        Set picDisplay.Picture = LoadPicture(vbNullString)
        DisplayPic.Render picDisplay.hDC, 0, 0, 120, 120, 0, 0, DisplayPic.Width, DisplayPic.Height, vbNull
        SavePicture picDisplay.Image, DpPath
        Dim DllPath As String
        DllPath = Replace$(GetSysDir & "\bmp2jpeg.dll", "\\", "\")
        If Not FileExists(DllPath) Then
            LoadDataIntoFile "BMP2JPEG", DllPath
        End If
        Bmp2Jpeg DpPath, DpPath, 50
        cmdOK.Enabled = True
    End If
    Exit Sub
Handler:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub cmdOK_Click()
    MakeSureDirectoryPathExists App.Path & "\Display Pics\"
    FileCopy DpPath, App.Path & "\Display Pics\" & frmMain.objMSN_NS.Login & ".dat"
    Kill DpPath
    SaveSettingX "Display Pics", frmMain.objMSN_NS.Login, Format(Now, "YYMMDDhhnnss")
    Dim frmIM As Form
    For Each frmIM In IMWindows
        Call frmIM.RefreshMyDP
        If frmIM.ChatBuddies.Count = 1 Then
            Call frmIM.OfferDP
        End If
    Next
    For Each frmIM In PendingIM
        Call frmIM.RefreshMyDP
    Next
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
    
    If frmMain.objMSN_NS.State = NsState_SignedIn Then
        Call LoadDP(frmMain.objMSN_NS.Login, picDisplay)
    Else
        cmdChange.Enabled = False
        cmdOK.Enabled = False
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    LastActive = Timer
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Kill DpPath
End Sub

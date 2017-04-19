VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Gilly Messenger"
   ClientHeight    =   5295
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   3750
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   ScaleHeight     =   353
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   250
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList imglstPictures 
      Left            =   2040
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   57
      ImageHeight     =   58
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CCA
            Key             =   "IMW_Resize"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":126E
            Key             =   "IMW_TopBarRight"
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrGMScript 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1560
      Top             =   3120
   End
   Begin VB.Timer tmrMsnFileKiller 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1560
      Top             =   2640
   End
   Begin VB.Timer tmrAnimator 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   2520
      Top             =   2640
   End
   Begin VB.Timer tmrPing 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   2040
      Top             =   2640
   End
   Begin MSWinsockLib.Winsock wskMSN 
      Left            =   600
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imglstStatus 
      Left            =   2040
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D08
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":205C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":23B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2704
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A58
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2D7C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imglstEmoticons 
      Left            =   2640
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   19
      ImageHeight     =   19
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   78
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":30A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3670
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3C40
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4210
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":47E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4CA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5278
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5848
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5E18
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":631C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":68EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6EBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":748C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7A5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":802C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":85FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8BCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":919C
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":976C
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9D3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A30C
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A7D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AC9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B26C
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B734
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BBFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C1CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C79C
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CD6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D33C
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D90C
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DEDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E4AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EA7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F04C
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F514
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F944
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FF14
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":104E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10AB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11084
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11654
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11C24
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":121F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":127C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12D94
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13364
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13934
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13F04
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":144D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14AA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15074
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15644
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15C14
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":161E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":167B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16D84
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17354
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17924
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17EF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":184C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18A94
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19064
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19634
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19C04
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A1D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A7A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1AD74
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B23C
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B80C
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BBA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BF00
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C494
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C95C
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CE24
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D2EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D7B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1DD84
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   0
      ScaleHeight     =   329
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   263
      TabIndex        =   4
      Top             =   0
      Width           =   3945
      Begin MSWinsockLib.Winsock wskSSL 
         Left            =   1080
         Top             =   2640
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.PictureBox picSignIn 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   1245
         MouseIcon       =   "frmMain.frx":1E24C
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":1EB16
         ScaleHeight     =   435
         ScaleWidth      =   1260
         TabIndex        =   9
         Top             =   1200
         Width           =   1260
      End
      Begin VB.PictureBox picTrayIcon 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   240
         Picture         =   "frmMain.frx":207E4
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   6
         Top             =   960
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox picEmpty 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         Picture         =   "frmMain.frx":20B6E
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   5
         Top             =   960
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblSigningIn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Signing In..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1365
         TabIndex        =   12
         Top             =   1320
         Width           =   990
      End
      Begin VB.Label lblGilly 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gilly"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008D2F11&
         Height          =   240
         Left            =   1500
         TabIndex        =   11
         Top             =   120
         Width           =   435
      End
      Begin VB.Label lblWelcome 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "welcome"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E9CAB1&
         Height          =   405
         Left            =   120
         TabIndex        =   10
         Top             =   0
         Width           =   1290
      End
   End
   Begin VB.PictureBox picStatus 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F5E2D6&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   249
      TabIndex        =   7
      Top             =   4920
      Width           =   3735
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00814D3C&
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   90
         UseMnemonic     =   0   'False
         Width           =   3540
      End
   End
   Begin MSComDlg.CommonDialog cdOpen 
      Left            =   600
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtNick 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   450
      MaxLength       =   129
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   180
      Visible         =   0   'False
      Width           =   3165
   End
   Begin MSComctlLib.TreeView tvwBuddies 
      Height          =   3810
      Left            =   60
      TabIndex        =   0
      Top             =   900
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   6720
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      Style           =   1
      ImageList       =   "imglstStatus"
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image imgStatus 
      Height          =   240
      Left            =   120
      Top             =   120
      Width           =   240
   End
   Begin VB.Image imgEmail 
      Height          =   240
      Left            =   60
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":20C40
      Top             =   570
      Width           =   270
   End
   Begin VB.Label lblEmail 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "No new e-mail messages"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   390
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.Label lblNick 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   450
      TabIndex        =   1
      Top             =   180
      UseMnemonic     =   0   'False
      Width           =   3165
   End
   Begin VB.Image imgTopRight 
      Height          =   450
      Left            =   2280
      Picture         =   "frmMain.frx":21002
      Top             =   0
      Width           =   1560
   End
   Begin VB.Menu mnuActions 
      Caption         =   "&Actions"
      Begin VB.Menu mnuSignIn 
         Caption         =   "Sign &In"
      End
      Begin VB.Menu mnuStatus 
         Caption         =   "&My Status"
         Enabled         =   0   'False
         Begin VB.Menu mnuStatusList 
            Caption         =   "&Online"
            Index           =   0
         End
         Begin VB.Menu mnuStatusList 
            Caption         =   "&Busy"
            Index           =   1
         End
         Begin VB.Menu mnuStatusList 
            Caption         =   "B&e Right Back"
            Index           =   2
         End
         Begin VB.Menu mnuStatusList 
            Caption         =   "&Away"
            Index           =   3
         End
         Begin VB.Menu mnuStatusList 
            Caption         =   "On The &Phone"
            Index           =   4
         End
         Begin VB.Menu mnuStatusList 
            Caption         =   "Out To &Lunch"
            Index           =   5
         End
         Begin VB.Menu mnuStatusList 
            Caption         =   "&Idle"
            Index           =   6
         End
         Begin VB.Menu mnuStatusList 
            Caption         =   "Appear O&ffline"
            Index           =   7
         End
      End
      Begin VB.Menu mnuOpenInbox 
         Caption         =   "Open &Inbox"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSendMessage 
         Caption         =   "&Send Message"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuChatRooms 
         Caption         =   "Goto &Chatrooms"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEditPassport 
         Caption         =   "Edit Passpor&t"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEditProfile 
         Caption         =   "Edit Profi&le"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuContacts 
      Caption         =   "&Contacts"
      Begin VB.Menu mnuAddContact 
         Caption         =   "&Add Contact"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "&Search Contact"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuViewContactsBy 
         Caption         =   "View &Contacts By"
         Enabled         =   0   'False
         Begin VB.Menu mnuViewContactsByDisplayName 
            Caption         =   "&Display name"
         End
         Begin VB.Menu mnuViewContactsByEmail 
            Caption         =   "&E-mail address"
         End
      End
      Begin VB.Menu mnuSaveContactList 
         Caption         =   "Save Contact &List"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuImportContactList 
         Caption         =   "Im&port Contact List"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuAutoMessage 
         Caption         =   "&Auto Message"
      End
      Begin VB.Menu mnuMessageAll 
         Caption         =   "&Message All"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuIgnoreAll 
         Caption         =   "&Ignore All"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuChatLogger 
         Caption         =   "Chat &Logger"
      End
      Begin VB.Menu mnuStatusLogger 
         Caption         =   "S&tatus Logger"
      End
      Begin VB.Menu mnuGMScript 
         Caption         =   "Run &GM Script"
         Enabled         =   0   'False
         Begin VB.Menu mnuScript 
            Caption         =   "No Scripts Available"
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu mnuScriptSep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuScriptOther 
            Caption         =   "&Other..."
         End
      End
      Begin VB.Menu mnuStopScript 
         Caption         =   "Stop &GM Script"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRemoteControl 
         Caption         =   "&Remote Control"
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "Sett&ings"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuReadme 
         Caption         =   "&Read Me"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuCrackSoft 
         Caption         =   "&CrackSoft"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
   End
   Begin VB.Menu mnuContact 
      Caption         =   "Contact"
      Visible         =   0   'False
      Begin VB.Menu mnuMsg 
         Caption         =   "&Send an Instant Message"
      End
      Begin VB.Menu mnuSendEmail 
         Caption         =   "Send &E-mail (Email)"
      End
      Begin VB.Menu mnuCopyNick 
         Caption         =   "Copy &Nick"
      End
      Begin VB.Menu mnuCopyEmail 
         Caption         =   "Copy &Email"
      End
      Begin VB.Menu mnuBlock 
         Caption         =   "(State)"
      End
      Begin VB.Menu mnuIgnore 
         Caption         =   "(Ignore)"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuProfile 
         Caption         =   "View &Profile"
      End
      Begin VB.Menu mnuViewLog 
         Caption         =   "&View &Log"
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "&Properties"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Dim TrayAnim As Integer
Dim MsnData As String, cSearch As String
Dim SSL_Error As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
    If cSearch <> vbNullString Then
        For X = tvwBuddies.SelectedItem.Index + 1 To tvwBuddies.Nodes.Count
            If tvwBuddies.Nodes(X).Key <> "NoneOnline" And tvwBuddies.Nodes(X).Key <> "NoneOffline" Then
                If (tvwBuddies.Nodes(X).Key Like cSearch) Or (tvwBuddies.Nodes(X).Text Like cSearch) Or (GetBuddyComment(tvwBuddies.Nodes(X).Key) Like cSearch) Then
                    tvwBuddies.Nodes(X).EnsureVisible
                    tvwBuddies.Nodes(X).Selected = True
                    Exit Sub
                End If
            End If
        Next X
        MsgBox "Contact not found!", vbInformation, "Search Contact"
    End If
ElseIf KeyCode = vbKeyN And Shift = ShiftConstants.vbAltMask And SignedIn = True Then
    Call lblNick_Click
ElseIf KeyCode = vbKeyE And Shift = ShiftConstants.vbAltMask And SignedIn = True Then
    Call lblEmail_Click
ElseIf KeyCode = vbKeyS And Shift = ShiftConstants.vbAltMask And SignedIn = False And picSignIn.Visible = True Then
    Call picSignIn_Click
End If
End Sub

Public Sub Form_Load()
On Error Resume Next
imglstPictures.ListImages.Add , "POPUP", LoadResPicture("POPUP", vbResBitmap)
imglstPictures.ListImages.Add , "IMW_Background", LoadResPicture("IMW_BACKGROUND", vbResBitmap)
imglstPictures.ListImages.Add , "IMW_TopBarLeftCorner", LoadResPicture("IMW_TOPBARLEFTCORNER", vbResBitmap)
imglstPictures.ListImages.Add , "IMW_TopBarMid", LoadResPicture("IMW_TOPBARMID", vbResBitmap)
InitStatus = True
StatusImage = 2
imgEmail.MouseIcon = picSignIn.MouseIcon
lblEmail.MouseIcon = picSignIn.MouseIcon
'Declare errors
MsnError.Add "Syntax MsnError", "200"
MsnError.Add "Invalid parameter.", "201"
MsnError.Add "User not found!", "205"
MsnError.Add "FQDN missing.", "206"
MsnError.Add "Already login.", "207"
MsnError.Add "Invalid username.", "208"
MsnError.Add "Invalid nick name.", "209"
MsnError.Add "Contact list is full.", "210"
MsnError.Add "Already in the mode.", "218"
MsnError.Add "Switch board failed.", "280"
MsnError.Add "Notify XFR failed.", "281"
MsnError.Add "Required fields are missing.", "300"
MsnError.Add "You are not signed in.", "302"
MsnError.Add "Error in internal server.", "500"
MsnError.Add "Error in DB server.", "501"
MsnError.Add "Error in file operation.", "510"
MsnError.Add "Error in memory allocation.", "520"
MsnError.Add "Server is busy.", "600"
MsnError.Add "Server is unavailable.", "601"
MsnError.Add "Peer NS down.", "602"
MsnError.Add "Error connecting to command base.", "603"
MsnError.Add "Server is going down.", "604"
MsnError.Add "Error creating connection.", "707"
MsnError.Add "Error writing block.", "711"
MsnError.Add "Session is overload.", "712"
MsnError.Add "User too active.", "713"
MsnError.Add "Too many sessions.", "714"
MsnError.Add "Unexpected error.", "715"
MsnError.Add "Bad friend file.", "717"
MsnError.Add "Status could not be changed.", "800"
MsnError.Add "Invalid username or password.", "911"
MsnError.Add "Not allowed when offline.", "913"
MsnError.Add "Not accepting new users.", "920"

'Add Emoticons
Emoticons(0, 0) = ":-)": Emoticons(0, 1) = "1"
Emoticons(1, 0) = ":)": Emoticons(1, 1) = "1"
Emoticons(2, 0) = ":-D": Emoticons(2, 1) = "2"
Emoticons(3, 0) = ":D": Emoticons(3, 1) = "2"
Emoticons(4, 0) = ":-O": Emoticons(4, 1) = "3"
Emoticons(5, 0) = ":O": Emoticons(5, 1) = "3"
Emoticons(6, 0) = ":-P": Emoticons(6, 1) = "4"
Emoticons(7, 0) = ":P": Emoticons(7, 1) = "4"
Emoticons(8, 0) = ";-)": Emoticons(8, 1) = "5"
Emoticons(9, 0) = ";)": Emoticons(9, 1) = "5"
Emoticons(10, 0) = ":-(": Emoticons(10, 1) = "6"
Emoticons(11, 0) = ":(": Emoticons(11, 1) = "6"
Emoticons(12, 0) = ":-S": Emoticons(12, 1) = "7"
Emoticons(13, 0) = ":S": Emoticons(13, 1) = "7"
Emoticons(14, 0) = ":-|": Emoticons(14, 1) = "8"
Emoticons(15, 0) = ":|": Emoticons(15, 1) = "8"
Emoticons(16, 0) = ":'(": Emoticons(16, 1) = "9"
Emoticons(17, 0) = ":-$": Emoticons(17, 1) = "10"
Emoticons(18, 0) = ":$": Emoticons(18, 1) = "10"
Emoticons(19, 0) = "(H)": Emoticons(19, 1) = "11"
Emoticons(20, 0) = ":-@": Emoticons(20, 1) = "12"
Emoticons(21, 0) = ":@": Emoticons(21, 1) = "12"
Emoticons(22, 0) = "(A)": Emoticons(22, 1) = "13"
Emoticons(23, 0) = "(6)": Emoticons(23, 1) = "14"
Emoticons(24, 0) = ":-#": Emoticons(24, 1) = "15"
Emoticons(25, 0) = "8O|": Emoticons(25, 1) = "16"
Emoticons(26, 0) = "8-|": Emoticons(26, 1) = "17"
Emoticons(27, 0) = "^O)": Emoticons(27, 1) = "18"
Emoticons(28, 0) = ":-*": Emoticons(28, 1) = "19"
Emoticons(29, 0) = "+O(": Emoticons(29, 1) = "20"
Emoticons(30, 0) = ":^)": Emoticons(30, 1) = "21"
Emoticons(31, 0) = "*-)": Emoticons(31, 1) = "22"
Emoticons(32, 0) = "<:O)": Emoticons(32, 1) = "23"
Emoticons(33, 0) = "8-)": Emoticons(33, 1) = "24"
Emoticons(34, 0) = "|-)": Emoticons(34, 1) = "25"
Emoticons(35, 0) = "(C)": Emoticons(35, 1) = "26"
Emoticons(36, 0) = "(Y)": Emoticons(36, 1) = "27"
Emoticons(37, 0) = "(N)": Emoticons(37, 1) = "28"
Emoticons(38, 0) = "(B)": Emoticons(38, 1) = "29"
Emoticons(39, 0) = "(D)": Emoticons(39, 1) = "30"
Emoticons(40, 0) = "(X)": Emoticons(40, 1) = "31"
Emoticons(41, 0) = "(Z)": Emoticons(41, 1) = "32"
Emoticons(42, 0) = "({)": Emoticons(42, 1) = "33"
Emoticons(43, 0) = "(})": Emoticons(43, 1) = "34"
Emoticons(44, 0) = ":-[": Emoticons(44, 1) = "35"
Emoticons(45, 0) = ":[": Emoticons(45, 1) = "35"
Emoticons(46, 0) = "(^)": Emoticons(46, 1) = "36"
Emoticons(47, 0) = "(L)": Emoticons(47, 1) = "37"
Emoticons(48, 0) = "(U)": Emoticons(48, 1) = "38"
Emoticons(49, 0) = "(K)": Emoticons(49, 1) = "39"
Emoticons(50, 0) = "(G)": Emoticons(50, 1) = "40"
Emoticons(51, 0) = "(F)": Emoticons(51, 1) = "41"
Emoticons(52, 0) = "(W)": Emoticons(52, 1) = "42"
Emoticons(53, 0) = "(P)": Emoticons(53, 1) = "43"
Emoticons(54, 0) = "(~)": Emoticons(54, 1) = "44"
Emoticons(55, 0) = "(@)": Emoticons(55, 1) = "45"
Emoticons(56, 0) = "(&)": Emoticons(56, 1) = "46"
Emoticons(57, 0) = "(T)": Emoticons(57, 1) = "47"
Emoticons(58, 0) = "(I)": Emoticons(58, 1) = "48"
Emoticons(59, 0) = "(8)": Emoticons(59, 1) = "49"
Emoticons(60, 0) = "(S)": Emoticons(60, 1) = "50"
Emoticons(61, 0) = "(*)": Emoticons(61, 1) = "51"
Emoticons(62, 0) = "(E)": Emoticons(62, 1) = "52"
Emoticons(63, 0) = "(O)": Emoticons(63, 1) = "53"
Emoticons(64, 0) = "(0)": Emoticons(64, 1) = "53"
Emoticons(65, 0) = "(M)": Emoticons(65, 1) = "54"
Emoticons(66, 0) = "(SN)": Emoticons(66, 1) = "55"
Emoticons(67, 0) = "(BAH)": Emoticons(67, 1) = "56"
Emoticons(68, 0) = "(PL)": Emoticons(68, 1) = "57"
Emoticons(69, 0) = "(||)": Emoticons(69, 1) = "58"
Emoticons(70, 0) = "(PI)": Emoticons(70, 1) = "59"
Emoticons(71, 0) = "(SO)": Emoticons(71, 1) = "60"
Emoticons(72, 0) = "(AU)": Emoticons(72, 1) = "61"
Emoticons(73, 0) = "(AP)": Emoticons(73, 1) = "62"
Emoticons(74, 0) = "(UM)": Emoticons(74, 1) = "63"
Emoticons(75, 0) = "(IP)": Emoticons(75, 1) = "64"
Emoticons(76, 0) = "(CO)": Emoticons(76, 1) = "65"
Emoticons(77, 0) = "(MP)": Emoticons(77, 1) = "66"
Emoticons(78, 0) = "(ST)": Emoticons(78, 1) = "67"
Emoticons(79, 0) = "(LI)": Emoticons(79, 1) = "68"
Emoticons(80, 0) = "(MO)": Emoticons(80, 1) = "69"
Emoticons(81, 0) = "(#)": Emoticons(81, 1) = "70"
Emoticons(82, 0) = "(R)": Emoticons(82, 1) = "71"
Emoticons(83, 0) = "(?)": Emoticons(83, 1) = "72"
Emoticons(84, 0) = "(BRB)": Emoticons(84, 1) = "73"
Emoticons(85, 0) = "(H5)": Emoticons(85, 1) = "74"
Emoticons(86, 0) = "(TU)": Emoticons(86, 1) = "75"
Emoticons(87, 0) = "(YN)": Emoticons(87, 1) = "76"
Emoticons(88, 0) = "(CI)": Emoticons(88, 1) = "77"
Emoticons(89, 0) = "(XX)": Emoticons(89, 1) = "78"

For X = 0 To 76
    frmEmoticons.imgEmoticon(X).Picture = imglstEmoticons.ListImages(X + 1).Picture
Next

'Get App Settings
Call LoadAppSettings

'Get server settings
wskMSN.RemoteHost = GetSetting("Gilly Messenger", "Server Settings", "IP Address", "messenger.hotmail.com")
wskMSN.RemotePort = GetSetting("Gilly Messenger", "Server Settings", "Port", "1863")
'Sign In Options
SignInMode = GetSetting("Gilly Messenger", "Sign In", "Mode", "Complete")
InitialStatus = GetSetting("Gilly Messenger", "Sign In", "Status", 7)
Status = InitialStatus
'Get cached login names
Dim LoginCache() As String
LoginCache = GetAllSettings("Gilly Messenger", "Login Cache")
Temp = Str(UBound(LoginCache))
For X = 0 To Val(Temp)
    frmSignIn.cmbLogin.AddItem LoginCache(X, 0)
    DoEvents
Next X
'Get temp path
TempPath = String$(100, Chr$(0))
GetTempPath 100, TempPath
TempPath = Left$(TempPath, InStr(TempPath, Chr$(0)) - 1)
'Load menus
Call LoadFileMenu(frmMain.mnuScript, App.Path & "\Scripts", "*.gms")
'Variable initialization
SystemParametersInfo SPI_GETWORKAREA, 0, R, 0
PopupHeight = R.Bottom * Screen.TwipsPerPixelY
TrialID = 1
Call AddIcon
ChDir "\"
'Common dialogs
cdOpen.InitDir = LastDir
cdOpen.Flags = cdlOFNHideReadOnly Or cdlOFNFileMustExist
If UCase$(Command$) = "/STARTUP" Then
    Me.Visible = False
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = vbAppWindows Or UnloadMode = vbAppTaskManager Then End
Me.Visible = False
If Me.WindowState <> vbMaximized Then
    Me.WindowState = vbNormal
End If
Cancel = 1
End Sub

Public Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    DoEvents
    imgTopRight.Left = Me.ScaleWidth - imgTopRight.Width
    'Nickname
    lblNick.Width = Me.ScaleWidth - imgStatus.Width - 20
    txtNick.Width = lblNick.Width
    'Outer Mask
    picMask.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - picStatus.Height
    picMask.Cls
    GradientFill picMask.hDC, 0, 0, picMask.ScaleWidth, picMask.ScaleHeight / 3, "FFFFFF", "C9D3F3", True
    GradientFill picMask.hDC, 0, picMask.ScaleHeight / 3, picMask.ScaleWidth, picMask.ScaleHeight, "C9D3F3", "FFFFFF", True
    'Contact List
    tvwBuddies.Move 4, 60, Me.ScaleWidth - 8, Me.ScaleHeight - tvwBuddies.Top - picStatus.Height - 8
    'Status Bar
    picStatus.Move 0, Me.ScaleHeight - picStatus.Height, Me.ScaleWidth
    lblStatus.Width = Me.ScaleWidth - 8
    GradientFill picStatus.hDC, 0, 0, picStatus.ScaleWidth / 2, picStatus.ScaleHeight, "CBD8EF", "F0F4FB", False
    GradientFill picStatus.hDC, picStatus.ScaleWidth / 2, 0, picStatus.ScaleWidth, picStatus.ScaleHeight, "F0F4FB", "CBD8EF", False
    picStatus.Line (0, 1)-(picStatus.ScaleWidth, 1), Val("&HE4BAA7")
    'Signin Button
    picSignIn.Move (picMask.ScaleWidth \ 2) - (picSignIn.Width \ 2), (picMask.ScaleHeight \ 3) - (picSignIn.Height \ 2)
    'SigningIn Label
    lblSigningIn.Move picSignIn.Left, picSignIn.Top
    'Main Window
    frmMain.Cls
    GradientFill frmMain.hDC, 0, 0, frmMain.ScaleWidth, 30, "FFFFFF", "E2E9FB", True
    frmMain.Line (0, 30)-(frmMain.ScaleWidth, 30), Val("&HE4BAA7")
    imglstStatus.ListImages(StatusImage).Draw Me.hDC, 8, 8, imlTransparent
End Sub

Private Sub imgEmail_Click()
Call mnuOpenInbox_Click
End Sub

Private Sub lblEmail_Click()
Call mnuOpenInbox_Click
End Sub

Public Sub lblNick_Click()
On Error Resume Next
If SignedIn = True Then
    txtNick.SelStart = 0
    txtNick.SelLength = Len(txtNick.Text)
    txtNick.Visible = True
    txtNick.SetFocus
End If
End Sub

Private Sub mnuChatRooms_Click()
    MsnUrlType.Add "02"
    MsnSend "URL " & TrialID & " CHAT 0x0409", TrialID, wskMSN
End Sub

Private Sub mnuEditPassport_Click()
    MsnUrlType.Add "03"
    MsnSend "URL " & TrialID & " PERSON 0x0409", TrialID, wskMSN
End Sub

Private Sub mnuImportContactList_Click()
    cdOpen.Filter = "Contact List (*.gcl)|*.gcl"
    cdOpen.ShowOpen
    If cdOpen.FileName <> vbNullString Then
        Dim FileNum As Integer
        FileNum = FreeFile
        Open cdOpen.FileName For Input As #FileNum
        Dim strContact As String
        Do Until EOF(FileNum) = True
            Input #FileNum, strContact
            If InStr(strContact, "@") > 0 And InStr(strContact, ".") > 0 Then
                If IsInList(strContact) = False Then
                    AddContact strContact
                End If
            End If
        Loop
        Close #FileNum
        LastDir = Left$(cdOpen.FileName, InStrRev(cdOpen.FileName, "\") - 1)
        cdOpen.FileName = vbNullString
        cdOpen.InitDir = LastDir
        MsgBox "Contacts imported successfully.", vbInformation
    End If
End Sub

Private Sub mnuSaveContactList_Click()
    cdOpen.Filter = "Contact List (*.gcl)|*.gcl"
    cdOpen.ShowSave
    If cdOpen.FileName <> vbNullString Then
        Dim FileNum As Integer
        FileNum = FreeFile
        Open cdOpen.FileName For Output As #FileNum
        For X = 3 To tvwBuddies.Nodes.Count
            If tvwBuddies.Nodes(X).Key <> "NoneOnline" And tvwBuddies.Nodes(X).Key <> "NoneOffline" Then
                Print #FileNum, tvwBuddies.Nodes(X).Key
            End If
        Next
        Close FileNum
        cdOpen.FileName = vbNullString
        MsgBox "Contacts saved successfully", vbInformation
    End If
End Sub

Private Sub mnuSendEmail_Click()
    MsnUrlType.Add "00"
    MsnSend "URL " & TrialID & " COMPOSE " & mnuSendEmail.Tag, TrialID, wskMSN
End Sub

Private Sub mnuViewContactsByDisplayName_Click()
    mnuViewContactsByDisplayName.Checked = True
    ViewContactsByEmail = False
    mnuViewContactsByEmail.Checked = False
    SaveSetting "Gilly Messenger", "App Settings\" & Login, "View Contacts By Email", ViewContactsByEmail
    RefreshList
End Sub

Private Sub mnuViewContactsByEmail_Click()
    mnuViewContactsByEmail.Checked = True
    ViewContactsByEmail = True
    mnuViewContactsByDisplayName.Checked = False
    SaveSetting "Gilly Messenger", "App Settings\" & Login, "View Contacts By Email", ViewContactsByEmail
    RefreshList
End Sub

Public Sub picSignIn_Click()
ShowMe frmMain
frmSignIn.Show 1, Me
End Sub

Private Sub lblStatus_DblClick()
    On Error Resume Next
    ShellExecute 0, "open", StatusLogDir & "\" & Login & ".txt", vbNullString, vbNullString, 1
End Sub

Private Sub lblStatus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then lblStatus.Caption = vbNullString
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show , Me
End Sub

Private Sub mnuAddContact_Click()
Temp = InputBox("Enter the email of the person you want to add.", "Add Contact")
If Trim$(Temp) <> vbNullString Then
    AddContact Temp
End If
End Sub

Public Sub mnuAutoMessage_Click()
If mnuAutoMessage.Checked = False Then
    Temp = InputBox("Enter the message.", "Auto Messeger")
    If Temp <> vbNullString Then
        SetAutoMsg Temp
        mnuAutoMessage.Checked = True
    End If
Else
    mnuAutoMessage.Checked = False
    AutoMsg = vbNullString
    lblStatus.Caption = "Auto message unset."
End If
End Sub

Public Sub mnuBlock_Click()
If mnuBlock.Caption = "&Block" Then
    Call Block(mnuBlock.Tag)
Else
    Call UnBlock(mnuBlock.Tag)
End If
End Sub

Public Sub mnuChatLogger_Click()
mnuChatLogger.Checked = Not mnuChatLogger.Checked
If mnuChatLogger.Checked = True Then
    lblStatus.Caption = "Chat logger activated."
Else
    lblStatus.Caption = "Chat logger deactivated."
End If
End Sub

Private Sub mnuCopyEmail_Click()
Clipboard.Clear
Clipboard.SetText mnuCopyEmail.Tag
End Sub

Private Sub mnuCopyNick_Click()
Clipboard.Clear
Clipboard.SetText mnuCopyNick.Tag
End Sub

Private Sub mnuCrackSoft_Click()
ShellExecute Me.hwnd, vbNullString, "http://www.cracksoft.net.pk/", vbNullString, vbNullString, 1
End Sub

Private Sub mnuDelete_Click()
DeleteContact mnuDelete.Tag
End Sub

Public Sub mnuEditProfile_Click()
MsnUrlType.Add "01"
MsnSend "URL " & TrialID & " PROFILE 0x1409", TrialID, wskMSN
End Sub

Public Sub mnuExit_Click()
On Error Resume Next
Call DeleteIcon
'Save App Settings
Call SaveAppSettings
End
End Sub

Private Sub mnuIgnore_Click()
If mnuIgnore.Caption = "&Ignore" Then
    Call Ignore(mnuIgnore.Tag)
    mnuIgnore.Caption = "&Unignore"
Else
    Call Unignore(mnuIgnore.Tag)
    mnuIgnore.Caption = "&Ignore"
End If
End Sub

Private Sub mnuIgnoreAll_Click()
For X = 3 To tvwBuddies.Nodes.Count
    If tvwBuddies.Nodes(X).Key <> "NoneOnline" And tvwBuddies.Nodes(X).Key <> "NoneOffline" Then
        Ignore tvwBuddies.Nodes(X).Key
    End If
Next
End Sub

Public Sub mnuMessageAll_Click()
Dim TmpMsg As String
TmpMsg = InputBox("Enter the message.", "Message All")
If TmpMsg <> vbNullString Then
    MessageAll TmpMsg
    lblStatus.Caption = "Message sent to all chat windows."
End If
End Sub

Private Sub mnuMsg_Click()
If mnuMsg.Caption = "&Send an Instant Message" Then
    StartChat mnuMsg.Tag
ElseIf mnuMsg.Caption = "&Block Check" Then
    StartChat mnuMsg.Tag, , True, True
End If
End Sub

Public Sub mnuOpenInbox_Click()
MsnUrlType.Add "00"
MsnSend "URL " & TrialID & " INBOX", TrialID, wskMSN
End Sub

Private Sub mnuProfile_Click()
ShellExecute Me.hwnd, vbNullString, "http://members.msn.com/" & mnuProfile.Tag, vbNullString, vbNullString, 1
End Sub

Private Sub mnuProperties_Click()
ShowBuddyPropPage (mnuProperties.Tag)
End Sub

Private Sub mnuReadme_Click()
ShellExecute Me.hwnd, vbNullString, App.Path & "\Readme.htm", vbNullString, vbNullString, 1
End Sub

Private Sub mnuRemoteControl_Click()
frmRemoteControl.Show , Me
End Sub

Private Sub mnuScript_Click(Index As Integer)
LoadScript mnuScript(Index).Tag
mnuGMScript.Visible = False
mnuStopScript.Visible = True
lblStatus.Caption = "Executing Script..."
End Sub

Private Sub mnuScriptOther_Click()
cdOpen.Filter = "GM Scripts (*.gms)|*.gms"
cdOpen.ShowOpen
If cdOpen.FileName <> vbNullString Then
    LoadScript cdOpen.FileName
    LastDir = Left$(cdOpen.FileName, InStrRev(cdOpen.FileName, "\"))
    cdOpen.FileName = vbNullString
    cdOpen.InitDir = LastDir
    mnuGMScript.Visible = False
    mnuStopScript.Visible = True
    lblStatus.Caption = "Executing Script..."
End If
End Sub

Private Sub mnuSearch_Click()
Temp = InputBox("Search for?", "Search Contact")
If Temp <> vbNullString Then
    cSearch = "*" & Temp & "*"
    For X = 3 To tvwBuddies.Nodes.Count
        If tvwBuddies.Nodes(X).Key <> "NoneOnline" And tvwBuddies.Nodes(X).Key <> "NoneOffline" Then
            If (tvwBuddies.Nodes(X).Key Like cSearch) Or (tvwBuddies.Nodes(X).Text Like cSearch) Or (GetBuddyComment(tvwBuddies.Nodes(X).Key) Like cSearch) Then
                tvwBuddies.Nodes(X).EnsureVisible
                tvwBuddies.Nodes(X).Selected = True
                Exit Sub
            End If
        End If
    Next X
    MsgBox "Contact not found!", vbInformation, "Search Contact"
End If
End Sub

Private Sub mnuSendMessage_Click()
Temp = InputBox("Enter the email of the person you want to message.", "Send Message")
If Trim$(Temp) <> vbNullString Then
    StartChat Temp
End If
End Sub

Private Sub mnuSettings_Click()
frmSettings.Show , Me
End Sub

Public Sub mnuSignIn_Click()
On Error Resume Next
Select Case mnuSignIn.Caption
Case "Sign &In"
    Call picSignIn_Click
Case "&Cancel Sign In"
    Call ResetSockets
Case "Sign &Out"
    LogStatus "Signed out."
    MsnSend "OUT", TrialID, wskMSN
    Call ResetSockets
End Select
End Sub

Private Sub mnuStatusList_Click(Index As Integer)
If mnuStatusList(Index).Checked = False Then ChangeStatus Index
End Sub

Private Sub mnuStatusLogger_Click()
mnuStatusLogger.Checked = Not mnuStatusLogger.Checked
If mnuStatusLogger.Checked = True Then
    lblStatus.Caption = "Status logger activated."
Else
    lblStatus.Caption = "Status logger deactivated."
End If
End Sub

Public Sub mnuStopScript_Click()
tmrGMScript.Enabled = False
Call EndScript
lblStatus.Caption = "Script Stoped."
End Sub

Private Sub mnuViewLog_Click()
    ShellExecute 0, "open", ChatLogDir & "\" & Login & "\" & mnuViewLog.Tag & ".txt", vbNullString, vbNullString, 1
End Sub

Private Sub imgStatus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then Me.PopupMenu mnuStatus
End Sub

Private Sub picTrayIcon_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim Msg As Long
Msg = X / Screen.TwipsPerPixelX
If Msg = WM_LBUTTONDBLCLK Then
    ShowMe frmMain
ElseIf Msg = WM_RBUTTONUP Then
    Me.PopupMenu mnuActions, MenuControlConstants.vbPopupMenuCenterAlign
End If
End Sub

Private Sub tmrAnimator_Timer()
If TrayAnim Mod 2 <> 0 Then
    ChangeIcon picEmpty.Picture.Handle
Else
    ChangeIcon picTrayIcon.Picture.Handle
End If
TrayAnim = TrayAnim + 1
If TrayAnim = 39 Then
    tmrAnimator.Enabled = False
    TrayAnim = 0
End If
End Sub

Public Sub tmrGMScript_Timer()
tmrGMScript.Interval = 500
Call ExecuteScript
End Sub

Private Sub tmrMsnFileKiller_Timer()
tmrMsnFileKiller.Enabled = False
On Error Resume Next
Kill MsnFile(1)
MsnFile.Remove 1
If MsnFile.Count > 0 Then tmrMsnFileKiller.Enabled = True
End Sub

Private Sub tmrPing_Timer()
If SignedIn = True Then
    wskMSN.SendData "PNG" & vbCrLf
End If
End Sub

Private Sub tvwBuddies_Collapse(ByVal Node As MSComctlLib.Node)
    Node.Image = 6
End Sub

Public Sub tvwBuddies_DblClick()
On Error Resume Next
If SignedIn = True Then
    If tvwBuddies.SelectedItem.Index > 2 Then
        If GetBuddyStatus(tvwBuddies.SelectedItem.Key) <> "Offline" Then
            StartChat tvwBuddies.SelectedItem.Key
        End If
    Else
        If tvwBuddies.SelectedItem.Expanded = True Then
            tvwBuddies.SelectedItem.Image = 5
        Else
            tvwBuddies.SelectedItem.Image = 6
        End If
    End If
End If
End Sub

Private Sub tvwBuddies_Expand(ByVal Node As MSComctlLib.Node)
    Node.Image = 5
End Sub

Private Sub tvwBuddies_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And Shift = ShiftConstants.vbAltMask Then
    If tvwBuddies.SelectedItem.Index > 2 And tvwBuddies.SelectedItem.Key <> "NoneOnline" And tvwBuddies.SelectedItem.Key <> "NoneOffline" Then
        ShowBuddyPropPage (tvwBuddies.SelectedItem.Key)
    End If
ElseIf KeyCode = vbKeyReturn Then
    Call tvwBuddies_DblClick
ElseIf KeyCode = vbKeyDelete Then
    If tvwBuddies.SelectedItem.Index > 2 And tvwBuddies.SelectedItem.Key <> "NoneOnline" And tvwBuddies.SelectedItem.Key <> "NoneOffline" Then
        Call DeleteContact(tvwBuddies.SelectedItem.Key)
    End If
End If
End Sub

Private Sub tvwBuddies_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = vbRightButton Then
    If tvwBuddies.SelectedItem.Key = "Online" Or tvwBuddies.SelectedItem.Key = "Offline" Or tvwBuddies.SelectedItem.Key = "NoneOnline" Or tvwBuddies.SelectedItem.Key = "NoneOffline" Then Exit Sub
    mnuSendEmail.Caption = "Send E-mail (" & tvwBuddies.SelectedItem.Key & ")"
    mnuSendEmail.Tag = tvwBuddies.SelectedItem.Key
    mnuMsg.Tag = tvwBuddies.SelectedItem.Key
    If GetBuddyStatus(tvwBuddies.SelectedItem.Key) = "Offline" Then
        mnuMsg.Caption = "&Block Check"
    Else
        mnuMsg.Caption = "&Send an Instant Message"
    End If
    If GetBuddyBlock(tvwBuddies.SelectedItem.Key) = "(Blocked)" Then
        mnuBlock.Caption = "&Unblock"
    Else
        mnuBlock.Caption = "&Block"
    End If
    If IsIgnored(tvwBuddies.SelectedItem.Key) = True Then
        mnuIgnore.Caption = "&Unignore"
    Else
        mnuIgnore.Caption = "&Ignore"
    End If
    mnuCopyNick.Tag = GetBuddyNick(tvwBuddies.SelectedItem.Key)
    mnuCopyEmail.Tag = tvwBuddies.SelectedItem.Key
    mnuBlock.Tag = tvwBuddies.SelectedItem.Key
    mnuIgnore.Tag = tvwBuddies.SelectedItem.Key
    mnuDelete.Tag = tvwBuddies.SelectedItem.Key
    mnuProfile.Tag = tvwBuddies.SelectedItem.Key
    mnuViewLog.Tag = tvwBuddies.SelectedItem.Key
    mnuProperties.Tag = tvwBuddies.SelectedItem.Key
    Me.PopupMenu mnuContact
End If
End Sub

Public Sub txtNick_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn And CBool(GetAsyncKeyState(vbKeyShift)) = False Then
    txtNick.Visible = False
    ChangeNick txtNick.Text
    KeyAscii = 0
ElseIf KeyAscii = vbKeyEscape Then
    txtNick.Visible = False
    txtNick.Text = Nick
End If
End Sub

Public Sub txtNick_LostFocus()
txtNick.Visible = False
txtNick.Text = Nick
End Sub

Public Sub wskMSN_Close()
On Error Resume Next
MsnData = vbNullString
ShowMe frmMain
If tmrGMScript.Enabled = True Then
    tmrGMScript.Enabled = False
End If
SignedIn = False
tmrPing.Enabled = False
InitStatus = True
If mnuStopScript.Visible = True Then
    Call EndScript
End If
picSignIn.Visible = True
If lblStatus.Tag = vbNullString Then
    lblStatus.Caption = vbNullString
Else
    lblStatus.Tag = vbNullString
End If
picMask.Visible = True
ChangeTip "Gilly Messenger"
TrialID = 1
tvwBuddies.Nodes.Clear
mnuStatus.Enabled = False
mnuStatusList(Status).Checked = False
Status = 7
mnuSignIn.Caption = "Sign &In"
mnuSendMessage.Enabled = False
lblNick.Caption = vbNullString
txtNick.Text = vbNullString
mnuOpenInbox.Enabled = False
mnuChatRooms.Enabled = False
mnuEditPassport.Enabled = False
mnuEditProfile.Enabled = False
mnuAddContact.Enabled = False
mnuViewContactsBy.Enabled = False
mnuSaveContactList.Enabled = False
mnuImportContactList.Enabled = False
mnuSearch.Enabled = False
mnuMessageAll.Enabled = False
mnuIgnoreAll.Enabled = False
mnuGMScript.Enabled = False
frmSettings.cmdClearIgnoreList.Enabled = False
frmSettings.cmdClearContactComments.Enabled = False
ResetCollection ContactList
ResetCollection BuddyIgnore
ResetCollection BuddyComment
ResetCollection CallForms
Inbox_Sid = 0
Inbox_Kv = 0
Inbox_Rru = vbNullString
Inbox_MSPAuth = vbNullString
InboxUnread = 0
FolderUnread = 0
NewMail_Folder = vbNullString
DumpMail = vbNullString
lblEmail.Visible = False
Call UpdateEmail
LastAddAlert = vbNullString
LastBlockAlert = vbNullString
LastDeleteAlert = vbNullString
End Sub

Public Sub wskMSN_Connect()
'Start logging in
MsnSend "VER " & TrialID & " MSNP8 CVRO", TrialID, wskMSN
End Sub

Public Sub wskMSN_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim Data As String
wskMSN.GetData Data
MsnData = MsnData & Data
If Right$(MsnData, 2) = vbCrLf Then
    Call Parse(MsnData)
End If
End Sub

Public Sub wskMSN_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
If wskMSN.State <> sckConnected Then
    If SignedIn = True Then
        LogStatus Description
    End If
    lblStatus.Caption = Description
    lblStatus.Tag = "sckError"
    Call ResetSockets
End If
End Sub

Public Sub wskSSL_Close()

' Handle SSL connection
'-----------------------------------------------

    Layer = 0
    wskSSL.Close
    Set SecureSession = Nothing
End Sub

Private Sub wskSSL_Connect()

' Handle SSL connection
'-----------------------------------------------

    Set SecureSession = New clsCrypto
    Call SendClientHello(wskSSL)

End Sub

Private Sub wskSSL_DataArrival(ByVal bytesTotal As Long)

' Decode SSL Information
' Passes result to the ProcessData() sub
'-----------------------------------------------

    'Parse each SSL Record
    Dim TheData As String
    Dim ReachLen As Long

    Do
    
        If SeekLen = 0 Then
            If bytesTotal >= 2 Then
                wskSSL.GetData TheData, vbString, 2
                SeekLen = BytesToLen(TheData)
                bytesTotal = bytesTotal - 2
            Else
                Exit Sub
            End If
        End If
        
        If bytesTotal >= SeekLen Then
            wskSSL.GetData TheData, vbString, SeekLen
            bytesTotal = bytesTotal - SeekLen
        Else
            Exit Sub
        End If
        
        
        Select Case Layer
            Case 0:
                ENCODED_CERT = Mid(TheData, 12, BytesToLen(Mid(TheData, 6, 2)))
                CONNECTION_ID = Right(TheData, BytesToLen(Mid(TheData, 10, 2)))
                Call IncrementRecv
                Call SendMasterKey(wskSSL)
            Case 1:
                TheData = SecureSession.RC4_Decrypt(TheData)
                If Right(TheData, Len(CHALLENGE_DATA)) = CHALLENGE_DATA Then
                    If VerifyMAC(TheData) Then Call SendClientFinish(wskSSL)
                Else
                    wskSSL.Close
                End If
             Case 2:
                TheData = SecureSession.RC4_Decrypt(TheData)
                If VerifyMAC(TheData) = False Then wskSSL.Close
                Layer = 3
                
             Case 3:
                TheData = SecureSession.RC4_Decrypt(TheData)
                If VerifyMAC(TheData) Then Call ProcessData(Mid(TheData, 17))
        End Select
    
        SeekLen = 0

    Loop Until bytesTotal = 0
End Sub

Public Sub ProcessData(strData As String)

strBuffer = strBuffer & strData

' MsgBox strBuffer

End Sub

Function DoSSL(strChallenge As String) As String

' Handles the SSL part of the authentication
'-----------------------------------------------
    
    Dim varLines As Variant
    Dim varURLS As Variant
    
    Dim intCurCookie As Integer
    
    Dim strAuthInfo As String
    Dim strHeader As String
    Dim strLoginServer As String
    Dim strLoginPage As String
    

    
    Dim colURLS As New Collection
    Dim colHeaders As New Collection


    
    'strChallenge = Replace(strChallenge, ",", "&")
    
'Connect to NEXUS:
'--------------------------------------------------
    SSL_Error = False
    strBuffer = ""
    
    wskSSL.Close
    wskSSL.Connect "nexus.passport.com", 443
    
    Do Until wskSSL.State = sckConnected
        DoEvents
        If SSL_Error = True Then
            DoSSL = "Error"
            Exit Function
        End If
    Loop
    
    ' Wait for the SSL layer to be established:
    
    Do Until Layer = 3
        DoEvents
        If SSL_Error = True Then
            DoSSL = "Error"
            Exit Function
        End If
    Loop

    'Obtain login information from NEXUS:
    
    Call SSLSend(wskSSL, "GET /rdr/pprdr.asp HTTP/1.0" & vbCrLf & vbCrLf)
    
    Do Until InStr(1, strBuffer, vbCrLf & vbCrLf) <> 0
        DoEvents
        If SSL_Error = True Then
            DoSSL = "Error"
            Exit Function
        End If
    Loop
    
    wskSSL.Close
    
'--------------------------------------------------
'Done with NEXUS
    
    
    
'Begin processing data from NEXUS:
'--------------------------------------------------
    
    intCurCookie = 0
    varLines = Split(strBuffer, vbCrLf)
    
    ' Search for the header "PasswordURLs:"
    
        For intCount = LBound(varLines) To UBound(varLines)
        
            ' Add the values for "PasswordURLs:" to a collection:
            
            If Left$(CStr(varLines(intCount)), InStr(1, varLines(intCount), " ")) = "PassportURLs: " Then
                colHeaders.Add Right$(CStr(varLines(intCount)), Len(varLines(intCount)) - InStr(1, varLines(intCount), " ")), Left(varLines(intCount), InStr(1, varLines(intCount), " "))
                Exit For
            End If
            
        Next intCount
        
    
    varURLS = Split(colHeaders.Item("PassportURLs: "), ",")
    
    For intCount = LBound(varURLS) To UBound(varURLS)
        colURLS.Add Right(varURLS(intCount), Len(varURLS(intCount)) - InStr(1, varURLS(intCount), "=")), Left(varURLS(intCount), InStr(1, varURLS(intCount), "="))
    Next intCount
    
    'Get the server and page for logging in:

    strLoginServer = Left$(colURLS("DALogin="), InStr(1, colURLS("DALogin="), "/") - 1)
    strLoginPage = Right$(colURLS("DALogin="), Len(colURLS("DALogin=")) - InStr(1, colURLS("DALogin="), "/") + 1)
    
'--------------------------------------------------
'End processing
    

    
ConnectLogin:

'Connect to login server
'--------------------------------------------------

    strBuffer = ""
    
    ' Layer resembles the state of the SSL connection:
    Layer = 0
    
    wskSSL.Close
    wskSSL.Connect strLoginServer, 443
    Do Until wskSSL.State <> sckConnected
        DoEvents
        If SSL_Error = True Then
            DoSSL = "Error"
            Exit Function
        End If
    Loop
    ' Wait for the SSL layer to be established:
    
    Do Until Layer = 3
        DoEvents
        If SSL_Error = True Then
            DoSSL = "Error"
            Exit Function
        End If
    Loop

    strHeader = "GET " & strLoginPage & " HTTP/1.1" & vbCrLf & _
                "Authorization: Passport1.4 OrgVerb=GET,OrgURL=http%3A%2F%2Fmessenger%2Emsn%2Ecom,sign-in=" & Replace(Login, "@", "%40") & ",pwd=" & URLEncode(Password) & "," & strChallenge & _
                "User-Agent: MSMSGS" & vbCrLf & _
                "Host: loginnet.passport.com" & vbCrLf & _
                "Connection: Keep-Alive" & vbCrLf & _
                "Cache-Control: no-cache" & vbCrLf & vbCrLf

    Call SSLSend(wskSSL, strHeader)

    ' Wait for the header to be recieved
    
    Do Until InStr(1, strBuffer, vbCrLf & vbCrLf) <> 0
        DoEvents
        If SSL_Error = True Then
            DoSSL = "Error"
            Exit Function
        End If
    Loop
    
    If InStr(strBuffer, "HTTP/1.1 401 Unauthorized") > 0 Then
        Call ResetSockets
        DoSSL = "False"
        Exit Function
    End If
    
    Dim strHeaderValue As String

    strHeaderValue = GetHeader("authentication-info:", strBuffer)
    
    If RequiresRedirect(strHeaderValue) = True Then
    
        strHeaderValue = GetHeader("location:", strBuffer)
        
        lngCharPos = InStr(strHeaderValue, "://")
        
        If (LCase$(Left$(strHeaderValue, lngCharPos - 1)) = "https") Then
        
            strLoginServer = Mid$(strHeaderValue, lngCharPos + 3, InStr(lngCharPos + 3, strHeaderValue, "/") - (lngCharPos + 3))
            strLoginPage = Right$(strHeaderValue, Len(strHeaderValue) - (InStr(lngCharPos + 3, strHeaderValue, "/") - 1))
            
            GoTo ConnectLogin
            
        End If
    
    Else
    
        DoSSL = ParseHash(strHeaderValue)
        wskSSL.Close
        Exit Function
        
    End If

'--------------------------------------------------
'Done with login server

End Function


Function GetHeader(strHeader As String, strData As String) As String

' Returns the value of a header-property
'-----------------------------------------------

Dim intCount As Integer
Dim varLines As Variant
Dim lngCharPos As Long
Dim strCurHeader As String

varLines = Split(strData, vbCrLf)

For intCount = LBound(varLines) To UBound(varLines)

If Len(varLines(intCount)) = 0 Then Exit For

    strCurHeader = varLines(intCount)
    lngCharPos = InStr(strCurHeader, " ")
    
    If LCase(Left(strCurHeader, lngCharPos - 1)) = LCase(strHeader) Then
        GetHeader = Right(strCurHeader, Len(strCurHeader) - lngCharPos)
        Exit Function
    End If
    

Next intCount

End Function

Function RequiresRedirect(strData As String) As Boolean

' Checks whether it's necessary to redirect to
' another server (using 'da-status' property)
'-----------------------------------------------

Dim intCount As Integer
Dim varProps As Variant
Dim lngCharPos As Long
Dim strCurItem As String
Dim strPropName As String
Dim strPropValue As String

lngCharPos = InStr(strData, " ")

If Left(strData, lngCharPos - 1) = "Passport1.4" Then

    strData = Right(strData, Len(strData) - lngCharPos)
    varProps = Split(strData, ",")
    
    For intCount = LBound(varProps) To UBound(varProps)
    
        strCurItem = varProps(intCount)
        lngCharPos = InStr(strCurItem, "=")
        
        strPropName = Left(strCurItem, lngCharPos - 1)
        strPropValue = Right(strCurItem, Len(strCurItem) - lngCharPos)
    
        If LCase$(strPropName) = "da-status" And LCase$(strPropValue) = "redir" Then
        
            RequiresRedirect = True
            Exit Function
            
        ElseIf LCase$(strPropName) = "da-status" And LCase$(strPropValue) = "success" Then
        
            RequiresRedirect = False
            Exit Function
        
        End If
        
    Next intCount

End If

End Function

Function ParseHash(strHeader As String) As String

' Returns the hash (from-pp) if the login has
' completed succesfully.
'-----------------------------------------------

Dim intCount As Integer
Dim varProps As Variant
Dim lngCharPos As Long
Dim strCurItem As String
Dim strPropName As String
Dim strPropValue As String

    varProps = Split(strHeader, ",")
    
    For intCount = LBound(varProps) To UBound(varProps)
    
        strCurItem = varProps(intCount)
        lngCharPos = InStr(strCurItem, "=")
        
        strPropName = Left(strCurItem, lngCharPos - 1)
        strPropValue = Right(strCurItem, Len(strCurItem) - lngCharPos)
    
        If LCase$(strPropName) = "from-pp" Then
        
            ParseHash = strPropValue
            'MsgBox ParseHash
            ParseHash = Left(ParseHash, Len(ParseHash) - 1)
            ParseHash = Right(ParseHash, Len(ParseHash) - 1)
            
            Exit Function
        
        End If
        
    Next intCount

End Function


Private Sub wskSSL_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
If wskSSL.State <> sckConnected Then
    SSL_Error = True
    If SignedIn = True Then
        LogStatus Description
    End If
    lblStatus.Caption = Description
    lblStatus.Tag = "sckError"
    Call ResetSockets
End If
End Sub

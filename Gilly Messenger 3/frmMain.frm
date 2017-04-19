VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Gilly Messenger"
   ClientHeight    =   5595
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   3765
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   ScaleHeight     =   373
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   251
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrGMScript_Events 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1080
      Top             =   3600
   End
   Begin VB.Timer tmrNewsScroller1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2040
      Top             =   3600
   End
   Begin MSWinsockLib.Winsock wskNews 
      Left            =   1440
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrNewsScroller2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2040
      Top             =   3120
   End
   Begin VB.PictureBox picNews 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FCF8F6&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   0
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   249
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4560
      Visible         =   0   'False
      Width           =   3735
      Begin VB.Label lblNews 
         AutoSize        =   -1  'True
         BackColor       =   &H00FCF8F6&
         Caption         =   "Join CrackSoft forums and stay updated."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00814D3C&
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   60
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   3465
      End
   End
   Begin VB.Timer tmrAutoIdle 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1560
      Top             =   3600
   End
   Begin VB.Timer tmrReconnect 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   600
      Top             =   3120
   End
   Begin MSComctlLib.ImageList imglstTrayIcons 
      Left            =   2640
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E64
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1998
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F32
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock wskSSL 
      Left            =   960
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wskNS 
      Left            =   480
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   0
      ScaleHeight     =   281
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   249
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   3735
      Begin VB.Timer tmrGMScript_Main 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   600
         Top             =   3600
      End
      Begin VB.PictureBox picSignIn 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   1080
         MouseIcon       =   "frmMain.frx":24CC
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":2D96
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   82
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1230
      End
      Begin VB.Timer tmrTrayAnim 
         Enabled         =   0   'False
         Interval        =   250
         Left            =   1560
         Tag             =   "20"
         Top             =   3120
      End
      Begin VB.PictureBox picSignInProgress 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   960
         ScaleHeight     =   10
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   93
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1920
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Timer tmrPing 
         Enabled         =   0   'False
         Interval        =   60000
         Left            =   1080
         Top             =   3120
      End
      Begin VB.PictureBox picTrayIcon 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   240
         ScaleHeight     =   180
         ScaleWidth      =   180
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   480
         Visible         =   0   'False
         Width           =   240
      End
      Begin MSComDlg.CommonDialog CommonDialog 
         Left            =   2640
         Top             =   1200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList imglstEmoticons 
         Left            =   2640
         Top             =   2520
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
               Picture         =   "frmMain.frx":4AE8
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":50B8
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":5688
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":5C58
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":6228
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":66F0
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":6CC0
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":7290
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":7860
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":7D64
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":8334
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":8904
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":8ED4
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":94A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":9A74
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":A044
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":A614
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":ABE4
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":B1B4
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":B784
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":BD54
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":C21C
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":C6E4
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":CCB4
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":D17C
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":D644
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":DC14
               Key             =   ""
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":E1E4
               Key             =   ""
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":E7B4
               Key             =   ""
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":ED84
               Key             =   ""
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":F354
               Key             =   ""
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":F924
               Key             =   ""
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":FEF4
               Key             =   ""
            EndProperty
            BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":104C4
               Key             =   ""
            EndProperty
            BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":10A94
               Key             =   ""
            EndProperty
            BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":10F5C
               Key             =   ""
            EndProperty
            BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1138C
               Key             =   ""
            EndProperty
            BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1195C
               Key             =   ""
            EndProperty
            BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":11F2C
               Key             =   ""
            EndProperty
            BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":124FC
               Key             =   ""
            EndProperty
            BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":12ACC
               Key             =   ""
            EndProperty
            BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1309C
               Key             =   ""
            EndProperty
            BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1366C
               Key             =   ""
            EndProperty
            BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":13C3C
               Key             =   ""
            EndProperty
            BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1420C
               Key             =   ""
            EndProperty
            BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":147DC
               Key             =   ""
            EndProperty
            BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":14DAC
               Key             =   ""
            EndProperty
            BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1537C
               Key             =   ""
            EndProperty
            BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1594C
               Key             =   ""
            EndProperty
            BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":15F1C
               Key             =   ""
            EndProperty
            BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":164EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":16ABC
               Key             =   ""
            EndProperty
            BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1708C
               Key             =   ""
            EndProperty
            BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1765C
               Key             =   ""
            EndProperty
            BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":17C2C
               Key             =   ""
            EndProperty
            BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":181FC
               Key             =   ""
            EndProperty
            BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":187CC
               Key             =   ""
            EndProperty
            BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":18D9C
               Key             =   ""
            EndProperty
            BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1936C
               Key             =   ""
            EndProperty
            BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1993C
               Key             =   ""
            EndProperty
            BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":19F0C
               Key             =   ""
            EndProperty
            BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1A4DC
               Key             =   ""
            EndProperty
            BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1AAAC
               Key             =   ""
            EndProperty
            BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1B07C
               Key             =   ""
            EndProperty
            BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1B64C
               Key             =   ""
            EndProperty
            BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1BC1C
               Key             =   ""
            EndProperty
            BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1C1EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1C7BC
               Key             =   ""
            EndProperty
            BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1CC84
               Key             =   ""
            EndProperty
            BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1D254
               Key             =   ""
            EndProperty
            BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1D5E8
               Key             =   ""
            EndProperty
            BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1D948
               Key             =   ""
            EndProperty
            BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1DEDC
               Key             =   ""
            EndProperty
            BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1E3A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1E86C
               Key             =   ""
            EndProperty
            BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1ED34
               Key             =   ""
            EndProperty
            BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1F1FC
               Key             =   ""
            EndProperty
            BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1F7CC
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imglstStatus 
         Left            =   2640
         Top             =   1800
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   15
         ImageHeight     =   16
         MaskColor       =   16711935
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1FC94
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1FFE8
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2033C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":20690
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":209E4
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":20DB8
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":210DC
               Key             =   ""
            EndProperty
         EndProperty
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
         Left            =   1200
         TabIndex        =   10
         Top             =   1560
         UseMnemonic     =   0   'False
         Width           =   990
      End
      Begin VB.Label lblWelcome 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "welcome"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E9CAB1&
         Height          =   360
         Left            =   120
         TabIndex        =   7
         Top             =   90
         UseMnemonic     =   0   'False
         Width           =   1320
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
         TabIndex        =   6
         Top             =   180
         UseMnemonic     =   0   'False
         Width           =   435
      End
   End
   Begin VB.PictureBox picStatus 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   249
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5040
      Width           =   3735
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00814D3C&
         Height          =   210
         Left            =   120
         TabIndex        =   3
         Top             =   90
         UseMnemonic     =   0   'False
         Width           =   3495
      End
   End
   Begin MSComctlLib.TreeView tvwContacts 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   7011
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      Style           =   1
      ImageList       =   "imglstStatus"
      Appearance      =   0
   End
   Begin VB.TextBox txtNick 
      Appearance      =   0  'Flat
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
      Height          =   225
      Left            =   390
      MaxLength       =   129
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Image imgStatus 
      Height          =   255
      Left            =   120
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lblEmail 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Goto my e-mail inbox"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   375
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   750
      UseMnemonic     =   0   'False
      Width           =   1755
   End
   Begin VB.Image imgEmail 
      Height          =   240
      Left            =   60
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":21400
      Top             =   720
      Width           =   270
   End
   Begin VB.Label lblNick 
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
      Height          =   225
      Left            =   390
      TabIndex        =   1
      Top             =   270
      UseMnemonic     =   0   'False
      Width           =   3135
   End
   Begin VB.Image imgTopRight 
      Height          =   450
      Left            =   2160
      Picture         =   "frmMain.frx":217C2
      Top             =   150
      Width           =   1560
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFile_SignIn 
         Caption         =   "S&ign In..."
      End
      Begin VB.Menu mnuFile_SignOut 
         Caption         =   "Sig&n Out"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFile_Seperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_MyStatus 
         Caption         =   "&My Status"
         Enabled         =   0   'False
         Begin VB.Menu mnuFile_MyStatus_Status 
            Caption         =   "&Online"
            Index           =   0
         End
         Begin VB.Menu mnuFile_MyStatus_Status 
            Caption         =   "&Busy"
            Index           =   1
         End
         Begin VB.Menu mnuFile_MyStatus_Status 
            Caption         =   "B&e Right Back"
            Index           =   2
         End
         Begin VB.Menu mnuFile_MyStatus_Status 
            Caption         =   "&Away"
            Index           =   3
         End
         Begin VB.Menu mnuFile_MyStatus_Status 
            Caption         =   "On The &Phone"
            Index           =   4
         End
         Begin VB.Menu mnuFile_MyStatus_Status 
            Caption         =   "Out To &Lunch"
            Index           =   5
         End
         Begin VB.Menu mnuFile_MyStatus_Status 
            Caption         =   "&Idle"
            Index           =   6
         End
         Begin VB.Menu mnuFile_MyStatus_Status 
            Caption         =   "Appear O&ffline"
            Index           =   7
         End
      End
      Begin VB.Menu mnuFile_Seperator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Goto 
         Caption         =   "&Go To"
         Begin VB.Menu mnuFile_Goto_MsnHome 
            Caption         =   "MSN &Home"
         End
         Begin VB.Menu mnuFile_Goto_MyEmailInbox 
            Caption         =   "My E-mail In&box"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuFile_Goto_MyProfile 
            Caption         =   "My Pr&ofile"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuFile_Goto_MyPassport 
            Caption         =   "My &Passport"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuFile_Goto_Chatrooms 
            Caption         =   "&Chat Rooms"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuFile_Goto_MsnToday 
            Caption         =   "&MSN Today"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuFile_Seperator4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_SendAFileOrPhoto 
         Caption         =   "Send a &File or Photo..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFile_OpenReceivedFiles 
         Caption         =   "&Open Received Files"
      End
      Begin VB.Menu mnuFile_OpenMessageHistory 
         Caption         =   "Open Message &History..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFile_Seperator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Close 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuFile_Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuContacts 
      Caption         =   "&Contacts"
      Begin VB.Menu mnuContacts_AddAContact 
         Caption         =   "&Add a Contact..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuContacts_SearchForAContact 
         Caption         =   "&Search for a Contact"
         Enabled         =   0   'False
         Begin VB.Menu mnuContacts_SearchForAContact_ContactList 
            Caption         =   "&Contact List"
         End
         Begin VB.Menu mnuContacts_SearchForAContact_AdvancedSearch 
            Caption         =   "&Advanced Search"
         End
         Begin VB.Menu mnuContacts_SearchForAContact_SearchByInterest 
            Caption         =   "Search by &Interest"
         End
      End
      Begin VB.Menu mnuContacts_Seperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContacts_ManageContacts 
         Caption         =   "&Manage Contacts"
         Enabled         =   0   'False
         Begin VB.Menu mnuContacts_ManageContacts_ViewContactsBy 
            Caption         =   "View &Contacts By"
            Begin VB.Menu mnuContacts_ManageContacts_ViewContactsBy_DisplayName 
               Caption         =   "&Display name"
            End
            Begin VB.Menu mnuContacts_ManageContacts_ViewContactsBy_EmailAddress 
               Caption         =   "&E-mail address"
            End
         End
      End
      Begin VB.Menu mnuContacts_ManageGroups 
         Caption         =   "Manage Groups"
         Enabled         =   0   'False
         Begin VB.Menu mnuContacts_ManageGroups_CreateNewGroup 
            Caption         =   "Create &New Group"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuContacts_ManageGroups_DeleteAGroup 
            Caption         =   "&Delete a Group"
            Enabled         =   0   'False
            Begin VB.Menu mnuContacts_ManageGroups_DeleteAGroup_Group 
               Caption         =   "(Group Name)"
               Enabled         =   0   'False
               Index           =   0
            End
         End
         Begin VB.Menu mnuContacts_ManageGroups_RenameAGroup 
            Caption         =   "&Rename a Group"
            Enabled         =   0   'False
            Begin VB.Menu mnuContacts_ManageGroups_RenameAGroup_Group 
               Caption         =   "(Group Name)"
               Enabled         =   0   'False
               Index           =   0
            End
         End
         Begin VB.Menu mnuContacts_ManageGroups_GroupOfflineContactsTogether 
            Caption         =   "Gro&up Offline Contacts Together"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuContacts_SortContactsBy 
         Caption         =   "Sort &Contacts By"
         Enabled         =   0   'False
         Begin VB.Menu mnuContacts_SortContactsBy_Groups 
            Caption         =   "&Groups"
         End
         Begin VB.Menu mnuContacts_SortContactsBy_OnlineOffline 
            Caption         =   "&Online / Offline"
         End
      End
      Begin VB.Menu mnuContacts_Seperator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContacts_SaveContactList 
         Caption         =   "Save Contact &List..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuContacts_ImportContactsFromASavedFile 
         Caption         =   "Im&port Contacts from a Saved File..."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuActions 
      Caption         =   "&Actions"
      Begin VB.Menu mnuActions_SendAnInstantMessage 
         Caption         =   "&Send an Instant Message..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuActions_SendAFileOrPhoto 
         Caption         =   "Send a &File or Photo..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuActions_Seperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActions_SendEmail 
         Caption         =   "Send &E-mail..."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuTools_ChangeDisplayPic 
         Caption         =   "Change Displa&y Picture..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuTools_AutoMessage 
         Caption         =   "&Auto Message"
      End
      Begin VB.Menu mnuTools_MessageAll 
         Caption         =   "&Message All"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuTools_IgnoreAll 
         Caption         =   "&Ignore All"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuTools_ChatBot 
         Caption         =   "&Chat Bot"
         Begin VB.Menu mnuTools_ChatBot_Bot 
            Caption         =   "No Bot Available"
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu mnuTools_ChatBot_Seperator1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTools_ChatBot_Other 
            Caption         =   "&Other..."
         End
      End
      Begin VB.Menu mnuTools_GMScript 
         Caption         =   "&GM Script"
         Enabled         =   0   'False
         Begin VB.Menu mnuTools_GMScript_Script 
            Caption         =   "No Script Available"
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu mnuTools_GMScript_Seperator1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTools_GMScript_Other 
            Caption         =   "&Other..."
            Index           =   0
         End
      End
      Begin VB.Menu mnuTools_RemoteControl 
         Caption         =   "&Remote Control"
      End
      Begin VB.Menu mnuTools_Options 
         Caption         =   "&Options"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp_Readme 
         Caption         =   "&Readme"
      End
      Begin VB.Menu mnuHelp_Seperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelp_CrackSoftWebsite 
         Caption         =   "CrackSoft &Website"
      End
      Begin VB.Menu mnuHelp_CrackSoftForums 
         Caption         =   "CrackSoft &Forums"
      End
      Begin VB.Menu mnuHelp_Seperator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelp_AboutGillyMessenger 
         Caption         =   "&About Gilly Messenger"
      End
   End
   Begin VB.Menu mnuGroup 
      Caption         =   "Group"
      Visible         =   0   'False
      Begin VB.Menu mnuGroup_RenameGroup 
         Caption         =   "&Rename Group"
      End
      Begin VB.Menu mnuGroup_DeleteGroup 
         Caption         =   "&Delete Group"
      End
      Begin VB.Menu mnuGroup_SaveGroupToAFile 
         Caption         =   "&Save Group to a File..."
      End
      Begin VB.Menu mnuGroup_Seperator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGroup_CreateNewGroup 
         Caption         =   "Create &New Group"
      End
   End
   Begin VB.Menu mnuContact 
      Caption         =   "Contact"
      Visible         =   0   'False
      Begin VB.Menu mnuContact_SendAnInstantMessage 
         Caption         =   "&Send an Instant Message"
      End
      Begin VB.Menu mnuContact_SendAFileOrPhoto 
         Caption         =   "Send a &File or Photo..."
      End
      Begin VB.Menu mnuContact_Seperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContact_CopyEmail 
         Caption         =   "Copy Email"
      End
      Begin VB.Menu mnuContact_CopyNick 
         Caption         =   "Copy &Nick"
      End
      Begin VB.Menu mnuContact_SendEmail 
         Caption         =   "Send &Email (Email)"
      End
      Begin VB.Menu mnuContact_Seperator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContact_OpenMessageHistory 
         Caption         =   "Open Message &History"
      End
      Begin VB.Menu mnuContact_Seperator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContact_Block 
         Caption         =   "&Block"
      End
      Begin VB.Menu mnuContact_Hide 
         Caption         =   "Hi&de"
      End
      Begin VB.Menu mnuContact_Ignore 
         Caption         =   "&Ignore"
      End
      Begin VB.Menu mnuContact_PopupFilter 
         Caption         =   "Add to Pop&up Filter"
      End
      Begin VB.Menu mnuContact_SoundFilter 
         Caption         =   "Add to S&ound Filter"
      End
      Begin VB.Menu mnuContact_CopyContactTo 
         Caption         =   "&Copy Contact to"
         Begin VB.Menu mnuContact_CopyContactTo_Group 
            Caption         =   "(Group)"
            Enabled         =   0   'False
            Index           =   0
         End
      End
      Begin VB.Menu mnuContact_MoveContactTo 
         Caption         =   "M&ove Contact to"
         Begin VB.Menu mnuContact_MoveContactTo_Group 
            Caption         =   "(Group)"
            Enabled         =   0   'False
            Index           =   0
         End
      End
      Begin VB.Menu mnuContact_RemoveContactFromGroup 
         Caption         =   "&Remove Contact from Group"
      End
      Begin VB.Menu mnuContact_DeleteContact 
         Caption         =   "&Delete Contact"
      End
      Begin VB.Menu mnuContact_ViewProfile 
         Caption         =   "&View Profile"
      End
      Begin VB.Menu mnuContact_Properties 
         Caption         =   "&Properties"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents objMSN_NS As clsMSN_NS
Attribute objMSN_NS.VB_VarHelpID = -1

Private InboxUnread As Integer, FoldersUnread As Integer

Private FirstStatus As Boolean
Private WindowLoaded As Boolean
Private NewsData As String
Private NewsLines() As String
Private NewsPointer As Integer
Private NewsInterval As Integer
Private LastNsError As String
Private LastAddAlert As String
Private LastDelAlert As String
Private LastBlockAlert As String

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    LastActive = Timer
    On Error Resume Next
    
    If objMSN_NS.State = NsState_Disconnected Then
        If KeyCode = vbKeyS And Shift = vbAltMask Then
            Call mnuFile_SignIn_Click
        End If
    ElseIf objMSN_NS.State = NsState_SignedIn Then
        If KeyCode = vbKeyF3 Then
            Call SearchForAContact
        ElseIf KeyCode = vbKeyN And Shift = vbAltMask Then
            Call lblNick_Click
        ElseIf KeyCode = vbKeyE And Shift = vbAltMask Then
            Call lblEmail_Click
        End If
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    Set objMSN_NS = New clsMSN_NS
    objMSN_NS.NsSocket = wskNS
    objMSN_NS.SslSocket = wskSSL
    Set DpTransfers = New Collection
    
    Set IMWindowBackground = LoadResPicture("IMWindowBackground", vbResBitmap)
    Set IMWindowTopLeft = LoadResPicture("IMWindowTopLeft", vbResBitmap)
    Set IMWindowTopMid = LoadResPicture("IMWindowTopMid", vbResBitmap)
    Set IMWindowTopRight = LoadResPicture("IMWindowTopRight", vbResBitmap)
    
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
    
    Dim i As Integer
    For i = 0 To 76
        frmEmoticons.imgEmoticon(i).Picture = imglstEmoticons.ListImages(i + 1).Picture
    Next
    
    IMWindowCommands(0, 0) = "/online": IMWindowCommands(0, 1) = "False"
    IMWindowCommands(1, 0) = "/vanish": IMWindowCommands(1, 1) = "False"
    IMWindowCommands(2, 0) = "/invite": IMWindowCommands(2, 1) = "True"
    IMWindowCommands(3, 0) = "/block": IMWindowCommands(3, 1) = "False"
    IMWindowCommands(4, 0) = "/unblock": IMWindowCommands(4, 1) = "False"
    IMWindowCommands(5, 0) = "/ignore": IMWindowCommands(5, 1) = "False"
    IMWindowCommands(6, 0) = "/unignore": IMWindowCommands(6, 1) = "False"
    IMWindowCommands(7, 0) = "/profile": IMWindowCommands(7, 1) = "False"
    IMWindowCommands(8, 0) = "/properties": IMWindowCommands(8, 1) = "False"
    IMWindowCommands(9, 0) = "/list": IMWindowCommands(9, 1) = "False"
    IMWindowCommands(10, 0) = "/view log": IMWindowCommands(10, 1) = "False"
    IMWindowCommands(11, 0) = "/find": IMWindowCommands(11, 1) = "True"
    IMWindowCommands(12, 0) = "/nick": IMWindowCommands(12, 1) = "True"
    IMWindowCommands(13, 0) = "/color": IMWindowCommands(13, 1) = "True"
    IMWindowCommands(14, 0) = "/automsg": IMWindowCommands(14, 1) = "True"
    IMWindowCommands(15, 0) = "/list online": IMWindowCommands(15, 1) = "False"
    IMWindowCommands(16, 0) = "/busy": IMWindowCommands(16, 1) = "False"
    IMWindowCommands(17, 0) = "/brb": IMWindowCommands(17, 1) = "False"
    IMWindowCommands(18, 0) = "/away": IMWindowCommands(18, 1) = "False"
    IMWindowCommands(19, 0) = "/phone": IMWindowCommands(19, 1) = "False"
    IMWindowCommands(20, 0) = "/lunch": IMWindowCommands(20, 1) = "False"
    IMWindowCommands(21, 0) = "/idle": IMWindowCommands(21, 1) = "False"
    IMWindowCommands(22, 0) = "/hide": IMWindowCommands(22, 1) = "False"
    IMWindowCommands(23, 0) = "/chat": IMWindowCommands(23, 1) = "True"
    IMWindowCommands(24, 0) = "/msg": IMWindowCommands(24, 1) = "True"
    IMWindowCommands(25, 0) = "/msgall": IMWindowCommands(25, 1) = "True"
    IMWindowCommands(26, 0) = "/sendfile": IMWindowCommands(26, 1) = "True"
    IMWindowCommands(27, 0) = "/signout": IMWindowCommands(27, 1) = "False"
    IMWindowCommands(28, 0) = "/signin": IMWindowCommands(28, 1) = "True"
    IMWindowCommands(29, 0) = "/ver": IMWindowCommands(29, 1) = "False"
    IMWindowCommands(30, 0) = "/msgr": IMWindowCommands(30, 1) = "False"
    IMWindowCommands(31, 0) = "/close": IMWindowCommands(31, 1) = "False"
    IMWindowCommands(32, 0) = "/script": IMWindowCommands(32, 1) = "True"
    IMWindowCommands(33, 0) = "/execute": IMWindowCommands(33, 1) = "True"
    IMWindowCommands(34, 0) = "/exit": IMWindowCommands(34, 1) = "False"
    IMWindowCommands(35, 0) = "/comment": IMWindowCommands(35, 1) = "True"
    IMWindowCommands(36, 0) = "/view comment": IMWindowCommands(36, 1) = "False"
    IMWindowCommands(37, 0) = "/buzz": IMWindowCommands(37, 1) = "False"
    IMWindowCommands(38, 0) = "/fakenick": IMWindowCommands(38, 1) = "True"
    IMWindowCommands(39, 0) = "/imitate": IMWindowCommands(39, 1) = "False"
    IMWindowCommands(40, 0) = "/view info": IMWindowCommands(40, 1) = "False"
    IMWindowCommands(41, 0) = "/font": IMWindowCommands(41, 1) = "True"
    IMWindowCommands(42, 0) = "/customnick": IMWindowCommands(42, 1) = "True"
    IMWindowCommands(43, 0) = "/addcontact": IMWindowCommands(43, 1) = "True"
    IMWindowCommands(44, 0) = "/bot": IMWindowCommands(44, 1) = "True"
    IMWindowCommands(45, 0) = "/email": IMWindowCommands(45, 1) = "False"
    IMWindowCommands(46, 0) = "/text": IMWindowCommands(46, 1) = "True"
    
    MyStatus = msnStatus_Offline
    imgEmail.MouseIcon = picSignIn.MouseIcon
    lblEmail.MouseIcon = picSignIn.MouseIcon
    lblNews.MouseIcon = picSignIn.MouseIcon
    
    Call AddTrayIcon
    
    objMSN_NS.Server = GetSettingX("Server Settings", "IPAddress", "messenger.hotmail.com")
    objMSN_NS.Port = Val(GetSettingX("Server Settings", "Port", 1863))
    
    AlertOnContactOnline = GetSettingX("App Settings", "Alert OnContactOnline", True)
    AlertOnMessageReceived = GetSettingX("App Settings", "Alert OnMessageReceived", True)
    AlertOnEmailReceived = GetSettingX("App Settings", "Alert OnEmailReceived", True)
    BlockAlert = GetSettingX("App Settings", "Block Alert", True)
    SoundAlerts = GetSettingX("App Settings", "Sound Alerts", True)
    
    Dim strSetting As String
    strSetting = GetSettingX("App Settings", "Online Sound", "1 " & ReadRegKey("HKEY_CURRENT_USER\AppEvents\Schemes\Apps\MSNMSGR\MSNMSGR_ContactOnline\.Current\"))
    boolOnlineSound = CBool(Split(strSetting)(0))
    strOnlineSound = Right$(strSetting, Len(strSetting) - 2)
    strSetting = GetSettingX("App Settings", "Offline Sound", "1")
    boolOfflineSound = CBool(Split(strSetting)(0))
    strOfflineSound = Right$(strSetting, Len(strSetting) - 2)
    strSetting = GetSettingX("App Settings", "Typing Sound", "1")
    boolTypingSound = CBool(Split(strSetting)(0))
    strTypingSound = Right$(strSetting, Len(strSetting) - 2)
    strSetting = GetSettingX("App Settings", "Message Sound", "1 " & ReadRegKey("HKEY_CURRENT_USER\AppEvents\Schemes\Apps\MSNMSGR\MSNMSGR_NewMessage\.Current\"))
    boolMessageSound = CBool(Split(strSetting)(0))
    strMessageSound = Right$(strSetting, Len(strSetting) - 2)
    strSetting = GetSettingX("App Settings", "Email Sound", "1 " & ReadRegKey("HKEY_CURRENT_USER\AppEvents\Schemes\Apps\MSNMSGR\MSNMSGR_NewAlert\.Current\"))
    boolEmailSound = CBool(Split(strSetting)(0))
    strEmailSound = Right$(strSetting, Len(strSetting) - 2)
    strSetting = GetSettingX("App Settings", "Alert Sound", "1 " & ReadRegKey("HKEY_CURRENT_USER\AppEvents\Schemes\Apps\MSNMSGR\MSNMSGR_NewMail\.Current\"))
    boolAlertSound = CBool(Split(strSetting)(0))
    strAlertSound = Right$(strSetting, Len(strSetting) - 2)
    BlockAlertsOnFullScrApp = CBool(GetSettingX("App Settings", "BlockAlerts Onnot fullscrapp", True))
   
    MainWindowWidth = Val(GetSettingX("App Settings", "MainWindow Width", Me.Width))
    MainWindowWidth = IIf(MainWindowWidth = 0, Me.Width, MainWindowWidth)
    MainWindowHeight = Val(GetSettingX("App Settings", "MainWindow Height", Me.Height))
    MainWindowHeight = IIf(MainWindowHeight = 0, Me.Height, MainWindowHeight)
    MainWindowMax = GetSettingX("App Settings", "MainWindow Max", False)
    
    ReceivedFilesFolder = GetSettingX("App Settings", "ReceivedFiles Folder", GetSpecialFolder(CSIDL_PERSONAL) & "\My Received Files\")
    FTPPort = GetSettingX("App Settings", "FTP Port", 1863)
    
    Dim RcAccounts() As String
    RcAccounts = GetAllSettings("Gilly Messenger", "RC Accounts")
    If Not ArraySize(RcAccounts) = -1 Then
        Dim RcAccount As Collection, Attrs() As String
        For i = 0 To UBound(RcAccounts)
            Set RcAccount = New Collection
            Attrs = Split(RcAccounts(i, 1))
            RcAccount.Add RcAccounts(i, 0), "login"
            RcAccount.Add XorDecrypt(Attrs(0), RcAccounts(i, 0)), "password"
            RcAccount.Add Attrs(1), "dirBrowsing"
            RcAccount.Add Attrs(2), "msgrControl"
            RcAccount.Add Attrs(3), "shellCommands"
            RC_Accounts.Add RcAccount, RcAccounts(i, 0)
            Set RcAccount = Nothing
        Next
    End If
    
    Me.Width = MainWindowWidth
    Me.Height = MainWindowHeight
    
    If MainWindowMax Then
        Me.WindowState = vbMaximized
    End If

    Transparency = Val(GetSettingX("App Settings", "Transparency", 0))
    
    If Not Transparency = 0 Then
        SetTransparency Me, Transparency
    End If
    
    Call LoadFileMenu(mnuTools_ChatBot_Bot, App.Path & "\Bots\", "*.gcb")
    Call LoadFileMenu(mnuTools_GMScript_Script, App.Path & "\Scripts\", "*.gms")
    
    i = Val(GetSettingX("App Settings", "MainWindow Left", (Screen.Width / 2) - (Me.Width / 2)))
    If Not i >= Screen.Width Then
        Me.Left = i
    End If
    i = Val(GetSettingX("App Settings", "MainWindow Top", (Screen.Height / 2) - (Me.Height / 2)))
    If Not i >= Screen.Height Then
        Me.Top = i
    End If
    
    If LCase$(Command$) = "/startup" And GetSettingX("App Settings", "Open MainWindow OnStart", True) = False Then
        Me.Visible = False
    End If
    
    WriteRegKey "HKEY_CURRENT_USER\Software\VB and VBA Program Settings\Gilly Messenger\Location", App.Path
    
    strDefaultBrowser = ReadRegKey("HKEY_CLASSES_ROOT\htmlfile\shell\open\command\")
    If Not strDefaultBrowser = vbNullString Then
        If Not InStr(strDefaultBrowser, """") = 0 Then
            strDefaultBrowser = Split(strDefaultBrowser, """")(1)
        End If
    End If
    
    strCustomBrowser = GetSettingX("App Settings", "Custom Browser")
    boolUseDefaultBrowser = GetSettingX("App Settings", "Use DefaultBrowser", True)
    
    strDefaultEmailApp = ReadRegKey("HKEY_CLASSES_ROOT\mailto\shell\open\command\")
    If Not strDefaultEmailApp = vbNullString Then
        If Not InStr(strDefaultEmailApp, """") = 0 Then
            strDefaultEmailApp = Split(strDefaultEmailApp, """")(1)
        End If
    Else
        strDefaultEmailApp = "msimn.exe"
    End If
        
    WindowLoaded = True
    
    Dim CmdParams() As String
    CmdParams = Split(Command$)
    If UBound(CmdParams) = 1 Then
        If IsEmail(CmdParams(0)) Then
            Call SignIn(CmdParams(0), CmdParams(1))
        End If
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    LastActive = Timer
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Me.Visible = False
        Cancel = 1
    Else
        Call TerminateGM
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.WindowState = vbMinimized Then
        Exit Sub
    End If
    
    If WindowLoaded Then
        MainWindowMax = (Me.WindowState = vbMaximized)
        If Not Me.WindowState = vbMaximized Then
            MainWindowWidth = Me.Width
            MainWindowHeight = Me.Height
        End If
    End If
    
    If objMSN_NS.State = NsState_SignedIn Then
        Me.Cls
        GradientFill Me.hDC, 0, 0, Me.ScaleWidth, 5, "D4DEF4", "D4DEF4", False
        Me.Line (0, 5)-(Me.ScaleWidth, 5), 14989991
        imgTopRight.Move Me.ScaleWidth - imgTopRight.Width
        GradientFill Me.hDC, 0, 6, Me.ScaleWidth, 40, "FFFFFF", "E2E9FB", True
        Me.Line (0, 40)-(Me.ScaleWidth, 40), 14989991
        imglstStatus.ListImages(StatusIcon(MyStatus)).Draw Me.hDC, 4, 14, imlTransparent
        lblNick.Width = Me.ScaleWidth - lblNick.Left - 8
        txtNick.Width = lblNick.Width
        lblNick.Caption = CropText(Me, lblNick.Width, objMSN_NS.Nick, " (" & StatusName(MyStatus) & ")")
        If lblNews.Visible Then
            tvwContacts.Move 4, tvwContacts.Top, Me.ScaleWidth - 8, Me.ScaleHeight - picStatus.Height - picNews.Height - tvwContacts.Top - 4
        Else
            tvwContacts.Move 4, tvwContacts.Top, Me.ScaleWidth - 8, Me.ScaleHeight - picStatus.Height - tvwContacts.Top - 4
        End If
    Else
        picMask.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - picStatus.Height
        picMask.Cls
        GradientFill picMask.hDC, 0, 0, picMask.ScaleWidth, picMask.ScaleHeight / 3, "FFFFFF", "C9D3F3", True
        GradientFill picMask.hDC, 0, picMask.ScaleHeight / 3, picMask.ScaleWidth, picMask.ScaleHeight, "C9D3F3", "FFFFFF", True
        If objMSN_NS.State = NsState_Disconnected Then
            picMask.Line (0, (picMask.ScaleHeight / 3) - 20)-(picMask.ScaleWidth, (picMask.ScaleHeight / 3) - 20), vbWhite
            GradientFill picMask.hDC, 0, (picMask.ScaleHeight / 3) - 19, Me.ScaleWidth, (picMask.ScaleHeight / 3) + 40, "E6EDFA", "E6EDFA", False
        End If
        picSignIn.Move (Me.ScaleWidth / 2) - (picSignIn.Width / 2), (Me.ScaleHeight / 3) - (picSignIn.Height / 2)
        picSignInProgress.Move (Me.ScaleWidth / 2) - (picSignInProgress.Width / 2), ((Me.ScaleHeight / 3) + 16) - (picSignInProgress.Height / 2)
        lblSigningIn.Move (Me.ScaleWidth / 2) - (TextWidth(lblSigningIn.Caption) / 2), (Me.ScaleHeight / 3) - (picSignIn.Height / 2)
    End If
    
    picStatus.Move 0, Me.ScaleHeight - picStatus.Height, Me.ScaleWidth
    picNews.Move 0, picStatus.Top - picNews.Height, Me.ScaleWidth
    picStatus.Cls
    picNews.Cls
    If lblNews.Visible Then
        picNews.Line (0, 0)-(picNews.ScaleWidth, 0), 14989991
        GradientFill picStatus.hDC, 0, 0, picStatus.ScaleWidth, 26, "FFFFFF", "C9D3F3", False
    Else
        picStatus.Line (0, 0)-(picStatus.ScaleWidth, 0), 14989991
        GradientFill picStatus.hDC, 0, 1, picStatus.ScaleWidth, 26, "FFFFFF", "C9D3F3", False
    End If
    picStatus.Line (0, 26)-(picStatus.ScaleWidth, 26), 14989991
    GradientFill picStatus.hDC, 0, 27, picStatus.ScaleWidth, picStatus.ScaleHeight, "D4DEF4", "D4DEF4", False
    lblStatus.Move 4, lblStatus.Top, Me.ScaleWidth - 8
End Sub

Private Sub imgStatus_Click()
    If mnuFile_MyStatus.Enabled Then
        PopupMenu mnuFile_MyStatus
    End If
End Sub

Private Sub lblEmail_Click()
    Call OpenMailBox
End Sub

Private Sub lblNews_Click()
    If Not lblNews.Tag = vbNullString Then
        If PathIsURL(lblNews.Tag) Then
            Call WebNavigate(lblNews.Tag)
        Else
            ShellExecute 0, "open", lblNews.Tag, vbNullString, vbNullString, 1
        End If
    End If
End Sub

Private Sub lblNick_Click()
    txtNick.Text = objMSN_NS.Nick
    txtNick.SelStart = 0
    txtNick.SelLength = Len(txtNick.Text)
    txtNick.Visible = True
    txtNick.SetFocus
End Sub

Private Sub lblStatus_Change()
    lblStatus.ToolTipText = lblStatus.Caption
End Sub

Private Sub lblStatus_DblClick()
    ShellExecute 0, "open", StatusHistoryFolder & "\" & objMSN_NS.Login & ".txt", vbNullString, vbNullString, 1
End Sub

Private Sub lblStatus_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        lblStatus.Caption = vbNullString
    End If
End Sub

Private Sub mnuActions_SendAFileOrPhoto_Click()
    Call mnuFile_SendAFileOrPhoto_Click
End Sub

Private Sub mnuActions_SendAnInstantMessage_Click()
    Dim strEmail As String
    strEmail = InputBox("Enter the email address of the person you want to message.", "Send an Instant Message...")
    If Not strEmail = vbNullString Then
        StartChat strEmail, , , True
    End If
End Sub

Private Sub mnuActions_SendEmail_Click()
    Dim strEmail As String
    strEmail = InputBox("Enter the email address of the person you want to send an email.", "Send E-mail...")
    If Not strEmail = vbNullString Then
        Call SendEmail(strEmail)
    End If
End Sub

Private Sub mnuContact_Block_Click()
    If mnuContact_Block.Caption = "&Block" Then
        BlockContact mnuContact.Tag
    Else
        UnblockContact mnuContact.Tag
    End If
End Sub

Private Sub mnuContact_CopyContactTo_Group_Click(Index As Integer)
    objMSN_NS.AddContact msnList_Forward, mnuContact.Tag, mnuContact.Tag, Val(mnuContact_CopyContactTo_Group(Index).Tag)
End Sub

Private Sub mnuContact_CopyEmail_Click()
    Clipboard.Clear
    Clipboard.SetText mnuContact.Tag
End Sub

Private Sub mnuContact_CopyNick_Click()
    Clipboard.Clear
    Clipboard.SetText GetContactAttr(mnuContact.Tag, "nick")
End Sub

Private Sub mnuContact_DeleteContact_Click()
    If MsgBox("Are you user you want to delete " & mnuContact.Tag & " from your contact list?", vbQuestion Or vbYesNo) = vbYes Then
        objMSN_NS.RemoveContact msnList_Forward, mnuContact.Tag
    End If
End Sub

Private Sub mnuContact_Hide_Click()
    Call HideContact(mnuContact.Tag)
End Sub

Private Sub mnuContact_Ignore_Click()
    If mnuContact_Ignore.Caption = "&Ignore" Then
        IgnoreContact mnuContact.Tag
    Else
        UnignoreContact mnuContact.Tag
    End If
End Sub

Private Sub mnuContact_MoveContactTo_Group_Click(Index As Integer)
    objMSN_NS.AddContact msnList_Forward, mnuContact.Tag, mnuContact.Tag, Val(mnuContact_MoveContactTo_Group(Index).Tag)
    objMSN_NS.RemoveContact msnList_Forward, mnuContact.Tag, Val(mnuContact_MoveContactTo.Tag)
End Sub

Private Sub mnuContact_OpenMessageHistory_Click()
    ShellExecute 0, "open", MessageHistoryFolder & "\" & mnuContact.Tag & ".txt", vbNullString, vbNullString, 1
End Sub

Private Sub mnuContact_PopupFilter_Click()
    On Error Resume Next
    
    If mnuContact_PopupFilter.Caption = "Add to Pop&up Filter" Then
        If PopupFilter.Count = 1 Then
            If PopupFilter(1) = "*@*.*" Then
                PopupFilterMode = Not PopupFilterMode
                SaveSettingX "Popup Filter\" & objMSN_NS.Login, "Mode", PopupFilterMode
                PopupFilter.Remove "*@*.*"
            End If
        End If
        PopupFilter.Add mnuContact.Tag, mnuContact.Tag
        SaveSettingX "Popup Filter\" & objMSN_NS.Login, mnuContact.Tag, mnuContact.Tag
    Else
        If PopupFilter.Count = 1 Then
            If PopupFilter(1) = mnuContact.Tag Then
                PopupFilterMode = Not PopupFilterMode
                SaveSettingX "Popup Filter\" & objMSN_NS.Login, "Mode", PopupFilterMode
                PopupFilter.Add "*@*.*", "*@*.*"
            End If
        End If
        PopupFilter.Remove mnuContact.Tag
        DeleteSetting "Gilly Messenger", "Popup Filter\" & objMSN_NS.Login, mnuContact.Tag
    End If
End Sub

Private Sub mnuContact_Properties_Click()
    ShowBuddyProperties Me, mnuContact.Tag
End Sub

Private Sub mnuContact_RemoveContactFromGroup_Click()
    objMSN_NS.RemoveContact msnList_Forward, mnuContact.Tag, Val(mnuContact_RemoveContactFromGroup.Tag)
End Sub

Private Sub mnuContact_SendAFileOrPhoto_Click()
    If Not GetUserFile("All Files|*.*", "Send a File to " & mnuContact.Tag) = vbNullString Then
        StartChat mnuContact.Tag, , CommonDialog.FileTitle & "|" & CommonDialog.FileName, True
    End If
End Sub

Private Sub mnuContact_SendAnInstantMessage_Click()
    StartChat mnuContact.Tag, , , True
End Sub

Private Sub mnuContact_SendEmail_Click()
    Call SendEmail(mnuContact.Tag)
End Sub

Private Sub mnuContact_SoundFilter_Click()
    On Error Resume Next
    
    If mnuContact_SoundFilter.Caption = "Add to S&ound Filter" Then
        If SoundFilter.Count = 1 Then
            If SoundFilter(1) = "*@*.*" Then
                SoundFilterMode = Not SoundFilterMode
                SaveSettingX "Sound Filter\" & objMSN_NS.Login, "Mode", SoundFilterMode
                SoundFilter.Remove "*@*.*"
            End If
        End If
        SoundFilter.Add mnuContact.Tag, mnuContact.Tag
        SaveSettingX "Sound Filter\" & objMSN_NS.Login, mnuContact.Tag, mnuContact.Tag
    Else
        If SoundFilter.Count = 1 Then
            If SoundFilter(1) = mnuContact.Tag Then
                SoundFilterMode = Not SoundFilterMode
                SaveSettingX "Sound Filter\" & objMSN_NS.Login, "Mode", SoundFilterMode
                SoundFilter.Add "*@*.*", "*@*.*"
            End If
        End If
        SoundFilter.Remove mnuContact.Tag
        DeleteSetting "Gilly Messenger", "Sound Filter\" & objMSN_NS.Login, mnuContact.Tag
    End If
End Sub

Private Sub mnuContact_ViewProfile_Click()
    Call WebNavigate("http://members.msn.com/" & mnuContact.Tag)
End Sub

Public Sub mnuContacts_AddAContact_Click()
    Dim strContact As String
    strContact = InputBox("Enter the email of the person you want to add.", "Add a Contact")
    If Not strContact = vbNullString Then
        If InStr(strContact, "@") = 0 Then
            strContact = strContact & "@hotmail.com"
        End If
        AddContact strContact
    End If
End Sub

Private Sub mnuContacts_ImportContactsFromASavedFile_Click()
    On Error GoTo Handler
    If Not GetUserFile("GM Contact List (*.gcl)|*.gcl", "Import Messenger Contact List") = vbNullString Then
        Dim FileNum As Integer
        FileNum = FreeFile
        Open CommonDialog.FileName For Input As #FileNum
        Dim strContact As String
        Do Until EOF(FileNum) = True
            Input #FileNum, strContact
            If InStr(strContact, "@") > 0 And InStr(strContact, ".") > 0 Then
                Call AddContact(strContact)
            End If
        Loop
        Close #FileNum
        MsgBox "Your contact list was successfully imported.", vbInformation
    End If
    Exit Sub
Handler:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub mnuContacts_ManageContacts_ViewContactsBy_DisplayName_Click()
    If ViewContactsByEmail Then
        ViewContactsByEmail = False
        mnuContacts_ManageContacts_ViewContactsBy_EmailAddress.Checked = False
        mnuContacts_ManageContacts_ViewContactsBy_DisplayName.Checked = True
        Call RefreshTreeview
        SaveSettingX "App Settings\" & objMSN_NS.Login, "ViewContactsByEmail", False
    End If
End Sub

Private Sub mnuContacts_ManageContacts_ViewContactsBy_EmailAddress_Click()
    If Not ViewContactsByEmail Then
        ViewContactsByEmail = True
        mnuContacts_ManageContacts_ViewContactsBy_EmailAddress.Checked = True
        mnuContacts_ManageContacts_ViewContactsBy_DisplayName.Checked = False
        Call RefreshTreeview
        SaveSettingX "App Settings\" & objMSN_NS.Login, "ViewContactsByEmail", True
    End If
End Sub

Private Sub mnuContacts_ManageGroups_CreateNewGroup_Click()
    Call mnuGroup_CreateNewGroup_Click
End Sub

Private Sub mnuContacts_ManageGroups_DeleteAGroup_Group_Click(Index As Integer)
    objMSN_NS.RemoveGroup Val(Split(mnuContacts_ManageGroups_DeleteAGroup_Group(Index).Tag)(1))
End Sub

Private Sub mnuContacts_ManageGroups_GroupOfflineContactsTogether_Click()
    GroupOfflineContactsTogether = Not GroupOfflineContactsTogether
    mnuContacts_ManageGroups_GroupOfflineContactsTogether.Checked = GroupOfflineContactsTogether
    SaveSettingX "App Settings\" & objMSN_NS.Login, "GroupOfflineContactsTogether", GroupOfflineContactsTogether
    Call RefreshTreeview
End Sub

Private Sub mnuContacts_ManageGroups_RenameAGroup_Group_Click(Index As Integer)
    RenameGroup Val(mnuContacts_ManageGroups_RenameAGroup_Group(Index).Tag)
End Sub

Private Sub mnuContacts_SaveContactList_Click()
    On Error GoTo Handler
    If Not GetUserFile("GM Contact List (*.gcl)|*.gcl", "Save Messenger Contact List", 1) = vbNullString Then
        Dim FileNum As Integer
        FileNum = FreeFile
        Open CommonDialog.FileName For Output As #FileNum
        Dim i As Integer
        For i = 1 To ContactList.Count
            If InList(ContactList(i).Item("lists"), msnList_Forward) Then
                Print #FileNum, ContactList(i).Item("email")
            End If
        Next
        Close #FileNum
        CommonDialog.FileName = vbNullString
        CommonDialog.DialogTitle = vbNullString
        MsgBox "Your contact list was successfully saved.", vbInformation
    End If
    Exit Sub
Handler:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub mnuContacts_SearchForAContact_AdvancedSearch_Click()
    Call WebNavigate("http://members.msn.com/find.msnw")
End Sub

Private Sub mnuContacts_SearchForAContact_ContactList_Click()
    mnuContacts_SearchForAContact.Tag = vbNullString
    Call SearchForAContact
End Sub

Private Sub mnuContacts_SearchForAContact_SearchByInterest_Click()
    Call WebNavigate("http://members.msn.com/rootcat.msnw")
End Sub

Private Sub mnuContacts_SortContactsBy_Groups_Click()
    If Not SortContactsByGroups Then
        SortContactsByGroups = True
        SaveSettingX "App Settings\" & objMSN_NS.Login, "SortContactsByGroups", True
        Call RefreshTreeview
    End If
    mnuContacts_SortContactsBy_Groups.Checked = True
    mnuContacts_SortContactsBy_OnlineOffline.Checked = False
    mnuContacts_ManageGroups_CreateNewGroup.Enabled = True
    mnuContacts_ManageGroups_DeleteAGroup.Enabled = True
    mnuContacts_ManageGroups_RenameAGroup.Enabled = True
    mnuContacts_ManageGroups_GroupOfflineContactsTogether.Enabled = True
End Sub

Private Sub mnuContacts_SortContactsBy_OnlineOffline_Click()
    If SortContactsByGroups Then
        SortContactsByGroups = False
        SaveSettingX "App Settings\" & objMSN_NS.Login, "SortContactsByGroups", False
        Call RefreshTreeview
    End If
    mnuContacts_SortContactsBy_Groups.Checked = False
    mnuContacts_SortContactsBy_OnlineOffline.Checked = True
    mnuContacts_ManageGroups_CreateNewGroup.Enabled = False
    mnuContacts_ManageGroups_DeleteAGroup.Enabled = False
    mnuContacts_ManageGroups_RenameAGroup.Enabled = False
    mnuContacts_ManageGroups_GroupOfflineContactsTogether.Enabled = False
End Sub

Private Sub mnuFile_Close_Click()
    If mnuFile_Close.Caption = "&Close" Then
        mnuFile_Close.Caption = "&Open Gilly Messenger"
        Me.Visible = False
    Else
        mnuFile_Close.Caption = "&Close"
        ActivateWindow Me
    End If
End Sub

Private Sub mnuFile_Exit_Click()
    Call TerminateGM
End Sub

Private Sub mnuFile_Goto_Chatrooms_Click()
    objMSN_NS.RequestURL "CHAT 0x0409"
End Sub

Private Sub mnuFile_Goto_MsnHome_Click()
    Call WebNavigate("http://www.msn.com")
End Sub

Private Sub mnuFile_Goto_MsnToday_Click()
    Call WebNavigate("http://msntoday.msn.com")
End Sub

Private Sub mnuFile_Goto_MyEmailInbox_Click()
    Call lblEmail_Click
End Sub

Private Sub mnuFile_Goto_MyPassport_Click()
    objMSN_NS.RequestURL "PERSON 0x0409"
End Sub

Private Sub mnuFile_Goto_MyProfile_Click()
    objMSN_NS.RequestURL "PROFILE 0x1409"
End Sub

Private Sub mnuFile_MyStatus_Status_Click(Index As Integer)
    If Not Index = MyStatus Then
        objMSN_NS.ChangeStatus Index
    End If
End Sub

Private Sub mnuFile_OpenMessageHistory_Click()
    ShellExecute 0, "open", MessageHistoryFolder, vbNullString, vbNullString, 1
End Sub

Private Sub mnuFile_OpenReceivedFiles_Click()
    ShellExecute 0, "open", ReceivedFilesFolder, vbNullString, vbNullString, 1
End Sub

Private Sub mnuFile_SendAFileOrPhoto_Click()
    Dim strContact As String
    strContact = InputBox("Enter the email of the person you want to send a file.", "Send a File or Photo")
    If Not strContact = vbNullString Then
        If Not GetUserFile("All Files|*.*", "Send a File to " & strContact) = vbNullString Then
            StartChat strContact, , CommonDialog.FileTitle & "|" & CommonDialog.FileName, True
        End If
    End If
End Sub

Private Sub mnuFile_SignIn_Click()
    On Error Resume Next
    
    If mnuFile_SignIn.Caption = "S&ign In..." Then
        frmSignIn.Show vbModal, Me
    ElseIf mnuFile_SignIn.Caption = "Cancel S&ign In" Then
        lblStatus.Caption = vbNullString
        objMSN_NS.Disconnect
        frmMain.tmrReconnect.Tag = 0
        frmMain.tmrReconnect.Enabled = False
    End If
End Sub

Private Sub mnuFile_SignOut_Click()
    Call Signout
End Sub

Private Sub mnuGroup_CreateNewGroup_Click()
    If Not GroupExists("New Group") Then
        objMSN_NS.AddGroup "New Group"
    Else
        Dim i As Integer
        i = 1
        Do
            If Not GroupExists("New Group " & i) Then
                objMSN_NS.AddGroup "New Group " & i
                Exit Do
            End If
            i = i + 1
        Loop
    End If
End Sub

Private Sub mnuGroup_DeleteGroup_Click()
    If MsgBox("Are you sure you want to delete this group?", vbYesNo) = vbYes Then
        objMSN_NS.RemoveGroup Val(mnuGroup.Tag)
    End If
End Sub

Private Sub mnuGroup_RenameGroup_Click()
    RenameGroup Val(mnuGroup.Tag)
End Sub

Private Sub mnuGroup_SaveGroupToAFile_Click()
    On Error GoTo Handler
    CommonDialog.DialogTitle = "Save Messenger Contact List"
    CommonDialog.Filter = "GM Contact List (*.gcl)|*.gcl"
    CommonDialog.ShowSave
    If Not CommonDialog.FileName = vbNullString Then
        Dim FileNum As Integer
        FileNum = FreeFile
        Open CommonDialog.FileName For Output As #FileNum
        Dim i As Integer
        For i = 1 To ContactList.Count
            If InCollection(ContactList(i).Item("groups"), "GRP " & mnuGroup.Tag) Then
                Print #FileNum, ContactList(i).Item("email")
            End If
        Next
        Close #FileNum
        CommonDialog.FileName = vbNullString
        CommonDialog.DialogTitle = vbNullString
        MsgBox "Your contact list was successfully saved.", vbInformation
    End If
    Exit Sub
Handler:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub mnuHelp_AboutGillyMessenger_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuHelp_CrackSoftForums_Click()
    Call WebNavigate("http://www.cracksoft.net/forums")
End Sub

Private Sub mnuHelp_CrackSoftWebsite_Click()
    Call WebNavigate("http://www.cracksoft.net")
End Sub

Private Sub mnuHelp_Readme_Click()
    ShellExecute 0, "open", App.Path & "\readme.htm", vbNullString, vbNullString, 1
End Sub

Private Sub mnuTools_AutoMessage_Click()
    If mnuTools_AutoMessage.Checked Then
        mnuTools_AutoMessage.Checked = False
    Else
        Dim strMessage As String
        strMessage = InputBox("Enter the message", "AutoMessage", mnuTools_AutoMessage.Tag)
        If Not strMessage = vbNullString Then
            mnuTools_AutoMessage.Tag = strMessage
            mnuTools_AutoMessage.Checked = True
        End If
    End If
End Sub

Private Sub mnuTools_ChangeDisplayPic_Click()
    frmDisplayPic.Show vbModal, Me
End Sub

Private Sub mnuTools_ChatBot_Bot_Click(Index As Integer)
    If mnuTools_ChatBot_Bot(Index).Checked Then
        Call LoadBot(vbNullString)
    Else
        Call LoadBot(mnuTools_ChatBot_Bot(Index).Tag)
    End If
End Sub

Private Sub mnuTools_ChatBot_Other_Click()
    If mnuTools_ChatBot_Other.Checked Then
        Call LoadBot(vbNullString)
    Else
        If Not GetUserFile("GM Chat Bots (*.gcb)|*.gcb") = vbNullString Then
            Call LoadBot(CommonDialog.FileName)
        End If
    End If
End Sub

Private Sub mnuTools_GMScript_Other_Click(Index As Integer)
    If Index = 0 Then
        If Not GetUserFile("GM Scripts (*.gms)|*.gms") = vbNullString Then
            Load mnuTools_GMScript_Other(mnuTools_GMScript_Other.UBound + 1)
            mnuTools_GMScript_Other(mnuTools_GMScript_Other.UBound).Tag = CommonDialog.FileName
            If InStr(CommonDialog.FileTitle, ".") > 0 Then
                mnuTools_GMScript_Other(mnuTools_GMScript_Other.UBound).Caption = Left$(CommonDialog.FileTitle, InStr(CommonDialog.FileTitle, ".") - 1)
            Else
                mnuTools_GMScript_Other(mnuTools_GMScript_Other.UBound).Caption = CommonDialog.FileTitle
            End If
            mnuTools_GMScript_Other(mnuTools_GMScript_Other.UBound).Checked = True
            Call LoadScript(CommonDialog.FileName)
        End If
    Else
        Call EndScript(GMScripts(mnuTools_GMScript_Other(Index).Tag))
    End If
End Sub

Private Sub mnuTools_GMScript_Script_Click(Index As Integer)
    mnuTools_GMScript_Script(Index).Checked = Not mnuTools_GMScript_Script(Index).Checked
    If mnuTools_GMScript_Script(Index).Checked Then
        Call LoadScript(mnuTools_GMScript_Script(Index).Tag)
    Else
        Call EndScript(GMScripts(mnuTools_GMScript_Script(Index).Tag))
    End If
End Sub

Private Sub mnuTools_IgnoreAll_Click()
    Dim i As Integer
    For i = 1 To ContactList.Count
        If InList(ContactList(i).Item("lists"), msnList_Forward) Then
            IgnoreContact ContactList(i).Item("email")
        End If
    Next
End Sub

Private Sub mnuTools_MessageAll_Click()
    Dim strMessage As String
    strMessage = InputBox("Enter the message for contacts in all conversations.", "Message All")
    If Not strMessage = vbNullString Then
        Call MessageAll(strMessage)
    End If
End Sub

Private Sub mnuTools_Options_Click()
    On Error Resume Next
    
    frmOptions.Show vbModal, Me
End Sub

Private Sub mnuTools_RemoteControl_Click()
    On Error Resume Next
    
    frmRemoteControl.Show vbModal, Me
End Sub

Private Sub objMSN_NS_ChatRequest(Email As String, Nick As String, SessionID As Double, AuthCode As String, ServerIP As String, ServerPort As Integer)
    On Error Resume Next
    
    If GetContactAttr(Email, "status") = msnStatus_Offline And InList(GetContactAttr(Email, "lists"), msnList_Forward) And BlockAlert Then
        LogStatus Nick & " [" & Email & "] has blocked you."
        If SoundAlerts And boolAlertSound And Not FullScrApp Then
            Call ContactSound(Email, strAlertSound, "alert")
        End If
        If Not LastBlockAlert = Email Then
            MsgBox Nick & " [" & Email & "] has blocked you.", vbInformation
            LastBlockAlert = Email
        End If
    Else
        If Not InCollection(IgnoreList, Email) Then
            Dim IMWindow As New frmChat
            
            Load IMWindow
            With IMWindow
                .WindowState = vbMinimized
                .objMSN_SB.Socket = Controls.Add("MSWinsock.Winsock", "wskSB" & Fix(Timer) & frmMain.Controls.Count)
                .objMSN_SB.Contact = Email
                .objMSN_SB.SessionID = SessionID
                .objMSN_SB.AuthCode = AuthCode
                .objMSN_SB.Server = ServerIP
                .objMSN_SB.Port = ServerPort
                .objMSN_SB.Login = objMSN_NS.Login
                .objMSN_SB.SessionType = SbSession_Ring
                .objMSN_SB.Connect
                
                .Caption = .BuddyNick & " - Conversation"
                .lblStatus.Caption = "[" & Time$ & "] " & Email & " has opened your window."
                Call LogChat(Email, "----" & vbCrLf & "[" & Now & "] " & Email & " has opened your window." & vbCrLf & "----")
            
                .lblBuddies.Caption = Email
                
                If Not GetSettingX("App Settings\" & frmMain.objMSN_NS.Login & "\Show DP", Email, True) Then
                    .imgBuddyDP.Width = 1
                    .imgBuddyDP.Height = .imgShowHideBuddyDP.Height
                    .imgBuddyDP.BorderStyle = vbBSNone
                    Set .imgBuddyDP.Picture = LoadPicture(vbNullString)
                    Call .Form_Resize
                End If
            End With
            
            Call QueScript(IMWindow, "imwindowopened", ConvArray(Email, Nick))
        End If
    End If
End Sub

Private Sub objMSN_NS_NickChanged()
    lblNick.Caption = CropText(Me, lblNick.Width, objMSN_NS.Nick, " (" & StatusName(MyStatus) & ")")
    Call LogStatus("Nick changed to " & objMSN_NS.Nick)
    Call QueScript(Me, "nickchanged", ConvArray(objMSN_NS.Nick))
End Sub

Private Sub objMSN_NS_SignInProgress(Percent As Integer)
    picSignInProgress.Cls
    GradientFill picSignInProgress.hDC, 0, 0, (picSignInProgress.ScaleWidth / 100) * Percent, picSignInProgress.ScaleHeight, "FFFFFF", "C9D3F3", False
    Dim r As RECT
    r.Left = 0
    r.Top = 0
    r.Right = (picSignInProgress.ScaleWidth / 100) * Percent
    r.Bottom = picSignInProgress.ScaleHeight
    DrawEdge picSignInProgress.hDC, r, BDR_RAISEDINNER, BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM
End Sub

Private Sub objMSN_NS_StateChanged()
    On Error Resume Next
    
    Select Case objMSN_NS.State
    Case NsState_Connecting
        Call ChangeGMStatus("Connecting...")
        Call Form_Resize
        
    Case NsState_Negotiating
        Call ChangeGMStatus("Negotiating...")
        
    Case NsState_Authenticating
        Call ChangeGMStatus("Authenticating...")
        
    Case NsState_SwitchingServer
        Call ChangeGMStatus("Switching Server...")
        SaveSettingX "Server Settings", "IPAddress", objMSN_NS.Server
        SaveSettingX "Server Settings", "Port", objMSN_NS.Port
        
    Case NsState_SignedIn
        If Val(tmrReconnect.Tag) = 0 Then
            MyStatus = InitialStatus
        Else
            MyStatus = LastStatus
        End If
        
        Call UpdateTrayIcon
        
        If SavePassword Then
            SaveSettingX "Login Cache", objMSN_NS.Login, InitialStatus & XorEncrypt(objMSN_NS.Password, objMSN_NS.Login)
        Else
            SaveSettingX "Login Cache", objMSN_NS.Login, InitialStatus
        End If
        
        Call LoadUserSettings
        
        lblEmail.Caption = "Goto my e-mail inbox"
        
        mnuFile_SignIn.Caption = "S&ign In..."
        mnuFile_SignIn.Enabled = False
        mnuFile_SignOut.Enabled = True
        mnuFile_MyStatus_Status(InitialStatus).Checked = True
        mnuFile_Goto_MyEmailInbox.Enabled = True
        mnuFile_Goto_MyProfile.Enabled = True
        mnuFile_Goto_MyPassport.Enabled = True
        mnuFile_Goto_Chatrooms.Enabled = True
        mnuFile_Goto_MsnToday.Enabled = True
        mnuFile_SendAFileOrPhoto.Enabled = True
        mnuFile_OpenMessageHistory = True
        mnuContacts_AddAContact.Enabled = True
        mnuContacts_SearchForAContact.Enabled = True
        mnuContacts_ManageContacts.Enabled = True
        mnuContacts_ManageGroups.Enabled = True
        mnuContacts_SortContactsBy.Enabled = True
        mnuContacts_SaveContactList.Enabled = True
        mnuContacts_ImportContactsFromASavedFile.Enabled = True
        mnuActions_SendAnInstantMessage.Enabled = True
        mnuActions_SendAFileOrPhoto.Enabled = True
        mnuActions_SendEmail.Enabled = True
        
        Call Form_Resize
        
        picMask.Visible = False
        Call ChangeGMStatus("Retreiving Contact List...")
        lblNick.Caption = CropText(Me, lblNick.Width, objMSN_NS.Nick, " (" & StatusName(MyStatus) & ")")
        
        Set ContactList = New Collection
        Set ContactGroups = New Collection
        Set ContactProperties = New Collection
        Set UserProperties = New Collection
        
        Set IMWindows = New Collection
        Set PendingIM = New Collection
        
        Set ContactComments = LoadSettingsInCollection("Comments\" & objMSN_NS.Login)
        Set ContactCustomNicks = LoadSettingsInCollection("Custom Nicks\" & objMSN_NS.Login)
        Set HiddenContacts = LoadSettingsInCollection("Hide List\" & objMSN_NS.Login)
        Set IgnoreList = LoadSettingsInCollection("Ignore List\" & objMSN_NS.Login)
        
        SoundFilterMode = GetSettingX("Sound Filter\" & objMSN_NS.Login, "Mode", False)
        Set SoundFilter = LoadSettingsInCollection("Sound Filter\" & objMSN_NS.Login, "Mode")
        If SoundFilter.Count = 0 Then
            SoundFilter.Add "*@*.*", "*@*.*"
        End If
        PopupFilterMode = GetSettingX("Popup Filter\" & objMSN_NS.Login, "Mode", False)
        Set PopupFilter = LoadSettingsInCollection("Popup Filter\" & objMSN_NS.Login, "Mode")
        If PopupFilter.Count = 0 Then
            PopupFilter.Add "*@*.*", "*@*.*"
        End If
    
        SortContactsByGroups = GetSettingX("App Settings\" & objMSN_NS.Login, "SortContactsByGroups", True)
        If SortContactsByGroups Then
            Call mnuContacts_SortContactsBy_Groups_Click
        Else
            Call mnuContacts_SortContactsBy_OnlineOffline_Click
        End If
        
        GroupOfflineContactsTogether = GetSettingX("App Settings\" & objMSN_NS.Login, "GroupOfflineContactsTogether", False)
        mnuContacts_ManageGroups_GroupOfflineContactsTogether.Checked = GroupOfflineContactsTogether
                
        ViewContactsByEmail = GetSettingX("App Settings\" & objMSN_NS.Login, "ViewContactsByEmail", False)
        mnuContacts_ManageContacts_ViewContactsBy_EmailAddress.Checked = ViewContactsByEmail
        mnuContacts_ManageContacts_ViewContactsBy_DisplayName.Checked = Not ViewContactsByEmail
        
        If Not SortContactsByGroups Then
            AddTVParentNode "online", "Online", "None of your contacts are online"
            AddTVParentNode "offline", "Not Online", "None of your contacts are offline"
        End If
        
        Call objMSN_NS.RequestContactList(0)
        
        NewsData = vbNullString
        NewsInterval = 0
        wskNews.Close
        wskNews.Connect "www.cracksoft.net", 80
        
        If Not ArraySize(NewsLines) = -1 Then
            Call initNews
        End If
        
    Case NsState_Disconnected
        tmrNewsScroller1.Enabled = False
        tmrNewsScroller2.Enabled = False
        picNews.Visible = False
        lblNews.Visible = False
        NewsData = vbNullString
        
        tmrAutoIdle.Enabled = False
        
        tmrPing.Enabled = False
        
        Call UpdateTrayIcon
  
        InboxUnread = 0
        FoldersUnread = 0
    
        Set ContactList = Nothing
        Set ContactGroups = Nothing
        Set ContactProperties = Nothing
        Set UserProperties = Nothing
        
        Dim frmTemp As Form
        For Each frmTemp In Forms
            If frmTemp.Name = "frmChat" Then
                If frmTemp.Visible = False And frmTemp.objMSN_SB.State = SbState_Disconnected Then
                    Unload frmTemp
                    Set frmTemp = Nothing
                End If
            ElseIf frmTemp.Name = "frmAddContact" Then
                Unload frmTemp
                Set frmTemp = Nothing
            End If
        Next
        
        Set IMWindows = Nothing
        Set PendingIM = Nothing
        
        Set ContactComments = Nothing
        Set HiddenContacts = Nothing
        Set IgnoreList = Nothing
        
        Set SoundFilter = Nothing
        Set PopupFilter = Nothing
    
        tvwContacts.Nodes.Clear
        lblNick.Caption = vbNullString
        mnuFile_SignIn.Caption = "S&ign In..."
    
        mnuFile_SignIn.Enabled = True
        mnuFile_SignOut.Enabled = False
        mnuFile_MyStatus_Status(MyStatus).Checked = False
        mnuFile_MyStatus.Enabled = False
        mnuFile_Goto_MyEmailInbox.Enabled = False
        mnuFile_Goto_MyProfile.Enabled = False
        mnuFile_Goto_MyPassport.Enabled = False
        mnuFile_Goto_Chatrooms.Enabled = False
        mnuFile_Goto_MsnToday.Enabled = False
        mnuFile_SendAFileOrPhoto.Enabled = False
        mnuFile_OpenMessageHistory.Enabled = False
        mnuContacts_AddAContact.Enabled = False
        mnuContacts_SearchForAContact.Enabled = False
        mnuContacts_ManageContacts.Enabled = False
        mnuContacts_ManageGroups.Enabled = False
        mnuContacts_SortContactsBy.Enabled = False
        mnuContacts_SaveContactList.Enabled = False
        mnuContacts_ImportContactsFromASavedFile.Enabled = False
        mnuActions_SendAnInstantMessage.Enabled = False
        mnuActions_SendAFileOrPhoto.Enabled = False
        mnuActions_SendEmail.Enabled = False
        mnuTools_ChangeDisplayPic.Enabled = False
        mnuTools_MessageAll.Enabled = False
        mnuTools_IgnoreAll.Enabled = False
        mnuTools_GMScript.Enabled = False
        
        ClearSubMenu mnuContacts_ManageGroups_DeleteAGroup_Group
        ClearSubMenu mnuContacts_ManageGroups_RenameAGroup_Group
    
        Call Form_Resize
        picSignIn.Visible = True
        picSignInProgress.Visible = False
        picMask.Visible = True
    
        If lblStatus.Tag = vbNullString Then
            lblStatus.Caption = vbNullString
        Else
            lblStatus.Tag = vbNullString
        End If
        
        Set GMScripts = Nothing
        Set GMScripts = New Collection
        Dim i As Integer
        For i = 1 To mnuTools_GMScript_Script.UBound
            mnuTools_GMScript_Script(i).Checked = False
        Next
        For i = 1 To mnuTools_GMScript_Other.UBound
            Unload mnuTools_GMScript_Other(i)
        Next
    End Select
End Sub

Private Sub objMSN_NS_ContactAdded(Email As String, Nick As String, list As Integer, GroupID As Integer)
    On Error Resume Next
    
    If list = msnList_Reverse Then
        If objMSN_NS.GTC = "A" Then
            If Not LastAddAlert = Email Then
                If Not InList(GetContactAttr(Email, "lists"), msnList_Forward) Then
                    Dim AddContactForm As New frmAddContact
                    AddContactForm.lblEmailHasAddedYouToHisHerContactList.Caption = Nick & " [" & Email & "] has added you to his/her contact list."
                    AddContactForm.ContactEmail = Email
                    AddContactForm.ContactNick = Nick
                    Call SaveFocus
                    AddContactForm.Show
                    Call RestoreFocus
                Else
                    If SoundAlerts And boolAlertSound And Not FullScrApp Then
                        Call ContactSound(Email, strAlertSound, "alert")
                    End If
                    MsgBox Nick & " [" & Email & "]  has added you to his/her contact list.", vbInformation
                End If
                LastAddAlert = Email
            End If
            Call LogStatus(Nick & " [" & Email & "]  has added you to his/her contact list.")
        Else
            objMSN_NS.AddContact msnList_Allow, Email, Nick
        End If
    End If

    On Error Resume Next
    
    If InCollection(ContactList, Email) Then
        SetCollectionItem ContactList(Email), "lists", GetContactAttr(Email, "lists") Or list
        
        If GetContactAttr(Email, "status") = msnStatus_Unknown Then
            SetCollectionItem ContactList(Email), "status", msnStatus_Offline
        End If
        
        If Not GroupID = -1 Then
            ContactList(Email).Item("groups").Add GroupID, "GRP " & GroupID
        End If
        
        If list = msnList_Forward Or (InList(GetContactAttr(Email, "lists"), msnList_Forward) And (list = msnList_Block Or list = msnList_Reverse)) Then
            RefreshContact Email
        End If
    Else
        Dim Groups As Collection
        Set Groups = New Collection
        
        Groups.Add GroupID
        
        Dim NewContact As Collection
        Set NewContact = New Collection
        
        NewContact.Add Email, "email"
        NewContact.Add Nick, "nick"
        NewContact.Add list, "lists"
        NewContact.Add Groups, "groups"
        NewContact.Add IIf(InList(list, msnList_Forward), msnStatus_Offline, msnStatus_Unknown), "status"
        
        ContactList.Add NewContact, Email
        
        Set NewContact = Nothing
        Set Groups = Nothing
    End If
    
    If frmOptions.Visible Then
        If list = msnList_Allow Then
            Call frmOptions.LoadAllowList
        ElseIf list = msnList_Block Then
            Call frmOptions.LoadBlockList
        ElseIf list = msnList_Reverse Then
            If frmReverseList.Visible Then
                Call frmOptions.LoadReverseList
            End If
        End If
    End If
End Sub

Private Sub objMSN_NS_ContactOffline(Email As String)
    SetCollectionItem ContactList(Email), "status", msnStatus_Offline
    RefreshContact Email
    Call ChangeGMStatus(Email & " appears to be offline.", True)
    Call UpdateContactStatusInIMWindow(Email)
    If SoundAlerts And boolOfflineSound Then
        ContactSound Email, strOfflineSound, "contactstatus"
    End If
    Call QueScript(Me, "contactstatuschanged", ConvArray(Email, GetContactAttr(Email, "nick"), msnStatus_Offline))
End Sub

Private Sub objMSN_NS_ContactPropertyReceived(Email As String, Property As String, Value As String)
    On Error Resume Next
    
    If InCollection(ContactProperties, Email) Then
        SetCollectionItem ContactProperties(Email), Property, Value
    Else
        Dim NewProperty As Collection
        Set NewProperty = New Collection
        NewProperty.Add Value, Property
        ContactProperties.Add NewProperty, Email
        Set NewProperty = Nothing
    End If
End Sub

Private Sub objMSN_NS_ContactReceived(Email As String, Nick As String, Lists As Integer, Groups As Collection)
    On Error Resume Next
    
    Dim NewContact As Collection
    Set NewContact = New Collection
    
    NewContact.Add Email, "email"
    NewContact.Add Nick, "nick"
    NewContact.Add Lists, "lists"
    NewContact.Add Groups, "groups"
    NewContact.Add IIf(InList(Lists, msnList_Forward), msnStatus_Offline, msnStatus_Unknown), "status"
    
    ContactList.Add NewContact, Email
    Set NewContact = Nothing
    
    If InList(Lists, msnList_Forward) Then
        RefreshContact Email
    End If
    
    If Lists = 8 Then
        If objMSN_NS.GTC = "A" Then
            If Not LastAddAlert = Email Then
                Dim AddContactForm As New frmAddContact
                AddContactForm.lblEmailHasAddedYouToHisHerContactList.Caption = Nick & " [" & Email & "] has added you to his/her contact list."
                AddContactForm.ContactEmail = Email
                AddContactForm.ContactNick = Nick
                Call SaveFocus
                AddContactForm.Show
                Call RestoreFocus
                LastAddAlert = Email
            End If
        Else
            objMSN_NS.AddContact msnList_Allow, Email, Nick
        End If
    End If
End Sub

Private Sub objMSN_NS_ContactRemoved(Email As String, list As Integer, GroupID As Integer)
    On Error Resume Next
    
    If list = msnList_Forward Then
        'Remove treeview nodes
        Call RemoveContact(Email, GroupID)
        
        'Update contact list collection
        If GroupID = -1 Then
            SetCollectionItem ContactList(Email), "groups", New Collection
        Else
            ContactList(Email).Item("groups").Remove "GRP " & GroupID
        End If
        Call LogStatus(Email & " deleted from contact list.")
    End If
    
    If list <> msnList_Forward Or ContactList(Email).Item("groups").Count = 0 Then
        SetCollectionItem ContactList(Email), "lists", GetContactAttr(Email, "lists") And Not list
    End If
    
    If (InList(GetContactAttr(Email, "lists"), msnList_Forward) And (list = msnList_Block Or list = msnList_Reverse)) Then
        RefreshContact Email
    End If
    
    If frmOptions.Visible Then
        If list = msnList_Allow Then
            Call frmOptions.LoadAllowList
        ElseIf list = msnList_Block Then
            Call frmOptions.LoadBlockList
        ElseIf list = msnList_Reverse Then
            If frmReverseList.Visible Then
                Call frmOptions.LoadReverseList
            End If
        End If
    End If

    If list = msnList_Reverse Then
        LogStatus (Email & " has deleted you from his/her contact list.")
        If SoundAlerts And boolAlertSound And Not FullScrApp Then
            Call ContactSound(Email, strAlertSound, "alert")
        End If
        If Not LastDelAlert = Email Then
            MsgBox Email & " has deleted you from his/her contact list.", vbInformation
            LastDelAlert = Email
        End If
    End If
End Sub

Private Sub objMSN_NS_ContactRenamed(Email As String, Nick As String)
    On Error Resume Next
    
    SetCollectionItem ContactList(Email), "nick", Nick
    If InList(GetContactAttr(Email, "lists"), msnList_Forward) Then
        RefreshContact Email
    End If
    Call LogStatus(Email & " renamed from '" & GetContactAttr(Email, "nick") & "' to '" & Nick & "'")
End Sub

Private Sub objMSN_NS_ContactStatusChanged(Email As String, Nick As String, Status As Integer)
    On Error Resume Next
    
    Dim PrevStatus As Integer
    PrevStatus = GetContactAttr(Email, "status")
    
    Dim ScriptIndex As Integer
    
    If Not PrevStatus = msnStatus_Offline Then
        If Not GetContactAttr(Email, "nick") = Nick Then
            Call LogStatus(Email & " changed nick to '" & Nick & "'.")
            Call QueScript(Me, "contactnickchanged", ConvArray(Email, Nick))
        End If
        If Not PrevStatus = Status Then
            Call LogStatus(Email & " changed status to " & StatusName(Status) & ".")
        End If
    End If
    
    Call QueScript(Me, "contactstatuschanged", ConvArray(Email, Nick, Status))
    
    SetCollectionItem ContactList(Email), "nick", Nick
    SetCollectionItem ContactList(Email), "status", Status
    
    RefreshContact Email
    
    Call UpdateContactStatusInIMWindow(Email)
    
    If PrevStatus = msnStatus_Offline Then
        Call ChangeGMStatus(Email & " has signed in.")
        tmrTrayAnim.Enabled = True
        If AlertOnContactOnline And PatternSearch(PopupFilter, Email, PopupFilterMode) Then
            If ViewContactsByEmail Then
                ShowPopup Me, "CHAT " & Email, PopupMessage(Email, True) & vbCrLf & "has signed in."
            Else
                ShowPopup Me, "CHAT " & Email, PopupMessage(GetCustomNick(Email, Nick), True) & vbCrLf & "has signed in."
            End If
        End If
        Call LogStatus(Nick & " [" & Email & "]  has signed in.")
        
        If SoundAlerts And boolOnlineSound Then
            ContactSound Email, strOnlineSound, "contactstatus"
        End If
        SaveSettingX "Statistics\" & Email, "Last Online", Now()
    End If
    
    If InCollection(IMWindows, Email) Then
        If InCollection(IMWindows(Email).ChatBuddies, Email) Then
            SetCollectionItem IMWindows(Email).ChatBuddies(Email), "nick", Nick
            IMWindows(Email).UpdateBuddies
        Else
            IMWindows(Email).Caption = Nick & " - Conversation"
        End If
    ElseIf InCollection(PendingIM, Email) Then
        PendingIM(Email).Caption = Nick & " - Conversation"
    End If
End Sub

Private Sub objMSN_NS_EmailNotification(FromName As String, FromEmail As String, MessageURL As String, PostURL As String, Subject As String, DestFolder As String, ID As Integer)
    If DestFolder = "ACTIVE" Then
        InboxUnread = InboxUnread + 1
    Else
        FoldersUnread = FoldersUnread + 1
    End If
    Call UpdateMailStatus
    Call ChangeGMStatus("New mail in " & IIf(DestFolder = "ACTIVE", "Inbox", DestFolder) & " from " & FromEmail)
    tmrTrayAnim.Enabled = True
    If AlertOnEmailReceived Then
        ShowPopup Me, "URL " & MessageURL & " " & PostURL & " " & ID, "You have received a new e-mail message " & vbCrLf & "from " & FromName & vbCrLf & vbCrLf & PopupMessage(Subject, False)
    End If
    Call LogStatus("You have received a new e-mail message from " & FromName & "<" & FromEmail & "> in " & IIf(DestFolder = "ACTIVE", "Inbox", DestFolder) & " with subject '" & Subject & "'")
    If SoundAlerts And boolEmailSound Then
        ContactSound FromEmail, strEmailSound, "email"
    End If
End Sub

Private Sub objMSN_NS_GroupAdded(GroupID As Integer, GroupName As String)
    Call objMSN_NS_GroupReceived(GroupID, GroupName)
    RenameGroup GroupID
    Call RefreshTreeview
    Call LogStatus("Group '" & GroupName & "' added.")
End Sub

Private Sub objMSN_NS_GroupReceived(GroupID As Integer, GroupName As String)
    On Error Resume Next
    
    Dim NewGroup As Collection
    Set NewGroup = New Collection
    NewGroup.Add GroupID, "id"
    If GroupID = 0 And GroupName = "Personen" Then
        GroupName = "Other Contacts"
    End If
    NewGroup.Add GroupName, "name"
    
    ContactGroups.Add NewGroup, "GRP " & GroupID
    
    If SortContactsByGroups Then
        AddTVParentNode "GRP " & GroupID, GroupName, IIf(GroupOfflineContactsTogether, "Everyone in this group is offline", "You have no contacts in this group")
        If ContactGroups.Count > 1 Then
            tvwContacts.Sorted = True
            tvwContacts.Sorted = False
            tvwContacts.Nodes.Remove "GRP 0"
            AddTVParentNode "GRP 0", ContactGroups(1).Item("name"), IIf(GroupOfflineContactsTogether, "Everyone in this group is offline", "You have no contacts in this group")
        End If
    End If
    
    If Not GroupID = 0 Then
        AddSubMenu mnuContacts_ManageGroups_DeleteAGroup_Group, GroupName, CStr(GroupID)
    End If
    
    AddSubMenu mnuContacts_ManageGroups_RenameAGroup_Group, GroupName, CStr(GroupID)
    
    Set NewGroup = Nothing
End Sub

Private Sub objMSN_NS_GroupRemoved(GroupID As Integer)
    Call LogStatus("Group '" & ContactGroups("GRP " & GroupID).Item("name") & "' deleted.")
    ContactGroups.Remove "GRP " & GroupID
    If SortContactsByGroups Then
        tvwContacts.Nodes.Remove "GRP " & GroupID
    End If
    RemoveSubMenu mnuContacts_ManageGroups_DeleteAGroup_Group, CStr(GroupID)
    RemoveSubMenu mnuContacts_ManageGroups_RenameAGroup_Group, CStr(GroupID)
End Sub

Private Sub objMSN_NS_GroupRenamed(GroupID As Variant, GroupName As String)
    On Error Resume Next
    
    Call LogStatus("Group '" & ContactGroups("GRP " & GroupID).Item("name") & "' renamed to '" & GroupName & "'")
    
    SetCollectionItem ContactGroups("GRP " & GroupID), "name", GroupName
    
    RenameSubMenu mnuContacts_ManageGroups_DeleteAGroup_Group, CStr(GroupID), GroupName
    RenameSubMenu mnuContacts_ManageGroups_RenameAGroup_Group, CStr(GroupID), GroupName
    
    If SortContactsByGroups Then
        tvwContacts.Nodes("GRP " & GroupID).Text = GetParentNodeTitle("GRP " & GroupID)
        Call RefreshTreeview
    End If
End Sub

Private Sub objMSN_NS_ContactInitialStatus(Email As String, Nick As String, Status As Integer)
    SetCollectionItem ContactList(Email), "status", Status
    SetCollectionItem ContactList(Email), "nick", Nick
    RefreshContact Email
    SaveSettingX "Statistics\" & Email, "Last Online", Now()
End Sub

Private Sub objMSN_NS_ListRetrievalComplete()
    FirstStatus = True
    
    If Val(tmrReconnect.Tag) = 0 Then
        objMSN_NS.ChangeStatus InitialStatus
    Else
        objMSN_NS.ChangeStatus LastStatus
        tmrReconnect.Tag = "0"
    End If
    
    tmrPing.Enabled = True
End Sub

Private Sub objMSN_NS_MailBoxNotification(SrcFolder As String, DestFolder As String, Messages As Integer)
    If SrcFolder = "ACTIVE" Then
        InboxUnread = InboxUnread - Messages
        If InboxUnread < 0 Then
            InboxUnread = 0
        End If
    ElseIf Not SrcFolder = "trAsH" Then
        FoldersUnread = FoldersUnread - Messages
        If FoldersUnread < 0 Then
            FoldersUnread = 0
        End If
    End If
    Call UpdateMailStatus
End Sub

Private Sub objMSN_NS_MailBoxStatus(InboxUnreadMsgs As Integer, FoldersUnreadMsgs As Integer)
    InboxUnread = InboxUnreadMsgs
    FoldersUnread = FoldersUnreadMsgs
    Call UpdateMailStatus
End Sub

Private Sub objMSN_NS_NsError(Error As String)
    On Error Resume Next
    
    ActivateWindow Me
    
    Select Case Left$(Error, 3)
    Case "911", "917"
        objMSN_NS.Disconnect
        MsgBox "Invalid username or password.", vbInformation
        If SavePassword Then
            SaveSettingX "Login Cache", objMSN_NS.Login, vbNullString
        End If
        frmSignIn.Show vbModal, Me
    Case "OUT"
        Select Case Right$(Error, 3)
        Case "OTH"
            MsgBox "You have been signed out because you signed in from another location.", vbInformation
        Case "SSD"
            MsgBox "Server is going down for maintenance.", vbInformation
        End Select
    Case "201", "205", "208"
        NS_Alert "Contact does not exist.", vbExclamation, "Server Warning"
    Case "206"
        MsgBox "Domain name missing.", vbExclamation, "Server Error"
    Case "209"
        NS_Alert "Illegal nickname.", vbExclamation, "Server Warning"
    Case "210"
        MsgBox "Contact list is full.", vbExclamation, "Server Warning"
    Case "223"
        MsgBox "Too many groups.", vbExclamation, "Server Warning"
    Case "224", "231"
        MsgBox "Invalid group.", vbExclamation, "Server Warning"
    Case "229"
        MsgBox "Group name too long.", vbExclamation, "Server Warning"
    Case "280"
        MsgBox "Switchboard failed.", vbExclamation, "Server Error"
    Case "281"
        MsgBox "Transfer to switchboard failed.", vbExclamation, "Server Error"
    Case "500"
        MsgBox "Internal server error.", vbExclamation, "Server Error"
    Case "501"
        MsgBox "Database server error.", vbExclamation, "Server Error"
    Case "510"
        MsgBox "File operation failed.", vbExclamation, "Server Error"
    Case "520"
        MsgBox "Memory allocation failed.", vbExclamation, "Server Error"
    Case "600", "910", "912", "918", "919", "921", "922"
        MsgBox "Server is busy.", vbExclamation, "Server Error"
    Case "601", "605"
        MsgBox "Server is unavailable.", vbExclamation, "Server Error"
    Case "602"
        MsgBox "Peer name server is down.", vbExclamation, "Server Error"
    Case "603"
        MsgBox "Database connection failed.", vbExclamation, "Server Error"
    Case "604"
        MsgBox "Server is going down.", vbExclamation, "Server Warning"
    Case "707"
        MsgBox "Could not create connection.", vbExclamation, "Server Error"
    Case "711"
        MsgBox "Write is blocking.", vbExclamation, "Server Error"
    Case "712"
        MsgBox "Session is overloaded.", vbExclamation, "Server Error"
    Case "713"
        NS_Alert "Calling too rapidly.", vbExclamation, "Server Warning"
    Case "714"
        MsgBox "Too many sessions.", vbExclamation, "Server Error"
    Case "717"
        NS_Alert "Bad friend file.", vbExclamation, "Server Error"
    Case "914", "915", "916"
        MsgBox "Server unavailable.", vbExclamation, "Server Error"
    Case "920"
        MsgBox "Not accepting new principles.", vbExclamation, "Server Warning"
    End Select
End Sub

Private Sub objMSN_NS_PropertyReceived(Property As String, Value As String)
    On Error Resume Next
    
    UserProperties.Add Value, Property
End Sub

Private Sub objMSN_NS_SocketError(Description As String)
    If Val(tmrReconnect.Tag) = 0 Then
        ActivateWindow Me
    End If
    
    lblStatus.Tag = False
    Call ChangeGMStatus(Description, True)
    
    If Val(tmrReconnect.Tag) < 5 Then
        tmrReconnect.Enabled = True
        tmrReconnect.Tag = Val(tmrReconnect.Tag) + 1
    End If
End Sub

Private Sub objMSN_NS_StatusChanged(Status As Integer)
    mnuFile_MyStatus_Status(MyStatus).Checked = False
    If Status_AutoIdle And Not Status = msnStatus_Idle Then
        Status_AutoIdle = False
    End If
    MyStatus = Status
    LastStatus = MyStatus
    mnuFile_MyStatus_Status(Status).Checked = True
    lblNick.Caption = CropText(Me, lblNick.Width, objMSN_NS.Nick, " (" & StatusName(MyStatus) & ")")
    Call Form_Resize
    
    If FirstStatus Then
        Call ChangeGMStatus("Signed In", True)
        If objMSN_NS.EmailEnabled Then
            Call UpdateMailStatus
        End If
        FirstStatus = False
        mnuFile_MyStatus.Enabled = True
        mnuTools_ChangeDisplayPic.Enabled = True
        mnuTools_MessageAll.Enabled = True
        mnuTools_IgnoreAll.Enabled = True
        mnuTools_GMScript.Enabled = True
        If frmOptions.Visible Then
            Call frmOptions.Form_Load
        End If
        If AutoIdle Then
            tmrAutoIdle.Enabled = True
        End If
        If FileExists(App.Path & "\Scripts\autoexec.gms") Then
            LoadScript App.Path & "\Scripts\autoexec.gms"
        End If
    Else
        Call LogStatus("Status changed to " & StatusName(Status) & ".")
        Call QueScript(Me, "statuschanged", ConvArray(MyStatus))
    End If
    
    Call UpdateTrayIcon
End Sub

Private Sub AddTVParentNode(Key As String, Text As String, Optional Message As String)
    tvwContacts.Nodes.Add , , Key, Text
    tvwContacts.Nodes(Key).Expanded = True
    tvwContacts.Nodes(Key).Sorted = True
    tvwContacts.Nodes(Key).Bold = True
    tvwContacts.Nodes(Key).ForeColor = 8388608
    tvwContacts.Nodes(Key).Image = 6
    tvwContacts.Nodes(Key).Tag = Message
    
    If Not Message = vbNullString Then
        tvwContacts.Nodes.Add Key, tvwChild, "MSG " & Key, Message
        tvwContacts.Nodes("MSG " & Key).ForeColor = 8421504
    End If
End Sub

Private Sub AddTVChildNode(Parent As String, Key As String, Text As String, Optional Icon As Integer, Optional BackColor As Long)
    On Error Resume Next
    
    If Not NodeExists(Parent) Then
        Select Case Parent
        Case "online"
            AddTVParentNode "online", "Online", "None of your contacts are online"
        Case "offline"
            AddTVParentNode "offline", "Not Online", "None of your contacts are offline"
        End Select
    End If
    
    tvwContacts.Nodes.Add Parent, tvwChild, Key, Text, Icon
    tvwContacts.Nodes(Key).BackColor = BackColor
    tvwContacts.Nodes(Parent).Text = GetParentNodeTitle(Parent)
End Sub

Private Sub DelTVChildNode(Key As String)
    On Error Resume Next
    
    Dim Temp As String
    Temp = tvwContacts.Nodes(Key).Parent.Key
    tvwContacts.Nodes.Remove Key
    tvwContacts.Nodes(Temp).Text = GetParentNodeTitle(Temp)
End Sub

Public Sub RefreshContact(Email As String)
    On Error Resume Next
    
    If InCollection(HiddenContacts, Email) Then
        Exit Sub
    End If
    
    Dim strSelectedItem As String, strTemp As String, strContactTitle As String
    
    strSelectedItem = tvwContacts.SelectedItem.Key
    
    If ViewContactsByEmail Then
        strContactTitle = Email
    Else
        strContactTitle = GetContactAttr(Email, "nick")
    End If
    
    strTemp = StatusName(GetContactAttr(Email, "status"))
    
    If (strTemp <> "Online" And strTemp <> "Offline") Or (SortContactsByGroups = True And GroupOfflineContactsTogether = False) Then
        strContactTitle = strContactTitle & " (" & strTemp & ")"
    End If
    
    If InList(GetContactAttr(Email, "lists"), msnList_Block) Then
        strContactTitle = strContactTitle & " (Blocked)"
    End If
    
    If InCollection(IgnoreList, Email) Then
        strContactTitle = strContactTitle & " (Ignored)"
    End If
    
    Dim lngContactBackColor As Long
    If InList(GetContactAttr(Email, "lists"), msnList_Reverse) Or Not HighlightFakeFriends Then
        lngContactBackColor = vbWindowBackground
    Else
        lngContactBackColor = 16119285
    End If
    
    If SortContactsByGroups Then
        Dim i As Integer, j As Integer
        Call RemoveContact(Email)
        If GetContactAttr(Email, "status") = msnStatus_Offline And GroupOfflineContactsTogether Then
            AddTVChildNode "offline", Email, strContactTitle, ContactIcon(Email), lngContactBackColor
        Else
            For i = 1 To ContactList(Email).Item("groups").Count
                j = ContactList(Email).Item("groups").Item(i)
                AddTVChildNode "GRP " & j, Email & " " & j, strContactTitle, ContactIcon(Email), lngContactBackColor
                SortContacts tvwContacts.Nodes("GRP " & j)
            Next
        End If
    Else
        Call RemoveContact(Email)
        If GetContactAttr(Email, "status") = msnStatus_Offline Then
            AddTVChildNode "offline", Email, strContactTitle, ContactIcon(Email), lngContactBackColor
        Else
            AddTVChildNode "online", Email, strContactTitle, ContactIcon(Email), lngContactBackColor
        End If
    End If
    
    tvwContacts.Nodes(strSelectedItem).Selected = True
End Sub

Public Sub RefreshTreeview()
    On Error Resume Next
    
    tvwContacts.Nodes.Clear
    Dim i As Integer
    
    If SortContactsByGroups Then
    
        For i = 2 To ContactGroups.Count
            AddTVParentNode "GRP " & ContactGroups(i).Item("id"), ContactGroups(i).Item("name"), IIf(GroupOfflineContactsTogether, "Everyone in this group is offline", "You have no contacts in this group")
        Next
        
        tvwContacts.Sorted = True
        tvwContacts.Sorted = False
        
        AddTVParentNode "GRP 0", ContactGroups(1).Item("name"), IIf(GroupOfflineContactsTogether, "Everyone in this group is offline", "You have no contacts in this group")
    
        If GroupOfflineContactsTogether Then
            AddTVParentNode "offline", "Not Online", "None of your contacts are offline"
        End If
    Else
        AddTVParentNode "online", "Online", "None of your contacts are online"
        AddTVParentNode "offline", "Not Online", "None of your contacts are offline"
    End If
    
    For i = 1 To ContactList.Count
        If InList(ContactList(i).Item("lists"), msnList_Forward) Then
            RefreshContact ContactList(i).Item("email")
        End If
    Next
End Sub

Private Sub objMSN_NS_SwitchboardReceived(IP As String, Port As Integer, AuthCode As String)
    On Error Resume Next
    
    PendingIM(1).objMSN_SB.Server = IP
    PendingIM(1).objMSN_SB.Port = Port
    PendingIM(1).objMSN_SB.AuthCode = AuthCode
    PendingIM(1).objMSN_SB.SessionType = SbSession_Call
    PendingIM(1).objMSN_SB.Connect
    If Not InCollection(IMWindows, PendingIM(1).objMSN_SB.Contact) Then
        IMWindows.Add PendingIM(1), PendingIM(1).objMSN_SB.Contact
    End If
    PendingIM.Remove 1
End Sub

Private Sub objMSN_NS_UrlReceived(rru As String, URL As String, ID As Integer)
    OpenMsnURL rru, URL, ID
End Sub

Private Sub picSignIn_Click()
    Call mnuFile_SignIn_Click
End Sub

Private Sub picTrayIcon_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    
    LastActive = Timer
    Dim Msg As Long
    Msg = x / Screen.TwipsPerPixelX
    If Msg = WM_LBUTTONDBLCLK Then
        ActivateWindow Me
    ElseIf Msg = WM_RBUTTONUP Then
        mnuFile_Close.Caption = "&Open Gilly Messenger"
        Me.PopupMenu mnuFile, MenuControlConstants.vbPopupMenuCenterAlign
    End If
End Sub

Private Sub UpdateMailStatus()
    If InboxUnread = 0 Then
        lblEmail.Caption = "No new e-mail messages"
    Else
        lblEmail.Caption = InboxUnread & " new e-mail messages"
    End If
    lblEmail.ToolTipText = "Inbox : " & InboxUnread & Space$(4) & "Folders : " & FoldersUnread
End Sub

Private Function ContactIcon(Email As String) As Integer
    If InList(ContactList(Email).Item("lists"), msnList_Block) Then
        ContactIcon = 5
    Else
        ContactIcon = StatusIcon(ContactList(Email).Item("status"))
    End If
End Function

Private Function StatusIcon(Status As Integer) As Integer
    Select Case Status
        Case msnStatus_Online
            StatusIcon = 1
        Case msnStatus_Offline
            StatusIcon = 2
        Case msnStatus_Busy, msnStatus_OneThePhone
            StatusIcon = 3
        Case Else
            StatusIcon = 4
    End Select
End Function

Private Function NodeExists(Key As String) As Boolean
    On Error GoTo Handler
    Dim strTemp As String
    strTemp = tvwContacts.Nodes(Key).Text
    NodeExists = True
    Exit Function
Handler:
    NodeExists = False
End Function

Private Sub SearchForAContact()
    On Error GoTo Handler
    Dim strKeyWord As String, strContact As String, i As Integer, j As Integer
    
    If mnuContacts_SearchForAContact.Tag = vbNullString Then
        strKeyWord = InputBox("Enter a keyword", "Search for a Contact")
        i = 1
    Else
        strKeyWord = mnuContacts_SearchForAContact.Tag
        i = tvwContacts.SelectedItem.Index + 1
    End If
    
    If Not strKeyWord = vbNullString Then
        strKeyWord = "*" & LCase$(strKeyWord) & "*"
        mnuContacts_SearchForAContact.Tag = strKeyWord
        For j = i To tvwContacts.Nodes.Count
            If NodeIsContact(tvwContacts.Nodes(j).Key) Then
                strContact = LCase$(Split(tvwContacts.Nodes(j).Key)(0))
                If (strContact Like strKeyWord) Or (GetContactComment(strContact) Like strKeyWord) Or (tvwContacts.Nodes(j).Text Like strKeyWord) Then
                    tvwContacts.Nodes(j).Selected = True
                    Exit Sub
                End If
            End If
        Next
    MsgBox "Contact not found!", vbInformation, "Search for a Contact"
    mnuContacts_SearchForAContact.Tag = vbNullString
    End If
Handler:
End Sub

Private Function GetParentNodeTitle(Key As String) As String
    Select Case Split(Key)(0)
    Case "online"
        GetParentNodeTitle = "Online"
    Case "offline"
        GetParentNodeTitle = "Not Online"
    Case "GRP"
        GetParentNodeTitle = ContactGroups("GRP " & Split(Key)(1)).Item("name")
    End Select
    
    Dim i As Integer
    i = tvwContacts.Nodes(Key).Children
        
    If i = 0 Then
        tvwContacts.Nodes.Add Key, tvwChild, "MSG " & Key, tvwContacts.Nodes(Key).Tag
        tvwContacts.Nodes("MSG " & Key).ForeColor = 8421504
    Else
        If NodeExists("MSG " & Key) And i > 1 Then
            tvwContacts.Nodes.Remove "MSG " & Key
            GetParentNodeTitle = GetParentNodeTitle & " (" & i - 1 & ")"
        ElseIf Not NodeExists("MSG " & Key) Then
            GetParentNodeTitle = GetParentNodeTitle & " (" & i & ")"
        End If
    End If
End Function

Private Function GroupExists(GroupName As String) As Boolean
    Dim i As Integer
    For i = 1 To ContactGroups.Count
        If ContactGroups(i).Item("name") = GroupName Then
            GroupExists = True
            Exit Function
        End If
    Next
End Function

Private Function NodeIsGroup(Key As String) As Boolean
    If Split(Key)(0) = "GRP" Then
        NodeIsGroup = True
    End If
End Function

Private Function NodeIsContact(Key As String) As Boolean
    Key = Split(Key)(0)
    If Key <> "GRP" And Key <> "MSG" And Key <> "offline" And Key <> "online" Then
        NodeIsContact = True
    End If
End Function

Private Sub tmrAutoIdle_Timer()
    'this routine is executed every second
    If (GetIdleTime \ 60) >= AutoIdle_Interval Then
        'if idle time reaches/acceeds the autoidle interval
        If MyStatus = msnStatus_Online And Not Status_AutoIdle Then
            'if status has not been changed to idle automatically already
            Status_AutoIdle = True
            objMSN_NS.ChangeStatus msnStatus_Idle
        End If
    ElseIf MyStatus = msnStatus_Idle And Status_AutoIdle Then
        'if some activity happened and the status was set to idle automatically then change it back to online
        Status_AutoIdle = False
        objMSN_NS.ChangeStatus msnStatus_Online
    End If
End Sub

Private Sub tmrGMScript_Events_Timer()
    On Error Resume Next
    
    tmrGMScript_Events.Enabled = False
    
    If Not ScriptQue.Count = 0 Then
        Dim ScriptIndex As Integer, strEvent As String
        ScriptIndex = ScriptQue(1).Item("script")
        strEvent = ScriptQue(1).Item("event")
        
        If strEvent = "messagesent" Then
            ScriptQue(1).Item("source").MsgSentProc = True
        End If
        
        SetCollectionItem GMScripts(ScriptIndex), "pos_" & strEvent, 1
        Do Until GMScripts(ScriptIndex).Item("pos_" & strEvent) > GMScripts(ScriptIndex).Item("script_" & strEvent).Count
            DoEvents
            Call ParseScript(ScriptQue(1).Item("source"), GMScripts(ScriptIndex), strEvent, ScriptQue(1).Item("params"))
        Loop

        If strEvent = "messagesent" Then
            ScriptQue(1).Item("source").MsgSentProc = False
        End If
        
        ScriptQue.Remove 1
    
        If Not ScriptQue.Count = 0 Then
            tmrGMScript_Events.Enabled = True
        End If
    End If
End Sub

Private Sub tmrGMScript_Main_Timer()
    tmrGMScript_Main.Enabled = False
    
    On Error GoTo Handler
    Dim ScriptIndex As Integer, DontStop As Boolean
    For ScriptIndex = 1 To GMScripts.Count
        With GMScripts(ScriptIndex)
            If .Item("script_main").Count > 0 And Not .Item("pos_main") > .Item("script_main").Count Then
                If Timer - Val(.Item("lastexec")) >= .Item("sleep") Then
                    SetCollectionItem GMScripts(ScriptIndex), "sleep", 0
                    If ParseScript(Me, GMScripts(ScriptIndex), "main") Then
                        If .Item("pos_main") > .Item("script_main").Count Then
                            If .Item("eventcount") = 0 Then
                                Call EndScript(GMScripts(ScriptIndex), True)
                            End If
                        Else
                            DontStop = True
                            SetCollectionItem GMScripts(ScriptIndex), "lastexec", Timer
                        End If
                    End If
                Else
                    DontStop = True
                End If
            End If
        End With
Handler:
        On Error Resume Next
    Next
    
    If GMScripts.Count > 0 And DontStop Then
        tmrGMScript_Main.Enabled = True
    End If
End Sub

Private Sub tmrNewsScroller1_Timer()
    On Error Resume Next
    
    lblNews.Top = lblNews.Top - 1
    If lblNews.Top = (picNews.Height - lblNews.Height) \ 2 Then
        tmrNewsScroller1.Enabled = False
        tmrNewsScroller2.Enabled = True
    End If
End Sub

Private Sub tmrNewsScroller2_Timer()
    On Error Resume Next
    
    If Not (Me.WindowState = vbMinimized Or Me.Visible = False) Then
        If ArraySize(NewsLines) >= 0 Then
            lblNews.Left = lblNews.Left - 1
            If lblNews.Left <= -lblNews.Width Then
                NewsPointer = NewsPointer + 1
                If NewsPointer > ArraySize(NewsLines) Then
                    picNews.Visible = False
                    lblNews.Visible = False
                    Erase NewsLines
                    Call Form_Resize
                    Exit Sub
                End If
                If PathIsURL(CStr(Split(NewsLines(NewsPointer))(0))) Then
                    lblNews.Tag = Left$(NewsLines(NewsPointer), InStr(NewsLines(NewsPointer), " ") - 1)
                    lblNews.Caption = Right$(NewsLines(NewsPointer), Len(NewsLines(NewsPointer)) - InStr(NewsLines(NewsPointer), " "))
                    lblNews.MousePointer = vbCustom
                Else
                    lblNews.Tag = vbNullString
                    lblNews.Caption = NewsLines(NewsPointer)
                    lblNews.MousePointer = vbDefault
                End If
                lblNews.Top = picNews.Height
                If Not lblNews.Width > picNews.Width Then
                    lblNews.Left = (picNews.Width - lblNews.Width) \ 2
                Else
                    lblNews.Left = 0
                End If
                tmrNewsScroller2.Enabled = False
                tmrNewsScroller1.Enabled = True
            End If
        End If
    End If
    NewsInterval = NewsInterval + 1
    If NewsInterval >= 12000 Then
        NewsInterval = 0
        wskNews.Close
        wskNews.Connect "www.cracksoft.net", 80
    End If
End Sub

Private Sub tmrPing_Timer()
    objMSN_NS.PingServer
End Sub

Private Sub tmrReconnect_Timer()
    tmrReconnect.Enabled = False
    If objMSN_NS.Login <> vbNullString And objMSN_NS.Password <> vbNullString And wskNS.LocalIP <> "127.0.0.1" And wskNS.LocalIP <> "0.0.0.0" Then
            Call SignIn(objMSN_NS.Login, objMSN_NS.Password)
    End If
End Sub

Private Sub tmrTrayAnim_Timer()
    If Val(tmrTrayAnim.Tag) Mod 2 = 0 Then
        Call ChangeTrayIcon(picTrayIcon.Picture.Handle)
    Else
        Call UpdateTrayIcon
    End If
    tmrTrayAnim.Tag = Val(tmrTrayAnim.Tag) - 1
    If tmrTrayAnim.Tag = "0" Then
        tmrTrayAnim.Tag = "20"
        tmrTrayAnim.Enabled = False
    End If
End Sub

Private Sub tvwContacts_Collapse(ByVal Node As MSComctlLib.Node)
    Node.Image = 7
End Sub

Private Sub tvwContacts_DblClick()
    On Error GoTo Handler
    If objMSN_NS.State = NsState_SignedIn Then
        If NodeIsContact(tvwContacts.SelectedItem.Key) Then
            Dim strContact As String
            strContact = CStr(Split(tvwContacts.SelectedItem.Key)(0))
            If IsOnline(strContact) Then
                StartChat strContact, , , True
            End If
        End If
    End If
Handler:
End Sub

Private Sub tvwContacts_Expand(ByVal Node As MSComctlLib.Node)
    Node.Image = 6
End Sub

Private Sub tvwContacts_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    Dim strSelectedItem As String
    strSelectedItem = tvwContacts.SelectedItem.Key
    If KeyCode = vbKeyF2 Then
        If NodeIsGroup(strSelectedItem) Then
            RenameGroup (Split(strSelectedItem)(1))
        End If
    ElseIf KeyCode = vbKeyReturn Then
        If Shift = vbAltMask Then
            If NodeIsContact(strSelectedItem) Then
                ShowBuddyProperties Me, CStr(Split(strSelectedItem)(0))
            End If
        Else
            Call tvwContacts_DblClick
        End If
    ElseIf KeyCode = vbKeyDelete Then
        If NodeIsGroup(strSelectedItem) Then
            If MsgBox("Are you sure you want to delete this group?", vbYesNo) = vbYes Then
                objMSN_NS.RemoveGroup Val(Split(strSelectedItem)(1))
            End If
        ElseIf NodeIsContact(strSelectedItem) Then
            If MsgBox("Are you sure you want to delete " & Split(strSelectedItem)(0) & " from your contact list?", vbYesNo) = vbYes Then
                objMSN_NS.RemoveContact msnList_Forward, CStr(Split(strSelectedItem)(0))
            End If
        End If
    End If
End Sub

Private Sub tvwContacts_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo Handler
    If Button = vbRightButton Then
        If NodeIsGroup(tvwContacts.SelectedItem.Key) Then
            mnuGroup.Tag = Split(tvwContacts.SelectedItem.Key)(1)
            mnuGroup_DeleteGroup.Enabled = IIf(tvwContacts.SelectedItem.Key = "GRP 0", False, True)
            mnuGroup_SaveGroupToAFile.Enabled = Not (NodeExists("MSG " & tvwContacts.SelectedItem.Key))
            PopupMenu mnuGroup
            
        ElseIf NodeIsContact(tvwContacts.SelectedItem.Key) Then
            Dim i As Integer
            
            mnuContact.Tag = Split(tvwContacts.SelectedItem.Key)(0)
            mnuContact_SendAnInstantMessage.Enabled = Not (GetContactAttr(mnuContact.Tag, "status") = msnStatus_Offline)
            mnuContact_SendAFileOrPhoto.Enabled = mnuContact_SendAnInstantMessage.Enabled
            mnuContact_SendEmail.Caption = "Send &Email (" & mnuContact.Tag & ")"
            mnuContact_Block.Caption = IIf(InList(GetContactAttr(mnuContact.Tag, "lists"), msnList_Block), "Un&block", "&Block")
            mnuContact_Ignore.Caption = IIf(InCollection(IgnoreList, mnuContact.Tag), "Un&ignore", "&Ignore")
            mnuContact_PopupFilter.Caption = IIf(InCollection(PopupFilter, mnuContact.Tag), "Remove from Pop&up Filter", "Add to Pop&up Filter")
            mnuContact_SoundFilter.Caption = IIf(InCollection(SoundFilter, mnuContact.Tag), "Remove from S&ound Filter", "Add to S&ound Filter")
            
            If Not (SortContactsByGroups = False Or (GroupOfflineContactsTogether = True And GetContactAttr(mnuContact.Tag, "status") = msnStatus_Offline)) Then
                mnuContact_CopyContactTo.Tag = Split(tvwContacts.SelectedItem.Parent.Key)(1)
                mnuContact_MoveContactTo.Tag = mnuContact_CopyContactTo.Tag
                mnuContact_RemoveContactFromGroup.Tag = mnuContact_MoveContactTo.Tag
                mnuContact_Hide.Tag = mnuContact_RemoveContactFromGroup.Tag
                
                ClearSubMenu mnuContact_CopyContactTo_Group
                ClearSubMenu mnuContact_MoveContactTo_Group
                
                For i = 1 To ContactGroups.Count
                    If Not InCollection(ContactList(mnuContact.Tag).Item("groups"), "GRP " & ContactGroups(i).Item("id")) Then
                        AddSubMenu mnuContact_CopyContactTo_Group, ContactGroups(i).Item("name"), ContactGroups(i).Item("id")
                        AddSubMenu mnuContact_MoveContactTo_Group, ContactGroups(i).Item("name"), ContactGroups(i).Item("id")
                    End If
                Next
                
                mnuContact_CopyContactTo.Enabled = (mnuContact_CopyContactTo_Group.Count > 1)
                mnuContact_MoveContactTo.Enabled = (mnuContact_MoveContactTo_Group.Count > 1)
                mnuContact_RemoveContactFromGroup.Enabled = True
            Else
                mnuContact_CopyContactTo.Enabled = False
                mnuContact_MoveContactTo.Enabled = False
                mnuContact_RemoveContactFromGroup.Enabled = False
            End If
            mnuContact_OpenMessageHistory.Enabled = FileExists(MessageHistoryFolder & "\" & mnuContact.Tag & ".txt")
            
            PopupMenu mnuContact
        End If
    End If
Handler:
End Sub

Private Sub RenameGroup(GroupID As Integer)
    On Error Resume Next
    
    Dim OldName As String, NewName As String
    OldName = ContactGroups("GRP " & GroupID).Item("name")
    NewName = InputBox("Enter the new name for group.", "Rename Group", OldName)
    If NewName <> OldName And NewName <> vbNullString Then
        objMSN_NS.RenameGroup GroupID, NewName
    End If
End Sub

Private Sub txtNick_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And GetKeyState(vbKeyShift) = False And GetKeyState(vbKeyControl) = False Then
        txtNick.Visible = False
        If Not txtNick.Text = vbNullString Then
            If InStr(txtNick.Text, vbCrLf) = 0 And InStr(txtNick.Text, vbLf) > 0 Then
                txtNick.Text = Replace$(txtNick.Text, vbLf, vbCrLf)
            End If
            objMSN_NS.ChangeNick txtNick.Text
        End If
        KeyAscii = 0
    ElseIf KeyAscii = vbKeyEscape Then
        txtNick.Visible = False
    End If
End Sub

Private Sub txtNick_LostFocus()
    txtNick.Visible = False
End Sub

Private Sub SwapNodes(Node1 As Node, Node2 As Node)
    Dim Expanded As Boolean, Tag As String, ForeColor As Long, BackColor As Long, Bold As Boolean, Sorted As Boolean, Key As String, Image As Integer, Text As String
    
    Expanded = Node1.Expanded
    Tag = Node1.Tag
    ForeColor = Node1.ForeColor
    BackColor = Node1.BackColor
    Bold = Node1.Bold
    Sorted = Node1.Sorted
    Key = Node1.Key
    Image = Node1.Image
    Text = Node1.Text
    
    Node1.Expanded = Node2.Expanded
    Node1.Tag = Node2.Tag
    Node1.ForeColor = Node2.ForeColor
    Node1.BackColor = Node2.BackColor
    Node1.Bold = Node2.Bold
    Node1.Sorted = Node2.Sorted
    Node1.Key = "_" & Node2.Key
    Node1.Image = Node2.Image
    Node1.Text = Node2.Text
    
    Node2.Expanded = Expanded
    Node2.Tag = Tag
    Node2.ForeColor = ForeColor
    Node2.BackColor = BackColor
    Node2.Bold = Bold
    Node2.Sorted = Sorted
    Node2.Key = Key
    Node2.Image = Image
    Node2.Text = Text
    
    Node1.Key = Right$(Node1.Key, Len(Node1.Key) - 1)
End Sub

Private Sub SortContacts(Group As Node)
    On Error GoTo Handler
    
    Dim TempNode As Node, i As Integer
    
    Group.Sorted = True
    Group.Sorted = False
    
    Set TempNode = Group.Child.FirstSibling
    
    Do
        Set TempNode = TempNode.Next
        If NodeIsContact(TempNode.Key) Then
            If Not GetContactAttr(CStr(Split(TempNode.Key)(0)), "status") = msnStatus_Offline Then
                BringContactUp TempNode
            End If
        End If
    Loop
Handler:
End Sub

Private Sub BringContactUp(Contact As Node)
    On Error GoTo Handler
    
    Dim Temp As String
    Do
        Temp = Contact.Previous.Key
        If NodeIsContact(Temp) Then
            If GetContactAttr(CStr(Split(Temp)(0)), "status") = msnStatus_Offline Then
                SwapNodes Contact, Contact.Previous
            Else
                Exit Do
            End If
            Set Contact = Contact.Previous
        End If
    Loop
Handler:
End Sub

Public Sub RemoveContact(Email As String, Optional GroupID = -1)
    On Error Resume Next
    
    If SortContactsByGroups Then
        If GroupID = -1 Then
            Dim i As Integer
            For i = 1 To ContactList(Email).Item("groups").Count
                DelTVChildNode Email & " " & ContactList(Email).Item("groups").Item(i)
            Next
            If GroupOfflineContactsTogether Then
                DelTVChildNode Email
            End If
        Else
            DelTVChildNode Email & " " & GroupID
        End If
    Else
        DelTVChildNode Email
    End If
End Sub

Private Sub wskNews_Close()
    wskNews.Close
    Dim intPos As Integer
    intPos = InStr(NewsData, vbCrLf & vbCrLf)
    If Not intPos = 0 Then
        If InStr(NewsData, "200 OK") > 0 Or InStr(NewsData, "100 Continue") > 0 Then
            NewsData = Right$(NewsData, Len(NewsData) - intPos - 3)
        Else
            NewsData = vbNullString
            wskNews.Close
        End If
    End If
    If Not Len(NewsData) = 0 Then
        If NewsData <> lblNews.Tag Then
            lblNews.Tag = NewsData
            NewsLines = Split(NewsData, vbLf)
            Dim i As Integer
            For i = 0 To UBound(NewsLines)
                NewsLines(i) = Replace(Trim$(NewsLines(i)), vbCr, vbNullString)
                If NewsLines(i) = vbNullString Or Len(NewsLines(i)) < 10 Then
                    NewsLines(i) = Chr$(0)
                End If
            Next
            NewsLines = Filter(NewsLines, Chr$(0), False)
            If ArraySize(NewsLines) >= 0 Then
                If objMSN_NS.State = NsState_SignedIn Then
                    Call initNews
                End If
            End If
        End If
    End If
    NewsData = vbNullString
End Sub

Private Sub wskNews_Connect()
    On Error Resume Next
    Dim Data As String
    Data = URL_Encode("id=GM " & App.Major & "." & App.Minor & "." & App.Revision)
    
    wskNews.SendData "POST /public/news.php HTTP/1.1" & vbCrLf & _
    "Host: www.cracksoft.net" & vbCrLf & _
    "Connection: close" & vbCrLf & _
    "Accept: */*" & vbCrLf & _
    "Content-Type: application/x-www-form-urlencoded" & vbCrLf & _
    "Content-Length: " & Len(Data) & vbCrLf & vbCrLf & Data
End Sub

Private Sub wskNews_DataArrival(ByVal bytesTotal As Long)
    If wskNews.State = sckConnected Then
        Dim Data As String
        wskNews.GetData Data
        NewsData = NewsData & Data
    End If
End Sub

Private Sub initNews()
    NewsPointer = 0
    If PathIsURL(CStr(Split(NewsLines(NewsPointer))(0))) Then
        lblNews.Tag = Left$(NewsLines(NewsPointer), InStr(NewsLines(NewsPointer), " ") - 1)
        lblNews.Caption = Right$(NewsLines(NewsPointer), Len(NewsLines(NewsPointer)) - InStr(NewsLines(NewsPointer), " "))
        lblNews.MousePointer = vbCustom
    Else
        lblNews.Tag = vbNullString
        lblNews.Caption = NewsLines(NewsPointer)
        lblNews.MousePointer = vbDefault
    End If
    lblNews.Top = picNews.Height
    If Not lblNews.Width > picNews.Width Then
        lblNews.Left = (picNews.Width - lblNews.Width) / 2
    Else
        lblNews.Left = 0
    End If
    picNews.Visible = True
    lblNews.Visible = True
    Call Form_Resize
    tmrNewsScroller1.Enabled = True
End Sub

Private Sub NS_Alert(Message As String, Buttons As VbMsgBoxStyle, Title As String)
    If Not Message = LastNsError Then
        MsgBox Message, Buttons, Title
        LastNsError = Message
    End If
End Sub

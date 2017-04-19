VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5835
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   465
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   389
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   4560
      TabIndex        =   60
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   330
      Left            =   3360
      TabIndex        =   59
      Top             =   6480
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   11033
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   529
      WordWrap        =   0   'False
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmOptions.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Line7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblSignIn"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Line8"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblAlerts"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Line9"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblSaveMyStatusHistoryInThisFolder"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkStartWithWindows"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "chkOpenMainWindowOnStart"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "chkAlertOnContactOnline"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "chkAlertOnMessageReceived"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chkAlertOnEmailReceived"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "chkSoundAlerts"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdSounds"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "chkStatusHistory"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtStatusHistoryFolder"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmdChangeStatusHistoryFolder"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "chkAutoIdle"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtAutoIdleInterval"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cmdPopupFilter"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "chkBlockAlert"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "chkSendDisplayPic"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "chkReceiveDisplayPic"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "chkBlockPopups"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).ControlCount=   23
      TabCaption(1)   =   "Messages"
      TabPicture(1)   =   "frmOptions.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblMessageText"
      Tab(1).Control(1)=   "Line1"
      Tab(1).Control(2)=   "lblChangeTheFontAndColorOfMyInstantMessages"
      Tab(1).Control(3)=   "lblFileTransfer"
      Tab(1).Control(4)=   "Line2"
      Tab(1).Control(5)=   "lblPutFilesReceivedFromOthersInThisFolder"
      Tab(1).Control(6)=   "lblSaveMyConversationsInThisFolder"
      Tab(1).Control(7)=   "Line3"
      Tab(1).Control(8)=   "lblMessageHistory"
      Tab(1).Control(9)=   "Line10"
      Tab(1).Control(10)=   "lblFTPPort"
      Tab(1).Control(11)=   "chkShowEmoticons"
      Tab(1).Control(12)=   "cmdChangeFont"
      Tab(1).Control(13)=   "cmdChangeColor"
      Tab(1).Control(14)=   "txtReceivedFilesFolder"
      Tab(1).Control(15)=   "cmdChangeReceivedFilesFolder"
      Tab(1).Control(16)=   "cmdChangeMessageHistoryFolder"
      Tab(1).Control(17)=   "txtMessageHistoryFolder"
      Tab(1).Control(18)=   "chkMessageHistory"
      Tab(1).Control(19)=   "chkShowIMWindowOnMsg"
      Tab(1).Control(20)=   "chkDisableTypingMsgNotification"
      Tab(1).Control(21)=   "txtFTPPort"
      Tab(1).ControlCount=   22
      TabCaption(2)   =   "Phone"
      TabPicture(2)   =   "frmOptions.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Line4"
      Tab(2).Control(1)=   "lblPhoneNumbers"
      Tab(2).Control(2)=   "lblTypeThePhoneNumbersThatYouWantPeopleOnYourAllowListToSee"
      Tab(2).Control(3)=   "lblMyCountryRegionCode"
      Tab(2).Control(4)=   "lblAreaCodeAndPhoneNumber"
      Tab(2).Control(5)=   "lblMyHomePhone"
      Tab(2).Control(6)=   "lblMyWorkPhone"
      Tab(2).Control(7)=   "lblMyMobilePhone"
      Tab(2).Control(8)=   "cmbCountryRegionCode"
      Tab(2).Control(9)=   "txtHomePhoneNumber"
      Tab(2).Control(10)=   "txtHomePhoneCode"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "txtWorkPhoneNumber"
      Tab(2).Control(12)=   "txtWorkPhoneCode"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "txtMobilePhoneNumber"
      Tab(2).Control(14)=   "txtMobilePhoneCode"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).ControlCount=   15
      TabCaption(3)   =   "Privacy"
      TabPicture(3)   =   "frmOptions.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "chkHighlightFakeFriends"
      Tab(3).Control(1)=   "cmdClean"
      Tab(3).Control(2)=   "chkGTC"
      Tab(3).Control(3)=   "cmdViewReverseList"
      Tab(3).Control(4)=   "cmdBlock"
      Tab(3).Control(5)=   "cmdAllow"
      Tab(3).Control(6)=   "lstBlock"
      Tab(3).Control(7)=   "lstAllow"
      Tab(3).Control(8)=   "chkBLP"
      Tab(3).Control(9)=   "lblSeeWhoHasAddedYouToTheirContactList"
      Tab(3).Control(10)=   "lblCantSendMeMessages"
      Tab(3).Control(11)=   "lblCantSeeMyOnlineStatus"
      Tab(3).Control(12)=   "lblCanSendMeMessages"
      Tab(3).Control(13)=   "lblCanSeeMyOnlineStatus"
      Tab(3).Control(14)=   "lblMyBlockList"
      Tab(3).Control(15)=   "lblMyAllowList"
      Tab(3).Control(16)=   "Line5"
      Tab(3).Control(17)=   "lblVisibiltyAndPrivacy"
      Tab(3).ControlCount=   18
      TabCaption(4)   =   "Cache"
      TabPicture(4)   =   "frmOptions.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lblResetServerSettings"
      Tab(4).Control(1)=   "lblResetAppSettings"
      Tab(4).Control(2)=   "lblClearLoginCache"
      Tab(4).Control(3)=   "lblResetSettings"
      Tab(4).Control(4)=   "Line6"
      Tab(4).Control(5)=   "cmdResetServerSettings"
      Tab(4).Control(6)=   "cmdResetAppSettings"
      Tab(4).Control(7)=   "cmdClearLoginCache"
      Tab(4).ControlCount=   8
      TabCaption(5)   =   "Miscellaneous"
      TabPicture(5)   =   "frmOptions.frx":0098
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "lblTransparency"
      Tab(5).Control(1)=   "SldrTransparency"
      Tab(5).Control(2)=   "fmBrowser"
      Tab(5).Control(3)=   "fmEmail"
      Tab(5).ControlCount=   4
      Begin VB.CheckBox chkHighlightFakeFriends 
         Caption         =   "&Highlight those people on my contact list who have deleted me"
         Height          =   255
         Left            =   -74760
         TabIndex        =   44
         Top             =   4320
         Width           =   5055
      End
      Begin VB.CheckBox chkBlockPopups 
         Caption         =   "Block alerts when i'm running a f&ull screen program"
         Height          =   255
         Left            =   840
         TabIndex        =   10
         Top             =   3480
         Width           =   4335
      End
      Begin VB.Frame fmEmail 
         Caption         =   "Email"
         Height          =   2055
         Left            =   -74760
         TabIndex        =   91
         Top             =   1920
         Width           =   5055
         Begin VB.Frame fmEmailOptions 
            BorderStyle     =   0  'None
            Height          =   975
            Left            =   240
            TabIndex        =   92
            Top             =   960
            Width           =   4695
            Begin VB.OptionButton optEmailWeb 
               Caption         =   "Webpage:"
               Enabled         =   0   'False
               Height          =   255
               Left            =   120
               TabIndex        =   54
               Top             =   135
               Width           =   1095
            End
            Begin VB.TextBox txtEmailWeb 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1320
               TabIndex        =   55
               Top             =   120
               Width           =   2175
            End
            Begin VB.OptionButton optEmailApp 
               Caption         =   "Application:"
               Enabled         =   0   'False
               Height          =   255
               Left            =   120
               TabIndex        =   56
               Top             =   525
               Width           =   1215
            End
            Begin VB.TextBox txtEmailApp 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   57
               TabStop         =   0   'False
               Top             =   510
               Width           =   2175
            End
            Begin VB.CommandButton cmdEmailAppBrowse 
               Caption         =   "Browse..."
               Enabled         =   0   'False
               Height          =   330
               Left            =   3600
               TabIndex        =   58
               Top             =   487
               Width           =   1095
            End
         End
         Begin VB.OptionButton optEmailDefault 
            Caption         =   "Default"
            Height          =   195
            Left            =   120
            TabIndex        =   52
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton optEmailCustom 
            Caption         =   "Custom"
            Height          =   195
            Left            =   120
            TabIndex        =   53
            Top             =   720
            Width           =   975
         End
      End
      Begin VB.Frame fmBrowser 
         Caption         =   "Browser"
         Height          =   1215
         Left            =   -74760
         TabIndex        =   90
         Top             =   600
         Width           =   5055
         Begin VB.CommandButton cmdBrowserBrowse 
            Caption         =   "Browse..."
            Enabled         =   0   'False
            Height          =   330
            Left            =   3840
            TabIndex        =   51
            Top             =   652
            Width           =   1095
         End
         Begin VB.TextBox txtBrowser 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   675
            Width           =   2655
         End
         Begin VB.OptionButton optBrowserCustom 
            Caption         =   "Custom"
            Height          =   195
            Left            =   120
            TabIndex        =   49
            Top             =   720
            Width           =   975
         End
         Begin VB.OptionButton optBrowserDefault 
            Caption         =   "Default"
            Height          =   195
            Left            =   120
            TabIndex        =   48
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.TextBox txtFTPPort 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -70200
         MaxLength       =   4
         TabIndex        =   23
         Top             =   2400
         Width           =   495
      End
      Begin VB.CheckBox chkReceiveDisplayPic 
         Caption         =   "Show display pictures from &others in IM conversations"
         Height          =   255
         Left            =   840
         TabIndex        =   12
         Top             =   4320
         Width           =   4335
      End
      Begin VB.CheckBox chkSendDisplayPic 
         Caption         =   "&Show my display picture and allow others to see it"
         Height          =   255
         Left            =   840
         TabIndex        =   11
         Top             =   3960
         Width           =   4335
      End
      Begin VB.CheckBox chkDisableTypingMsgNotification 
         Caption         =   "Disable message typing notification"
         Height          =   375
         Left            =   -74760
         TabIndex        =   28
         Top             =   4800
         Width           =   4335
      End
      Begin VB.CheckBox chkBlockAlert 
         Caption         =   "Display block alerts when &offline contact start conversation"
         Height          =   255
         Left            =   840
         TabIndex        =   7
         Top             =   2640
         Width           =   4455
      End
      Begin VB.CommandButton cmdPopupFilter 
         Caption         =   "&Popup Filter"
         Height          =   330
         Left            =   4200
         TabIndex        =   4
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CheckBox chkShowIMWindowOnMsg 
         Caption         =   "Show IM Window only when contact sends a message"
         Height          =   375
         Left            =   -74760
         TabIndex        =   27
         Top             =   4440
         Width           =   4335
      End
      Begin VB.TextBox txtAutoIdleInterval 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4080
         TabIndex        =   14
         Top             =   4665
         Width           =   285
      End
      Begin VB.CheckBox chkAutoIdle 
         Caption         =   "Sho&w me as ""Idle"" when I'm inactive for            minutes"
         Height          =   255
         Left            =   840
         TabIndex        =   13
         Top             =   4680
         Width           =   4455
      End
      Begin VB.CommandButton cmdChangeStatusHistoryFolder 
         Caption         =   "Cha&nge..."
         Height          =   330
         Left            =   4200
         TabIndex        =   17
         Top             =   5640
         Width           =   1095
      End
      Begin VB.TextBox txtStatusHistoryFolder 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   5640
         Width           =   3255
      End
      Begin VB.CheckBox chkStatusHistory 
         Caption         =   "Automatically &keep a history of my Messenger status"
         Height          =   375
         Left            =   840
         TabIndex        =   15
         Top             =   5040
         Width           =   4095
      End
      Begin VB.CommandButton cmdSounds 
         Caption         =   "&Sounds..."
         Height          =   330
         Left            =   4200
         TabIndex        =   9
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CheckBox chkSoundAlerts 
         Caption         =   "&Play a sound when contacts sign in or send a message"
         Height          =   375
         Left            =   840
         TabIndex        =   8
         Top             =   3000
         Width           =   3135
      End
      Begin VB.CommandButton cmdClearLoginCache 
         Caption         =   "Clear &Login Cache"
         Height          =   375
         Left            =   -73155
         TabIndex        =   45
         Top             =   1305
         Width           =   1695
      End
      Begin VB.CommandButton cmdResetAppSettings 
         Caption         =   "Reset &App Settings"
         Height          =   375
         Left            =   -73155
         TabIndex        =   46
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CommandButton cmdResetServerSettings 
         Caption         =   "Reset &Server Settings"
         Height          =   375
         Left            =   -73155
         TabIndex        =   47
         Top             =   3000
         Width           =   1695
      End
      Begin VB.CheckBox chkAlertOnEmailReceived 
         Caption         =   "Display alerts when &e-mail is received"
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   2280
         Width           =   4335
      End
      Begin VB.CheckBox chkAlertOnMessageReceived 
         Caption         =   "Display alerts when a &message is received"
         Height          =   255
         Left            =   840
         TabIndex        =   5
         Top             =   1920
         Width           =   4335
      End
      Begin VB.CheckBox chkAlertOnContactOnline 
         Caption         =   "&Display alerts when contacts come online"
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   1560
         Width           =   4335
      End
      Begin VB.CheckBox chkOpenMainWindowOnStart 
         Caption         =   "&Open Messenger main window when Messenger starts"
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   1080
         Width           =   4335
      End
      Begin VB.CheckBox chkStartWithWindows 
         Caption         =   "Automatically &run Messenger when I log on to Windows"
         Height          =   255
         Left            =   840
         TabIndex        =   1
         Top             =   720
         Width           =   4335
      End
      Begin VB.CommandButton cmdClean 
         Caption         =   "&Clean"
         Enabled         =   0   'False
         Height          =   330
         Left            =   -72720
         TabIndex        =   40
         Top             =   3000
         Width           =   975
      End
      Begin VB.CheckBox chkGTC 
         Caption         =   "Alert &me when other people add me to their contact list"
         Height          =   255
         Left            =   -74760
         TabIndex        =   43
         Top             =   3960
         Width           =   4335
      End
      Begin VB.CommandButton cmdViewReverseList 
         Caption         =   "&View..."
         Height          =   330
         Left            =   -70680
         TabIndex        =   42
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton cmdBlock 
         Caption         =   "Bloc&k >>"
         Enabled         =   0   'False
         Height          =   330
         Left            =   -72720
         TabIndex        =   39
         Top             =   2520
         Width           =   975
      End
      Begin VB.CommandButton cmdAllow 
         Caption         =   "<< A&llow"
         Enabled         =   0   'False
         Height          =   330
         Left            =   -72720
         TabIndex        =   38
         Top             =   2040
         Width           =   975
      End
      Begin VB.ListBox lstBlock 
         Appearance      =   0  'Flat
         Height          =   1395
         Left            =   -71640
         TabIndex        =   41
         Top             =   2040
         Width           =   1935
      End
      Begin VB.ListBox lstAllow 
         Appearance      =   0  'Flat
         Height          =   1395
         Left            =   -74760
         TabIndex        =   37
         Top             =   2040
         Width           =   1935
      End
      Begin VB.CheckBox chkBLP 
         Caption         =   "&Only people on my Allow List can see my status and send me messages"
         Height          =   375
         Left            =   -74760
         TabIndex        =   36
         Top             =   720
         Width           =   5055
      End
      Begin VB.TextBox txtMobilePhoneCode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73320
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox txtMobilePhoneNumber 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72720
         TabIndex        =   35
         Top             =   3000
         Width           =   3015
      End
      Begin VB.TextBox txtWorkPhoneCode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73320
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox txtWorkPhoneNumber 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72720
         TabIndex        =   33
         Top             =   2640
         Width           =   3015
      End
      Begin VB.TextBox txtHomePhoneCode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73320
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox txtHomePhoneNumber 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72720
         TabIndex        =   31
         Top             =   2280
         Width           =   3015
      End
      Begin VB.ComboBox cmbCountryRegionCode 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmOptions.frx":00B4
         Left            =   -74760
         List            =   "frmOptions.frx":00B6
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1560
         Width           =   5055
      End
      Begin VB.CheckBox chkMessageHistory 
         Caption         =   "Automatically &keep a history of my conversations"
         Height          =   375
         Left            =   -74760
         TabIndex        =   24
         Top             =   3120
         Width           =   3735
      End
      Begin VB.TextBox txtMessageHistoryFolder 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -74760
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   3720
         Width           =   3855
      End
      Begin VB.CommandButton cmdChangeMessageHistoryFolder 
         Caption         =   "Cha&nge..."
         Height          =   330
         Left            =   -70800
         TabIndex        =   26
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CommandButton cmdChangeReceivedFilesFolder 
         Caption         =   "&Change..."
         Height          =   330
         Left            =   -71400
         TabIndex        =   22
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox txtReceivedFilesFolder 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -74760
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   2400
         Width           =   3255
      End
      Begin VB.CommandButton cmdChangeColor 
         Caption         =   "Change C&olor"
         Height          =   330
         Left            =   -70920
         TabIndex        =   19
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdChangeFont 
         Caption         =   "Change &Font"
         Height          =   330
         Left            =   -72240
         TabIndex        =   18
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CheckBox chkShowEmoticons 
         Caption         =   "&Show emoticons in instant messages"
         Height          =   375
         Left            =   -74760
         TabIndex        =   20
         Top             =   1440
         Width           =   3015
      End
      Begin MSComctlLib.Slider SldrTransparency 
         Height          =   255
         Left            =   -73680
         TabIndex        =   93
         Top             =   4200
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   44
         SmallChange     =   20
         Max             =   220
         TickFrequency   =   20
      End
      Begin VB.Label lblTransparency 
         AutoSize        =   -1  'True
         Caption         =   "Transparency:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74760
         TabIndex        =   94
         Top             =   4200
         Width           =   1020
      End
      Begin VB.Label lblFTPPort 
         AutoSize        =   -1  'True
         Caption         =   "Port:"
         Height          =   195
         Left            =   -70200
         TabIndex        =   89
         Top             =   2160
         Width           =   330
      End
      Begin VB.Line Line10 
         BorderColor     =   &H80000011&
         X1              =   -74760
         X2              =   -69720
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Label lblSaveMyStatusHistoryInThisFolder 
         AutoSize        =   -1  'True
         Caption         =   "Save my status history in this folder"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   840
         TabIndex        =   88
         Top             =   5400
         Width           =   2460
      End
      Begin VB.Line Line9 
         BorderColor     =   &H80000011&
         X1              =   240
         X2              =   5280
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Line Line6 
         BorderColor     =   &H80000011&
         X1              =   -73680
         X2              =   -69720
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lblResetSettings 
         AutoSize        =   -1  'True
         Caption         =   "Reset Settings"
         Height          =   195
         Left            =   -74760
         TabIndex        =   87
         Top             =   480
         Width           =   1035
      End
      Begin VB.Label lblClearLoginCache 
         AutoSize        =   -1  'True
         Caption         =   "Clear the list of logins cached from signin dialog"
         Height          =   195
         Left            =   -74760
         TabIndex        =   86
         Top             =   960
         Width           =   3330
      End
      Begin VB.Label lblResetAppSettings 
         AutoSize        =   -1  'True
         Caption         =   "Reset your personal settings"
         Height          =   195
         Left            =   -74760
         TabIndex        =   85
         Top             =   1800
         Width           =   1995
      End
      Begin VB.Label lblResetServerSettings 
         AutoSize        =   -1  'True
         Caption         =   "Reset your server settings"
         Height          =   195
         Left            =   -74760
         TabIndex        =   84
         Top             =   2640
         Width           =   1830
      End
      Begin VB.Label lblAlerts 
         AutoSize        =   -1  'True
         Caption         =   "Alerts"
         Height          =   195
         Left            =   240
         TabIndex        =   83
         Top             =   1320
         Width           =   390
      End
      Begin VB.Line Line8 
         BorderColor     =   &H80000011&
         X1              =   690
         X2              =   5280
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label lblSignIn 
         AutoSize        =   -1  'True
         Caption         =   "Sign In"
         Height          =   195
         Left            =   240
         TabIndex        =   82
         Top             =   480
         Width           =   495
      End
      Begin VB.Line Line7 
         BorderColor     =   &H80000011&
         X1              =   800
         X2              =   5280
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lblSeeWhoHasAddedYouToTheirContactList 
         AutoSize        =   -1  'True
         Caption         =   "See who has added you to their contact list"
         Height          =   195
         Left            =   -74760
         TabIndex        =   81
         Top             =   3600
         Width           =   3060
      End
      Begin VB.Label lblCantSendMeMessages 
         AutoSize        =   -1  'True
         Caption         =   "Can't send me messages"
         Height          =   195
         Left            =   -71640
         TabIndex        =   80
         Top             =   1800
         Width           =   1755
      End
      Begin VB.Label lblCantSeeMyOnlineStatus 
         AutoSize        =   -1  'True
         Caption         =   "Can't see my online status"
         Height          =   195
         Left            =   -71640
         TabIndex        =   79
         Top             =   1560
         Width           =   1830
      End
      Begin VB.Label lblCanSendMeMessages 
         AutoSize        =   -1  'True
         Caption         =   "Can send me messages"
         Height          =   195
         Left            =   -74760
         TabIndex        =   78
         Top             =   1800
         Width           =   1680
      End
      Begin VB.Label lblCanSeeMyOnlineStatus 
         AutoSize        =   -1  'True
         Caption         =   "Can see my online status"
         Height          =   195
         Left            =   -74760
         TabIndex        =   77
         Top             =   1560
         Width           =   1755
      End
      Begin VB.Label lblMyBlockList 
         AutoSize        =   -1  'True
         Caption         =   "My Block List"
         Height          =   195
         Left            =   -71640
         TabIndex        =   76
         Top             =   1200
         Width           =   945
      End
      Begin VB.Label lblMyAllowList 
         AutoSize        =   -1  'True
         Caption         =   "My Allow List"
         Height          =   195
         Left            =   -74760
         TabIndex        =   75
         Top             =   1200
         Width           =   915
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000011&
         X1              =   -73290
         X2              =   -69720
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lblVisibiltyAndPrivacy 
         AutoSize        =   -1  'True
         Caption         =   "Visibility and Privacy"
         Height          =   195
         Left            =   -74760
         TabIndex        =   74
         Top             =   480
         Width           =   1425
      End
      Begin VB.Label lblMyMobilePhone 
         AutoSize        =   -1  'True
         Caption         =   "My mobile phone:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   73
         Top             =   3000
         Width           =   1245
      End
      Begin VB.Label lblMyWorkPhone 
         AutoSize        =   -1  'True
         Caption         =   "My work phone:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   72
         Top             =   2640
         Width           =   1140
      End
      Begin VB.Label lblMyHomePhone 
         AutoSize        =   -1  'True
         Caption         =   "My home phone:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   71
         Top             =   2280
         Width           =   1185
      End
      Begin VB.Label lblAreaCodeAndPhoneNumber 
         AutoSize        =   -1  'True
         Caption         =   "Area code and phone number:"
         Height          =   195
         Left            =   -73320
         TabIndex        =   70
         Top             =   2040
         Width           =   2160
      End
      Begin VB.Label lblMyCountryRegionCode 
         AutoSize        =   -1  'True
         Caption         =   "My Country/region code:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   69
         Top             =   1320
         Width           =   1755
      End
      Begin VB.Label lblTypeThePhoneNumbersThatYouWantPeopleOnYourAllowListToSee 
         Caption         =   "Type the phone numbers that you want people on your Allow List to see."
         Height          =   435
         Left            =   -74760
         TabIndex        =   68
         Top             =   840
         Width           =   5055
      End
      Begin VB.Label lblPhoneNumbers 
         AutoSize        =   -1  'True
         Caption         =   "Phone Numbers"
         Height          =   195
         Left            =   -74760
         TabIndex        =   67
         Top             =   480
         Width           =   1140
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000011&
         X1              =   -73560
         X2              =   -69720
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lblMessageHistory 
         AutoSize        =   -1  'True
         Caption         =   "Message History"
         Height          =   195
         Left            =   -74760
         TabIndex        =   66
         Top             =   2880
         Width           =   1170
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000011&
         X1              =   -73560
         X2              =   -69720
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Label lblSaveMyConversationsInThisFolder 
         AutoSize        =   -1  'True
         Caption         =   "Save my conversations in this folder"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74760
         TabIndex        =   65
         Top             =   3480
         Width           =   2535
      End
      Begin VB.Label lblPutFilesReceivedFromOthersInThisFolder 
         AutoSize        =   -1  'True
         Caption         =   "Put files received from others in this folder:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   64
         Top             =   2160
         Width           =   2970
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000011&
         X1              =   -73875
         X2              =   -69720
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label lblFileTransfer 
         AutoSize        =   -1  'True
         Caption         =   "File Transfer"
         Height          =   195
         Left            =   -74760
         TabIndex        =   63
         Top             =   1920
         Width           =   870
      End
      Begin VB.Label lblChangeTheFontAndColorOfMyInstantMessages 
         AutoSize        =   -1  'True
         Caption         =   "Change the font and color of my instant messages"
         Height          =   195
         Left            =   -74760
         TabIndex        =   62
         Top             =   720
         Width           =   3525
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000011&
         X1              =   -73725
         X2              =   -69720
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lblMessageText 
         AutoSize        =   -1  'True
         Caption         =   "Message Text"
         Height          =   195
         Left            =   -74760
         TabIndex        =   61
         Top             =   480
         Width           =   1005
      End
   End
   Begin VB.Menu mnuContact 
      Caption         =   "[Contact]"
      Visible         =   0   'False
      Begin VB.Menu mnuContact_AddToContacts 
         Caption         =   "&Add to Contacts"
      End
      Begin VB.Menu mnuContact_Move 
         Caption         =   "[Move]"
      End
      Begin VB.Menu mnuContact_Hide 
         Caption         =   "Hi&de"
      End
      Begin VB.Menu mnuContact_Ignore 
         Caption         =   "&Ignore"
      End
      Begin VB.Menu mnuContact_Delete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuContact_Seperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContact_Properties 
         Caption         =   "P&roperties"
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private AllowList As Collection, BlockList As Collection

Private Sub chkStartWithWindows_Click()
    If chkStartWithWindows.Value = vbChecked Then
        chkOpenMainWindowOnStart.Enabled = True
    Else
        chkOpenMainWindowOnStart.Enabled = False
    End If
End Sub

Private Sub cmbCountryRegionCode_Click()
    If Not cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListIndex) = 0 Then
        txtHomePhoneCode.Text = cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListIndex)
        txtWorkPhoneCode.Text = txtHomePhoneCode.Text
        txtMobilePhoneCode.Text = txtWorkPhoneCode.Text
        If cmbCountryRegionCode.ItemData(0) = 0 Then
            cmbCountryRegionCode.RemoveItem 0
        End If
    End If
End Sub

Private Sub cmdAllow_Click()
    On Error Resume Next
    
    Call UnblockContact(BlockList(lstBlock.ListIndex + 1))
End Sub

Private Sub cmdBlock_Click()
    On Error Resume Next
    
    Call BlockContact(AllowList(lstAllow.ListIndex + 1))
End Sub

Private Sub cmdBrowserBrowse_Click()
    If Not GetUserFile("Program Files (*.*)", "Custom Browser") = vbNullString Then
        txtBrowser.Text = frmMain.CommonDialog.FileName
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdChangeColor_Click()
    With frmMain.CommonDialog
        .Color = IMFontColor
        .Flags = cdlCCFullOpen Or cdlCCRGBInit
        .ShowColor
        IMFontColor = .Color
    End With
End Sub

Private Sub cmdChangeFont_Click()
    On Error Resume Next
    
    With frmMain.CommonDialog
        .Flags = cdlCFScreenFonts
        .FontName = IMFontName
        .FontBold = IMFontBold
        .FontItalic = IMFontItalic
        .ShowFont
        IMFontName = .FontName
        IMFontBold = .FontBold
        IMFontItalic = .FontItalic
    End With
End Sub

Private Sub cmdChangeMessageHistoryFolder_Click()
    On Error Resume Next
    
    frmSelectFolder.Drive1.Drive = Left$(MessageHistoryFolder, InStr(MessageHistoryFolder, "\"))
    frmSelectFolder.Dir1.Path = MessageHistoryFolder
    Set frmSelectFolder.srcTextBox = txtMessageHistoryFolder
    frmSelectFolder.Show vbModal, Me
End Sub

Private Sub cmdChangeReceivedFilesFolder_Click()
    On Error Resume Next
    
    frmSelectFolder.Drive1.Drive = Left$(ReceivedFilesFolder, InStr(ReceivedFilesFolder, "\"))
    frmSelectFolder.Dir1.Path = ReceivedFilesFolder
    Set frmSelectFolder.srcTextBox = txtReceivedFilesFolder
    frmSelectFolder.Show vbModal, Me
End Sub

Private Sub cmdChangeStatusHistoryFolder_Click()
    On Error Resume Next
    
    frmSelectFolder.Drive1.Drive = Left$(StatusHistoryFolder, InStr(StatusHistoryFolder, "\"))
    frmSelectFolder.Dir1.Path = StatusHistoryFolder
    Set frmSelectFolder.srcTextBox = txtStatusHistoryFolder
    frmSelectFolder.Show vbModal, Me
End Sub

Private Sub cmdClean_Click()
    frmContactListCleaner.Show vbModal, Me
End Sub

Private Sub cmdClearLoginCache_Click()
    On Error Resume Next
    
    DeleteSetting "Gilly Messenger", "Login Cache"
    MsgBox "Login cache cleared!", vbInformation
End Sub

Private Sub cmdEmailAppBrowse_Click()
    If Not GetUserFile("Program Files (*.*)", "Custom Email Application") = vbNullString Then
        txtEmailApp.Text = frmMain.CommonDialog.FileName
    End If
End Sub

Private Sub cmdOK_Click()
    boolUseDefaultBrowser = optBrowserDefault.Value
    strCustomBrowser = txtBrowser.Text
    
    If chkStartWithWindows.Value = vbChecked Then
        WriteRegKey "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run\Gilly Messenger", """" & App.Path & "\" & App.EXEName & ".exe"" /startup"
    Else
        DeleteRegKey "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run\Gilly Messenger"
    End If
    AlertOnContactOnline = IIf(chkAlertOnContactOnline.Value = vbChecked, True, False)
    AlertOnMessageReceived = IIf(chkAlertOnMessageReceived.Value = vbChecked, True, False)
    AlertOnEmailReceived = IIf(chkAlertOnEmailReceived.Value = vbChecked, True, False)
    BlockAlert = IIf(chkBlockAlert.Value = vbChecked, True, False)
    SoundAlerts = IIf(chkSoundAlerts.Value = vbChecked, True, False)
    BlockAlertsOnFullScrApp = IIf(chkBlockPopups.Value = vbChecked, True, False)
    
    Transparency = SldrTransparency.Value
    SaveSettingX "App Settings", "Transparency", Transparency
    Dim Window As Form
    For Each Window In Forms
        If Window.Visible Then
            SetTransparency Window, Transparency
        End If
    Next
    ReceivedFilesFolder = txtReceivedFilesFolder.Text
    SaveSettingX "App Settings", "ReceivedFiles Folder", ReceivedFilesFolder
    FTPPort = Val(txtFTPPort.Text)
    SaveSettingX "App Settings", "FTP Port", FTPPort
    
    If chkStatusHistory.Enabled And frmMain.objMSN_NS.State = NsState_SignedIn Then
        SendDisplayPic = (chkSendDisplayPic.Value = vbChecked)
        ReceiveDisplayPic = (chkReceiveDisplayPic.Value = vbChecked)
        AutoIdle = (chkAutoIdle.Value = vbChecked)
        AutoIdle_Interval = Val(txtAutoIdleInterval.Text)
        frmMain.tmrAutoIdle.Enabled = AutoIdle
        SaveStatusHistory = (chkStatusHistory.Value = vbChecked)
        StatusHistoryFolder = txtStatusHistoryFolder.Text
        SaveMessageHistory = (chkMessageHistory.Value = vbChecked)
        MessageHistoryFolder = txtMessageHistoryFolder.Text
        ShowEmoticons = (chkShowEmoticons.Value = vbChecked)
        ShowIMWindowOnMsg = (chkShowIMWindowOnMsg.Value = vbChecked)
        TypingNotification = Not (chkDisableTypingMsgNotification.Value = vbChecked)
        If Not HighlightFakeFriends = (chkHighlightFakeFriends.Value = vbChecked) Then
            HighlightFakeFriends = (chkHighlightFakeFriends.Value = vbChecked)
            Call frmMain.RefreshTreeview
        End If
        boolUseDefaultEmailApp = optEmailDefault.Value
        boolUseCustomEmailWeb = optEmailWeb.Value
        strCustomEmailWeb = txtEmailWeb.Text
        strCustomEmailApp = txtEmailApp.Text

        SaveSettingX "App Settings\" & frmMain.objMSN_NS.Login, "Send DisplayPic", SendDisplayPic
        SaveSettingX "App Settings\" & frmMain.objMSN_NS.Login, "Receive DisplayPic", ReceiveDisplayPic
        SaveSettingX "App Settings\" & frmMain.objMSN_NS.Login, "AutoIdle", AutoIdle
        SaveSettingX "App Settings\" & frmMain.objMSN_NS.Login, "AutoIdle Interval", AutoIdle_Interval
        SaveSettingX "App Settings\" & frmMain.objMSN_NS.Login, "Save StatusHistory", SaveStatusHistory
        SaveSettingX "App Settings\" & frmMain.objMSN_NS.Login, "StatusHistory Folder", StatusHistoryFolder
        SaveSettingX "App Settings\" & frmMain.objMSN_NS.Login, "Show Emoticons", ShowEmoticons
        SaveSettingX "App Settings\" & frmMain.objMSN_NS.Login, "Save MessageHistory", SaveMessageHistory
        SaveSettingX "App Settings\" & frmMain.objMSN_NS.Login, "MessageHistory Folder", MessageHistoryFolder
        SaveSettingX "App Settings\" & frmMain.objMSN_NS.Login, "Show IMWindow OnMsg", ShowIMWindowOnMsg
        SaveSettingX "App Settings\" & frmMain.objMSN_NS.Login, "Disable MsgTypingNotification", Not TypingNotification
        SaveSettingX "App Settings\" & frmMain.objMSN_NS.Login, "Highlight FakeFriends", HighlightFakeFriends
        SaveSettingX "App Settings\" & frmMain.objMSN_NS.Login, "Use DefaultEmailApp", boolUseDefaultEmailApp
        SaveSettingX "App Settings\" & frmMain.objMSN_NS.Login, "Custom EmailApp", strCustomEmailApp
        SaveSettingX "App Settings\" & frmMain.objMSN_NS.Login, "Custom EmailWeb", strCustomEmailWeb
        SaveSettingX "App Settings\" & frmMain.objMSN_NS.Login, "Use CustomEmailWeb", boolUseCustomEmailWeb
        
        If Not ((chkGTC.Value = vbChecked And frmMain.objMSN_NS.GTC = "A") Or (chkGTC.Value = vbUnchecked And frmMain.objMSN_NS.GTC = "N")) Then
            If chkGTC.Value = vbChecked Then
                frmMain.objMSN_NS.ChangeGTC "A"
            Else
                frmMain.objMSN_NS.ChangeGTC "N"
            End If
        End If
        If Not ((chkBLP.Value = vbChecked And frmMain.objMSN_NS.BLP = "BL") Or (chkBLP.Value = vbUnchecked And frmMain.objMSN_NS.BLP = "AL")) Then
            If chkBLP.Value = vbChecked Then
                frmMain.objMSN_NS.ChangeBLP "BL"
            Else
                frmMain.objMSN_NS.ChangeBLP "AL"
            End If
        End If
        
        Dim frm As Form
        For Each frm In IMWindows
            frm.RefreshMyDP
            frm.RefreshBuddyDP
        Next
        For Each frm In PendingIM
            frm.RefreshMyDP
            frm.RefreshBuddyDP
        Next
    End If
    
    SaveSettingX "App Settings", "Open MainWindow OnStart", (chkOpenMainWindowOnStart.Value = vbChecked)
    SaveSettingX "App Settings", "Alert OnContactOnline", AlertOnContactOnline
    SaveSettingX "App Settings", "Alert OnMessageReceived", AlertOnMessageReceived
    SaveSettingX "App Settings", "Alert OnEmailReceived", AlertOnEmailReceived
    SaveSettingX "App Settings", "Block Alert", BlockAlert
    SaveSettingX "App Settings", "Sound Alerts", SoundAlerts
    SaveSettingX "App Settings", "BlockAlerts OnFullScrApp", BlockAlertsOnFullScrApp
    SaveSettingX "App Settings", "Custom Browser", strCustomBrowser
    SaveSettingX "App Settings", "Use DefaultBrowser", boolUseDefaultBrowser
    Unload Me
End Sub

Private Sub cmdPopupFilter_Click()
    On Error Resume Next
    
    frmPopupFilter.Show vbModal, Me
End Sub

Private Sub cmdResetAppSettings_Click()
    On Error Resume Next
    
    DeleteSetting "Gilly Messenger", "App Settings\" & frmMain.objMSN_NS.Login
    MsgBox "Application settings reset!", vbInformation
End Sub

Private Sub cmdResetServerSettings_Click()
    On Error Resume Next
    
    DeleteSetting "Gilly Messenger", "Server Settings"
    frmMain.objMSN_NS.Server = "messenger.hotmail.com"
    frmMain.objMSN_NS.Port = 1863
    MsgBox "Server settings reset!", vbInformation
End Sub

Private Sub cmdSounds_Click()
    frmSounds.Show vbModal, Me
End Sub

Private Sub cmdViewReverseList_Click()
    Call LoadReverseList
    frmReverseList.Show vbModal, Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    LastActive = Timer
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Public Sub Form_Load()
    On Error Resume Next
    
    If ReadRegKey("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run\Gilly Messenger") = """" & App.Path & "\" & App.EXEName & ".exe"" /startup" Then
        chkStartWithWindows.Value = vbChecked
    Else
        chkOpenMainWindowOnStart.Enabled = False
    End If
    
    chkOpenMainWindowOnStart.Value = IIf(GetSettingX("App Settings", "Open MainWindow OnStart", True), vbChecked, vbUnchecked)
    chkAlertOnContactOnline.Value = IIf(AlertOnContactOnline, vbChecked, vbUnchecked)
    chkAlertOnMessageReceived.Value = IIf(AlertOnMessageReceived, vbChecked, vbUnchecked)
    chkAlertOnEmailReceived.Value = IIf(AlertOnEmailReceived, vbChecked, vbUnchecked)
    chkBlockAlert.Value = IIf(BlockAlert, vbChecked, vbUnchecked)
    chkSoundAlerts.Value = IIf(SoundAlerts, vbChecked, vbUnchecked)
    chkBlockPopups.Value = IIf(BlockAlertsOnFullScrApp, vbChecked, vbUnchecked)
    
    SldrTransparency.Value = Transparency
    txtReceivedFilesFolder.Text = ReceivedFilesFolder
    
    If GetWindowsVersion < 5 Then
        SldrTransparency.Enabled = False
        lblTransparency.ForeColor = vbGrayText
    End If
    
    txtReceivedFilesFolder.Text = ReceivedFilesFolder
    txtFTPPort.Text = FTPPort
    SetNumbered txtFTPPort.hwnd, True
    
    If boolUseDefaultBrowser Then
        optBrowserDefault.Value = True
    Else
        optBrowserCustom.Value = True
        txtBrowser.Enabled = True
        cmdBrowserBrowse.Enabled = True
    End If
    txtBrowser.Text = strCustomBrowser
    
    If frmMain.objMSN_NS.State = NsState_SignedIn Then
        cmdPopupFilter.Enabled = True
        Call LoadCountryRegionCodes
        chkStatusHistory.Value = IIf(SaveStatusHistory, vbChecked, vbUnchecked)
        txtStatusHistoryFolder.Text = StatusHistoryFolder
        chkSendDisplayPic.Value = IIf(SendDisplayPic, vbChecked, vbUnchecked)
        chkReceiveDisplayPic.Value = IIf(ReceiveDisplayPic, vbChecked, vbUnchecked)
        chkAutoIdle.Value = IIf(AutoIdle, vbChecked, vbUnchecked)
        txtAutoIdleInterval.Text = AutoIdle_Interval
        chkShowEmoticons.Value = IIf(ShowEmoticons, vbChecked, vbUnchecked)
        chkMessageHistory.Value = IIf(SaveMessageHistory, vbChecked, vbUnchecked)
        txtMessageHistoryFolder.Text = MessageHistoryFolder
        cmdChangeMessageHistoryFolder.Enabled = True
        chkShowIMWindowOnMsg.Value = IIf(ShowIMWindowOnMsg, vbChecked, vbUnchecked)
        chkDisableTypingMsgNotification.Value = IIf(Not TypingNotification, vbChecked, vbUnchecked)
        chkHighlightFakeFriends.Value = IIf(HighlightFakeFriends, vbChecked, vbUnchecked)
        
        Call LoadAllowList
        Call LoadBlockList
        If InCollection(UserProperties, "PHH") Then
            txtHomePhoneCode.Text = Split(UserProperties("PHH"))(0)
            Dim i As Integer
            For i = 1 To cmbCountryRegionCode.ListCount
                If cmbCountryRegionCode.ItemData(i) = txtHomePhoneCode.Text Then
                    cmbCountryRegionCode.ListIndex = i
                    cmbCountryRegionCode.RemoveItem 1
                    Exit For
                End If
            Next
            txtHomePhoneNumber.Text = Split(UserProperties("PHH"))(1)
        End If
        If InCollection(UserProperties, "PHW") Then
            txtWorkPhoneCode = Split(UserProperties("PHW"))(0)
            txtWorkPhoneNumber.Text = Split(UserProperties("PHW"))(1)
        End If
        If InCollection(UserProperties, "PHM") Then
            txtMobilePhoneCode.Text = Split(UserProperties("PHM"))(0)
            txtMobilePhoneNumber.Text = Split(UserProperties("PHM"))(1)
        End If
        chkBLP.Value = IIf(frmMain.objMSN_NS.BLP = "BL", vbChecked, vbUnchecked)
        chkGTC.Value = IIf(frmMain.objMSN_NS.GTC = "A", vbChecked, vbUnchecked)
        cmdClean.Enabled = True
        If boolUseDefaultEmailApp Then
            optEmailDefault.Value = True
        Else
            optEmailCustom.Value = True
            optEmailWeb.Enabled = True
            optEmailApp.Enabled = True
            If boolUseCustomEmailWeb Then
                optEmailWeb.Value = True
            Else
                optEmailApp.Value = True
            End If
        End If
        txtEmailWeb.Text = strCustomEmailWeb
        txtEmailApp.Text = strCustomEmailApp
    Else
        cmdPopupFilter.Enabled = False
        chkStatusHistory.Enabled = False
        Call DisableControl(txtStatusHistoryFolder)
        chkSendDisplayPic.Enabled = False
        chkReceiveDisplayPic.Enabled = False
        chkAutoIdle.Enabled = False
        Call DisableControl(txtAutoIdleInterval)
        cmdChangeStatusHistoryFolder.Enabled = False
        cmdChangeColor.Enabled = False
        cmdChangeFont.Enabled = False
        chkShowEmoticons.Enabled = False
        chkMessageHistory.Enabled = False
        Call DisableControl(txtMessageHistoryFolder)
        cmdChangeMessageHistoryFolder.Enabled = False
        chkShowIMWindowOnMsg.Enabled = False
        chkDisableTypingMsgNotification.Enabled = False
        chkHighlightFakeFriends.Enabled = False
        cmbCountryRegionCode.Enabled = False
        Call DisableControl(txtHomePhoneCode)
        Call DisableControl(txtHomePhoneNumber)
        Call DisableControl(txtWorkPhoneCode)
        Call DisableControl(txtWorkPhoneNumber)
        Call DisableControl(txtMobilePhoneCode)
        Call DisableControl(txtMobilePhoneNumber)
        chkBLP.Enabled = False
        chkGTC.Enabled = False
        cmdViewReverseList.Enabled = False
        cmdResetAppSettings.Enabled = False
        optEmailDefault.Enabled = False
        optEmailCustom.Enabled = False
        txtEmailWeb.Enabled = False
    End If
        
    If Not Transparency = 0 Then
        SetTransparency Me, Transparency
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    LastActive = Timer
End Sub

Private Sub lstAllow_Click()
    cmdAllow.Enabled = False
    cmdBlock.Enabled = True
End Sub

Private Sub lstAllow_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    
    lstBlock.ListIndex = -1
    If Button = vbRightButton And Not lstAllow.ListIndex = -1 Then
        mnuContact.Tag = "allow" & " " & AllowList(lstAllow.ListIndex + 1)
        mnuContact_Move.Caption = "&Move to Block List"
        If Not InList(GetContactAttr(AllowList(lstAllow.ListIndex + 1), "lists"), msnList_Forward) Then
            mnuContact_AddToContacts.Enabled = True
            mnuContact_Hide.Enabled = False
        Else
            mnuContact_AddToContacts.Enabled = False
            mnuContact_Hide.Enabled = True
            If Not InCollection(HiddenContacts, AllowList(lstAllow.ListIndex + 1)) Then
                mnuContact_Hide.Caption = "Hi&de"
            Else
                mnuContact_Hide.Caption = "Unhi&de"
            End If
        End If
        If Not InCollection(IgnoreList, AllowList(lstAllow.ListIndex + 1)) Then
            mnuContact_Ignore.Caption = "&Ignore"
        Else
            mnuContact_Ignore.Caption = "Un&ignore"
        End If
        mnuContact_Delete.Enabled = Not InList(GetContactAttr(AllowList(lstAllow.ListIndex + 1), "lists"), msnList_Reverse)
        PopupMenu mnuContact
    End If
End Sub

Private Sub lstBlock_Click()
    cmdAllow.Enabled = True
    cmdBlock.Enabled = False
End Sub

Private Sub lstBlock_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    
    lstAllow.ListIndex = -1
    If Button = vbRightButton And Not lstBlock.ListIndex = -1 Then
        mnuContact.Tag = "block" & " " & BlockList(lstBlock.ListIndex + 1)
        mnuContact_Move.Caption = "&Move to Allow List"
        If Not InList(GetContactAttr(BlockList(lstBlock.ListIndex + 1), "lists"), msnList_Forward) Then
            mnuContact_AddToContacts.Enabled = True
            mnuContact_Hide.Enabled = False
        Else
            mnuContact_AddToContacts.Enabled = False
            mnuContact_Hide.Enabled = True
            If Not InCollection(HiddenContacts, BlockList(lstBlock.ListIndex + 1)) Then
                mnuContact_Hide.Caption = "Hi&de"
            Else
                mnuContact_Hide.Caption = "Unhi&de"
            End If
        End If
        If Not InCollection(IgnoreList, BlockList(lstBlock.ListIndex + 1)) Then
            mnuContact_Ignore.Caption = "&Ignore"
        Else
            mnuContact_Ignore.Caption = "Un&ignore"
        End If
        mnuContact_Delete.Enabled = Not InList(GetContactAttr(BlockList(lstBlock.ListIndex + 1), "lists"), msnList_Reverse)
        PopupMenu mnuContact
    End If
End Sub

Private Sub mnuContact_AddToContacts_Click()
    On Error Resume Next
    
    Call AddContact(CStr(Split(mnuContact.Tag)(1)))
End Sub

Private Sub mnuContact_Delete_Click()
    On Error Resume Next
    
    If Split(mnuContact.Tag)(0) = "allow" Then
        frmMain.objMSN_NS.RemoveContact msnList_Allow, CStr(Split(mnuContact.Tag)(1))
    Else
        frmMain.objMSN_NS.RemoveContact msnList_Block, CStr(Split(mnuContact.Tag)(1))
    End If
End Sub

Private Sub mnuContact_Hide_Click()
    Select Case mnuContact_Hide.Caption
    Case "Hi&de"
        Call HideContact(CStr(Split(mnuContact.Tag)(1)))
    Case "Unhi&de"
        Call UnhideContact(CStr(Split(mnuContact.Tag)(1)))
    End Select
End Sub

Private Sub mnuContact_Ignore_Click()
    Select Case mnuContact_Ignore.Caption
    Case "&Ignore"
        Call IgnoreContact(CStr(Split(mnuContact.Tag)(1)))
    Case "&Unignore"
        Call UnignoreContact(CStr(Split(mnuContact.Tag)(1)))
    End Select
End Sub

Private Sub mnuContact_Move_Click()
    On Error Resume Next
    
    If Split(mnuContact.Tag)(0) = "allow" Then
        Call BlockContact(CStr(Split(mnuContact.Tag)(1)))
    Else
        Call UnblockContact(CStr(Split(mnuContact.Tag)(1)))
    End If
End Sub

Private Sub mnuContact_Properties_Click()
    On Error Resume Next
    
    ShowBuddyProperties Me, CStr(Split(mnuContact.Tag)(1))
End Sub

Public Sub LoadAllowList()
    lstAllow.Clear
    Set AllowList = Nothing
    Set AllowList = New Collection
    Dim i As Integer
    For i = 1 To ContactList.Count
        If InList(ContactList(i).Item("lists"), msnList_Allow) Then
            lstAllow.AddItem ContactList(i).Item("nick")
            lstAllow.ItemData(lstAllow.ListCount - 1) = i
            AllowList.Add ContactList(i).Item("email")
        End If
    Next
End Sub

Public Sub LoadBlockList()
    lstBlock.Clear
    Set BlockList = Nothing
    Set BlockList = New Collection
    Dim i As Integer
    For i = 1 To ContactList.Count
        If InList(ContactList(i).Item("lists"), msnList_Block) Then
            lstBlock.AddItem ContactList(i).Item("nick")
            lstBlock.ItemData(lstBlock.ListCount - 1) = i
            BlockList.Add ContactList(i).Item("email")
        End If
    Next
End Sub

Public Sub LoadReverseList()
    frmReverseList.lstReverse.Clear
    Set frmReverseList.ReverseList = Nothing
    Set frmReverseList.ReverseList = New Collection
    Dim i As Integer
    For i = 1 To ContactList.Count
        If InList(ContactList(i).Item("lists"), msnList_Reverse) Then
            frmReverseList.lstReverse.AddItem ContactList(i).Item("nick") & IIf(InList(ContactList(i).Item("lists"), msnList_Block), " (Blocked)", vbNullString)
            frmReverseList.lstReverse.ItemData(frmReverseList.lstReverse.ListCount - 1) = i
            frmReverseList.ReverseList.Add ContactList(i).Item("email")
        End If
    Next
End Sub

Private Sub LoadCountryRegionCodes()
    cmbCountryRegionCode.AddItem "Chose a country or region"
    
    cmbCountryRegionCode.AddItem "Afghanistan (93)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 93
    cmbCountryRegionCode.AddItem "Albania (355)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 355
    cmbCountryRegionCode.AddItem "Algeria (213)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 213
    cmbCountryRegionCode.AddItem "American Samoa (684)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 684
    cmbCountryRegionCode.AddItem "Andorra (376)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 376
    cmbCountryRegionCode.AddItem "Angola (244)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 244
    cmbCountryRegionCode.AddItem "Anguilla (1264)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 1264
    cmbCountryRegionCode.AddItem "Antarctica (672)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 672
    cmbCountryRegionCode.AddItem "Antigua and Barbuda (1268)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 1268
    cmbCountryRegionCode.AddItem "Argentina (54)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 54
    cmbCountryRegionCode.AddItem "Armenia (374)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 374
    cmbCountryRegionCode.AddItem "Aruba (297)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 297
    cmbCountryRegionCode.AddItem "Ascension (247)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 247
    cmbCountryRegionCode.AddItem "Australia (61)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 61
    cmbCountryRegionCode.AddItem "Austria (43)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 43
    cmbCountryRegionCode.AddItem "Azerbaijan (994)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 994
    cmbCountryRegionCode.AddItem "Bahamas, The (1242)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 1242
    cmbCountryRegionCode.AddItem "Bahrain (973)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 973
    cmbCountryRegionCode.AddItem "Bangladesh (880)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 880
    cmbCountryRegionCode.AddItem "Barbados (1246)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 1246
    cmbCountryRegionCode.AddItem "Belarus (375)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 375
    cmbCountryRegionCode.AddItem "Belgium (32)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 32
    cmbCountryRegionCode.AddItem "Belize (501)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 501
    cmbCountryRegionCode.AddItem "Benin (229)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 229
    cmbCountryRegionCode.AddItem "Bermuda (1441)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 1441
    cmbCountryRegionCode.AddItem "Bhutan (975)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 975
    cmbCountryRegionCode.AddItem "Bolivia (591)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 591
    cmbCountryRegionCode.AddItem "Bosnia and Herzegovina (387)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 387
    cmbCountryRegionCode.AddItem "Botswana (267)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 267
    cmbCountryRegionCode.AddItem "Brazil (55)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 55
    cmbCountryRegionCode.AddItem "Brit. Ind. Ocean Terr. (873)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 873
    cmbCountryRegionCode.AddItem "Brunei (673)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 673
    cmbCountryRegionCode.AddItem "Bulgaria (359)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 359
    cmbCountryRegionCode.AddItem "Burkina Faso (226)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 226
    cmbCountryRegionCode.AddItem "Burundi (257)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 257
    cmbCountryRegionCode.AddItem "Cambodia (855)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 855
    cmbCountryRegionCode.AddItem "Cameroon (237)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 237
    cmbCountryRegionCode.AddItem "Canada (1)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 1
    cmbCountryRegionCode.AddItem "Cape Verde (238)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 238
    cmbCountryRegionCode.AddItem "Cayman Islands (1345)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 1345
    cmbCountryRegionCode.AddItem "Central African Republic (236)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 236
    cmbCountryRegionCode.AddItem "Chad (235)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 235
    cmbCountryRegionCode.AddItem "Chile (56)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 56
    cmbCountryRegionCode.AddItem "China (86)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 86
    cmbCountryRegionCode.AddItem "Christmas Island (61)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 61
    cmbCountryRegionCode.AddItem "Cocos (keeling) ls. (61)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 61
    cmbCountryRegionCode.AddItem "Colombia (57)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 57
    cmbCountryRegionCode.AddItem "Comoros (269)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 269
    cmbCountryRegionCode.AddItem "Congo (242)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 242
    cmbCountryRegionCode.AddItem "Congo (DRC) (243)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 243
    cmbCountryRegionCode.AddItem "Cook Islands (682)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 682
    cmbCountryRegionCode.AddItem "Costa Rica (506)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 506
    cmbCountryRegionCode.AddItem "Cote d'lvoire (255)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 255
    cmbCountryRegionCode.AddItem "Croatia (385)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 385
    cmbCountryRegionCode.AddItem "Cuba (53)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 53
    cmbCountryRegionCode.AddItem "Cyprus (357)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 357
    cmbCountryRegionCode.AddItem "Czech Republic (420)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 420
    cmbCountryRegionCode.AddItem "Denmark (45)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 45
    cmbCountryRegionCode.AddItem "Diego Garcia (246)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 246
    cmbCountryRegionCode.AddItem "Djbouti (253)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 253
    cmbCountryRegionCode.AddItem "Dominica (1767)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 1767
    cmbCountryRegionCode.AddItem "Dominican Republic (1809)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 1809
    cmbCountryRegionCode.AddItem "Ecuador (593)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 593
    cmbCountryRegionCode.AddItem "Egypt (20)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 20
    cmbCountryRegionCode.AddItem "El Salvandor (503)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 503
    cmbCountryRegionCode.AddItem "Equatorial Guinea (240)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 240
    cmbCountryRegionCode.AddItem "Eritrea (291)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 291
    cmbCountryRegionCode.AddItem "Estonia (372)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 372
    cmbCountryRegionCode.AddItem "Ethiopia (251)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 251
    cmbCountryRegionCode.AddItem "Falkland ls. (Malvinas) (500)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 500
    cmbCountryRegionCode.AddItem "Faroe Islands (298)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 298
    cmbCountryRegionCode.AddItem "Fiji Islands (679)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 679
    cmbCountryRegionCode.AddItem "Finland (358)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 358
    cmbCountryRegionCode.AddItem "France (33)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 33
    cmbCountryRegionCode.AddItem "French Guiana (594)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 594
    cmbCountryRegionCode.AddItem "French Polynesia (689)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 689
    cmbCountryRegionCode.AddItem "Gabon (241)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 241
    cmbCountryRegionCode.AddItem "Gambia, The (220)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 220
    cmbCountryRegionCode.AddItem "Georgia (995)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 995
    cmbCountryRegionCode.AddItem "Germany (49)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 49
    cmbCountryRegionCode.AddItem "Ghana (233)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 233
    cmbCountryRegionCode.AddItem "Gibraltar (350)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 350
    cmbCountryRegionCode.AddItem "Greece (30)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 30
    cmbCountryRegionCode.AddItem "Greenland (299)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 299
    cmbCountryRegionCode.AddItem "Grenada (1473)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 1473
    cmbCountryRegionCode.AddItem "Guadeloupe (590)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 590
    cmbCountryRegionCode.AddItem "Guam (1671)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 1671
    cmbCountryRegionCode.AddItem "Guatemala (502)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 502
    cmbCountryRegionCode.AddItem "Guinea (244)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 244
    cmbCountryRegionCode.AddItem "Guinea-Bissau (245)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 245
    cmbCountryRegionCode.AddItem "Guyana (592)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 592
    cmbCountryRegionCode.AddItem "Haiti (509)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 509
    cmbCountryRegionCode.AddItem "Honduras (504)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 504
    cmbCountryRegionCode.AddItem "Hong Kong SAR (852)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 852
    cmbCountryRegionCode.AddItem "Hungary (36)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 36
    cmbCountryRegionCode.AddItem "Iceland (354)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 354
    cmbCountryRegionCode.AddItem "India (91)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 91
    cmbCountryRegionCode.AddItem "Indonesia (62)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 62
    cmbCountryRegionCode.AddItem "Iran (98)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 98
    cmbCountryRegionCode.AddItem "Iraq (964)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 964
    cmbCountryRegionCode.AddItem "Ireland (353)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 353
    cmbCountryRegionCode.AddItem "Israel (972)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 972
    cmbCountryRegionCode.AddItem "Italy (39)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 39
    cmbCountryRegionCode.AddItem "Jamaica (1876)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 1876
    cmbCountryRegionCode.AddItem "Japan (81)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 81
    cmbCountryRegionCode.AddItem "Jordan (962)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 962
    cmbCountryRegionCode.AddItem "Kazakhstan (7)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 7
    cmbCountryRegionCode.AddItem "Kenya (254)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 254
    cmbCountryRegionCode.AddItem "Kiribati (686)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 686
    cmbCountryRegionCode.AddItem "Korea (82)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 82
    cmbCountryRegionCode.AddItem "Korea, North (850)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 850
    cmbCountryRegionCode.AddItem "Kuwait (965)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 965
    cmbCountryRegionCode.AddItem "Kyrgyzstan (996)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 996
    cmbCountryRegionCode.AddItem "Laos (856)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 856
    cmbCountryRegionCode.AddItem "Latvia (371)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 371
    cmbCountryRegionCode.AddItem "Lebanon (961)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 961
    cmbCountryRegionCode.AddItem "Lesotho (266)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 266
    cmbCountryRegionCode.AddItem "Liberia (231)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 231
    cmbCountryRegionCode.AddItem "Libya (218)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 218
    cmbCountryRegionCode.AddItem "Liechtenstein (423)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 423
    cmbCountryRegionCode.AddItem "Lithuania (370)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 370
    cmbCountryRegionCode.AddItem "Luxembourg (352)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 352
    cmbCountryRegionCode.AddItem "Macao SAR (853)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 853
    cmbCountryRegionCode.AddItem "Macedonia, FYRO (389)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 389
    cmbCountryRegionCode.AddItem "Madagascar (261)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 261
    cmbCountryRegionCode.AddItem "Malawi (265)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 265
    cmbCountryRegionCode.AddItem "Malaysia (60)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 60
    cmbCountryRegionCode.AddItem "Maldives (960)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 960
    cmbCountryRegionCode.AddItem "Mali (223)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 223
    cmbCountryRegionCode.AddItem "Malta (356)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 356
    cmbCountryRegionCode.AddItem "Marshall Islands (692)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 692
    cmbCountryRegionCode.AddItem "Martinique (596)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 596
    cmbCountryRegionCode.AddItem "Mauritania (222)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 222
    cmbCountryRegionCode.AddItem "Mauritius (230)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 230
    cmbCountryRegionCode.AddItem "Mayotte (269)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 269
    cmbCountryRegionCode.AddItem "Mexico (52)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 52
    cmbCountryRegionCode.AddItem "Micronesia (691)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 691
    cmbCountryRegionCode.AddItem "Moldova (373)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 373
    cmbCountryRegionCode.AddItem "Monaco (377)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 377
    cmbCountryRegionCode.AddItem "Mongolia (976)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 976
    cmbCountryRegionCode.AddItem "Montserrat (1664)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 1664
    cmbCountryRegionCode.AddItem "Morocco (212)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 212
    cmbCountryRegionCode.AddItem "Mozambique (258)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 258
    cmbCountryRegionCode.AddItem "Myanmar (95)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 95
    cmbCountryRegionCode.AddItem "Namibia (264)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 264
    cmbCountryRegionCode.AddItem "Nauru (674)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 674
    cmbCountryRegionCode.AddItem "Nepal (977)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 977
    cmbCountryRegionCode.AddItem "Netherlands Antilles (599)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 599
    cmbCountryRegionCode.AddItem "Netherlands, The (31)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 31
    cmbCountryRegionCode.AddItem "New Caledonia (687)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 687
    cmbCountryRegionCode.AddItem "New Zealand (64)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 64
    cmbCountryRegionCode.AddItem "Nicaragua (505)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 505
    cmbCountryRegionCode.AddItem "Niger (227)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 227
    cmbCountryRegionCode.AddItem "Nigeria (234)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 234
    cmbCountryRegionCode.AddItem "Niue (683)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 683
    cmbCountryRegionCode.AddItem "Norfolk Island (6723)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 6723
    cmbCountryRegionCode.AddItem "Northern Mariana ls. (1670)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 1670
    cmbCountryRegionCode.AddItem "Norway (47)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 47
    cmbCountryRegionCode.AddItem "Oman (968)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 968
    cmbCountryRegionCode.AddItem "Pakistan (92)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 92
    cmbCountryRegionCode.AddItem "Palau (680)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 680
    cmbCountryRegionCode.AddItem "Panama (507)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 507
    cmbCountryRegionCode.AddItem "Papua New Guinea (675)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 675
    cmbCountryRegionCode.AddItem "Paraguay (595)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 595
    cmbCountryRegionCode.AddItem "Peru (51)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 51
    cmbCountryRegionCode.AddItem "Philippines (63)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 63
    cmbCountryRegionCode.AddItem "Poland (48)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 48
    cmbCountryRegionCode.AddItem "Portugal (351)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 351
    cmbCountryRegionCode.AddItem "Puerto Rico (1787)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 1787
    cmbCountryRegionCode.AddItem "Qatar (974)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 974
    cmbCountryRegionCode.AddItem "Reunion (262)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 262
    cmbCountryRegionCode.AddItem "Romania (40)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 40
    cmbCountryRegionCode.AddItem "Russia (7)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 7
    cmbCountryRegionCode.AddItem "Rwanda (250)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 250
    cmbCountryRegionCode.AddItem "Saint Helena (290)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 290
    cmbCountryRegionCode.AddItem "Samoa (685)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 685
    cmbCountryRegionCode.AddItem "San Marino (378)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 378
    cmbCountryRegionCode.AddItem "Sao Tome & Principe (239)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 239
    cmbCountryRegionCode.AddItem "Saudi Arabia (966)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 966
    cmbCountryRegionCode.AddItem "Senegal (221)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 221
    cmbCountryRegionCode.AddItem "Serbia and Montenegro (381)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 381
    cmbCountryRegionCode.AddItem "Seychelles (248)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 248
    cmbCountryRegionCode.AddItem "Sierra Leone (232)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 232
    cmbCountryRegionCode.AddItem "Singapore (65)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 65
    cmbCountryRegionCode.AddItem "Slovakia (421)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 421
    cmbCountryRegionCode.AddItem "Slovenia (386)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 386
    cmbCountryRegionCode.AddItem "Solomon Islands (677)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 677
    cmbCountryRegionCode.AddItem "Somalia (252)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 252
    cmbCountryRegionCode.AddItem "South Africa (27)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 27
    cmbCountryRegionCode.AddItem "Spain (34)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 34
    cmbCountryRegionCode.AddItem "Sri Lanka (94)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 94
    cmbCountryRegionCode.AddItem "St. Kitts & Nevis (1869)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 1869
    cmbCountryRegionCode.AddItem "St. Lucia (1758)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 1758
    cmbCountryRegionCode.AddItem "St. Pierre & Miquelon (508)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 508
    cmbCountryRegionCode.AddItem "St. Vincent & the Gren. (1784)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 1784
    cmbCountryRegionCode.AddItem "Sudan (249)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 249
    cmbCountryRegionCode.AddItem "Suriname (597)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 597
    cmbCountryRegionCode.AddItem "Swaziland (268)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 268
    cmbCountryRegionCode.AddItem "Sweden (46)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 46
    cmbCountryRegionCode.AddItem "Switzerland (41)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 41
    cmbCountryRegionCode.AddItem "Syria (963)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 963
    cmbCountryRegionCode.AddItem "Taiwan (886)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 886
    cmbCountryRegionCode.AddItem "Tajikistan (992)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 992
    cmbCountryRegionCode.AddItem "Tanzania (255)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 255
    cmbCountryRegionCode.AddItem "Thailand (66)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 66
    cmbCountryRegionCode.AddItem "Timor-Leste (670)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 670
    cmbCountryRegionCode.AddItem "Togo (228)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 228
    cmbCountryRegionCode.AddItem "Tokelau (690)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 690
    cmbCountryRegionCode.AddItem "Tonga (676)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 676
    cmbCountryRegionCode.AddItem "Trinidad and Tobago (1868)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 1868
    cmbCountryRegionCode.AddItem "Tunisia (216)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 216
    cmbCountryRegionCode.AddItem "Turkey (90)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 90
    cmbCountryRegionCode.AddItem "Turkmenistan (993)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 993
    cmbCountryRegionCode.AddItem "Turks and Caicos ls. (1649)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 1649
    cmbCountryRegionCode.AddItem "Tuvalu (688)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 688
    cmbCountryRegionCode.AddItem "Uganda (256)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 256
    cmbCountryRegionCode.AddItem "Ukraine (380)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 380
    cmbCountryRegionCode.AddItem "United Arab Emirates (971)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 971
    cmbCountryRegionCode.AddItem "United Kingdom (44)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 44
    cmbCountryRegionCode.AddItem "United States (1)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 1
    cmbCountryRegionCode.AddItem "Uruguay (598)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 598
    cmbCountryRegionCode.AddItem "Uzbekistan (998)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 998
    cmbCountryRegionCode.AddItem "Vanuatu (678)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 678
    cmbCountryRegionCode.AddItem "Vatican City (3906)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 3906
    cmbCountryRegionCode.AddItem "Venezuela (58)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 58
    cmbCountryRegionCode.AddItem "Vietnam (84)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 84
    cmbCountryRegionCode.AddItem "Virgin Islands (1340)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 1340
    cmbCountryRegionCode.AddItem "Virgin Islands, British (1284)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 1284
    cmbCountryRegionCode.AddItem "Wallis and Futuna (681)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 681
    cmbCountryRegionCode.AddItem "Yemen (967)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 967
    cmbCountryRegionCode.AddItem "Zambia (260)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 260
    cmbCountryRegionCode.AddItem "Zimbabwe (263)": cmbCountryRegionCode.ItemData(cmbCountryRegionCode.ListCount - 1) = 263
    
    cmbCountryRegionCode.ListIndex = 0
End Sub

Private Sub optBrowserCustom_Click()
    txtBrowser.Enabled = True
    cmdBrowserBrowse.Enabled = True
End Sub

Private Sub optBrowserDefault_Click()
    txtBrowser.Enabled = False
    cmdBrowserBrowse.Enabled = False
End Sub

Private Sub optEmailApp_Click()
    txtEmailWeb.Enabled = False
    txtEmailApp.Enabled = True
    cmdEmailAppBrowse.Enabled = True
End Sub

Private Sub optEmailCustom_Click()
    optEmailWeb.Enabled = True
    txtEmailWeb.Enabled = True
    optEmailApp.Enabled = True
    txtEmailApp.Enabled = True
    cmdEmailAppBrowse.Enabled = True
End Sub

Private Sub optEmailDefault_Click()
    optEmailWeb.Enabled = False
    txtEmailWeb.Enabled = False
    optEmailApp.Enabled = False
    txtEmailApp.Enabled = False
    cmdEmailAppBrowse.Enabled = False
End Sub

Private Sub optEmailWeb_Click()
    txtEmailWeb.Enabled = True
    txtEmailApp.Enabled = False
    cmdEmailAppBrowse.Enabled = False
End Sub

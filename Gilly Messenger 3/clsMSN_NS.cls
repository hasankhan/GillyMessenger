VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMSN_NS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'SSL
    Private SecureSession As clsSSL
    
'Sockets
    Private WithEvents wskNS As Winsock
Attribute wskNS.VB_VarHelpID = -1
    Private SocketClosed As Boolean

'Enumurations
    Public Enum NsStateConstant
        NsState_Disconnected
        NsState_Connecting
        NsState_Connected
        NsState_Negotiating
        NsState_SwitchingServer
        NsState_Authenticating
        NsState_SignedIn
    End Enum

    Public Enum msnStatusConstant
        msnStatus_Online
        msnStatus_Busy
        msnStatus_BeRightBack
        msnStatus_Away
        msnStatus_OneThePhone
        msnStatus_OutToLunch
        msnStatus_Idle
        msnStatus_Offline
        msnStatus_Unknown
    End Enum
    
    Public Enum MsnContactListConstant
        msnList_Forward = 1
        msnList_Allow = 2
        msnList_Block = 4
        msnList_Reverse = 8
    End Enum
    
'Properties
    'Server
    Private NS_Server As String
    Private NS_Port As Integer
    Private NS_Version As String
    'Account
    Private NS_Login As String
    Private NS_Password As String
    Private NS_Nick As String
    'Profile
    Private NS_LoginTime As Long
    Private NS_EmailEnabled As Boolean
    Private NS_MemberIdHigh As Double
    Private NS_MemberIdLow As Double
    Private NS_lang_preference As Integer
    Private NS_preferredEmail As String
    Private NS_country As String * 2
    Private NS_PostalCode As Integer
    Private NS_Gender As String * 1
    Private NS_Kid As Boolean
    Private NS_Age As Integer
    Private NS_BDayPre As Integer
    Private NS_Birthday As Integer
    Private NS_Wallet As Boolean
    Private NS_Flags As Integer
    Private NS_sid As Integer
    Private NS_kv As Integer
    Private NS_MSPAuth As String
    Private NS_ClientIP As String
    Private NS_ClientPort As Long
    Private NS_GTC As String * 1
    Private NS_BLP As String * 2
    
    'Email
    Private NS_InboxUnread As Integer
    Private NS_FoldersUnread As Integer
    Private NS_InboxURL As String
    Private NS_FoldersURL As String
    Private NS_PostURL As String
    
'Events
    Public Event NsError(Error As String)
    Public Event SocketError(Description As String)
    Public Event StateChanged()
    Public Event StatusChanged(Status As Integer)
    Public Event NickChanged()
    Public Event MailBoxStatus(InboxUnreadMsgs As Integer, FoldersUnreadMsgs As Integer)
    Public Event EmailNotification(FromName As String, FromEmail As String, MessageURL As String, PostURL As String, Subject As String, DestFolder As String, ID As Integer)
    Public Event MailBoxNotification(SrcFolder As String, DestFolder As String, Messages As Integer)
    Public Event ContactListCacheID(CacheID As Integer)
    Public Event PropertyReceived(Property As String, Value As String)
    Public Event GroupReceived(GroupID As Integer, GroupName As String)
    Public Event ContactReceived(Email As String, Nick As String, Lists As Integer, Groups As Collection)
    Public Event ContactPropertyReceived(Email As String, Property As String, Value As String)
    Public Event GTCChanged()
    Public Event BLPChanged()
    Public Event ContactAdded(Email As String, Nick As String, list As Integer, GroupID As Integer)
    Public Event ContactRemoved(Email As String, list As Integer, GroupID As Integer)
    Public Event ContactRenamed(Email As String, Nick As String)
    Public Event ChatRequest(Email As String, Nick As String, SessionID As Double, AuthCode As String, ServerIP As String, ServerPort As Integer)
    Public Event UrlReceived(rru As String, URL As String, ID As Integer)
    Public Event ContactInitialStatus(Email As String, Nick As String, Status As Integer)
    Public Event ContactStatusChanged(Email As String, Nick As String, Status As Integer)
    Public Event ContactOffline(Email As String)
    Public Event GroupAdded(GroupID As Integer, GroupName As String)
    Public Event GroupRemoved(GroupID As Integer)
    Public Event GroupRenamed(GroupID, GroupName As String)
    Public Event ListRetrievalComplete()
    Public Event SwitchboardReceived(IP As String, Port As Integer, AuthCode As String)
    Public Event SignInProgress(Percent As Integer)
    
'Server Vars
    Private NS_Data As String
    Private NS_TrialID As Double
    Private NS_State As Integer
    
'General Vars
    Private NS_ContactCount As Integer
    Private NS_LastContact As String
    Private NS_LoginStamp As Date
    Private SSL_Socket As Winsock
    Private xmlTag As String
    
Public Property Let Server(Data As String)
    NS_Server = Data
End Property

Public Property Get Server() As String
    Server = NS_Server
End Property

Public Property Let Port(Data As Integer)
    NS_Port = Data
End Property

Public Property Get Port() As Integer
    Port = NS_Port
End Property

Public Property Get State() As Integer
    State = NS_State
End Property

Public Property Get Login() As String
    Login = NS_Login
End Property

Public Property Let Login(ByVal Data As String)
    If NS_State = NsState_Disconnected Then
        NS_Login = Data
    End If
End Property

Public Property Get Password() As String
    Password = NS_Password
End Property

Public Property Let Password(ByVal Data As String)
    If NS_State = NsState_Disconnected Then
        If Len(Data) > 16 Then
                NS_Password = Left$(Data, 16)
        Else
                NS_Password = Data
        End If
    End If
End Property

Public Property Get Version() As String
    Version = NS_Version
End Property

Public Property Let Version(Data As String)
    If NS_State = NsState_Disconnected Then
        NS_Version = Data
    End If
End Property

Public Property Get Nick() As String
    Nick = NS_Nick
End Property

Public Property Get LoginTime() As Long
    LoginTime = NS_LoginTime
End Property

Public Property Get EmailEnabled() As Boolean
    EmailEnabled = NS_EmailEnabled
End Property

Public Property Get MemberIdHigh() As Long
    MemberIdHigh = NS_MemberIdHigh
End Property

Public Property Get MemberIdLow() As Long
    MemberIdLow = NS_MemberIdLow
End Property
    
Public Property Get lang_preference() As Integer
    lang_preference = NS_lang_preference
End Property

Public Property Get preferredEmail() As String
    preferredEmail = NS_preferredEmail
End Property

Public Property Get country() As String
    country = NS_country
End Property
    
Public Property Get PostalCode() As Integer
    PostalCode = NS_PostalCode
End Property

Public Property Get Gender() As String
    Gender = NS_Gender
End Property

Public Property Get Kid() As Boolean
    Kid = NS_Kid
End Property
    
Public Property Get Age() As Integer
    Age = NS_Age
End Property
    
Public Property Get BDayPre() As Integer
    BDayPre = NS_BDayPre
End Property

Public Property Get Birthday() As Integer
    Birthday = NS_Birthday
End Property
    
Public Property Get Wallet() As Boolean
    Wallet = NS_Wallet
End Property
    
Public Property Get Flags() As Integer
    Flags = NS_Flags
End Property

Public Property Get sid() As Integer
    sid = NS_sid
End Property

Public Property Get kv() As Integer
    kv = NS_kv
End Property

Public Property Get MSPAuth() As String
    MSPAuth = NS_MSPAuth
End Property

Public Property Get ClientIP() As String
    ClientIP = NS_ClientIP
End Property

Public Property Get ClientPort() As Integer
    ClientPort = NS_ClientPort
End Property
    
Public Property Get GTC() As String
    GTC = NS_GTC
End Property

Public Property Get BLP() As String
    BLP = NS_BLP
End Property
    
Public Property Get InboxURL() As String
    InboxURL = NS_InboxURL
End Property

Public Property Get FoldersURL() As String
    FoldersURL = NS_FoldersURL
End Property

Public Property Get PostURL() As String
    PostURL = NS_PostURL
End Property

Public Property Get sl() As Double
    If NS_State = NsState_SignedIn Then
        sl = DateDiff("s", NS_LoginStamp, Now())
    Else
        sl = 0
    End If
End Property

Public Property Get TrialID() As Double
    TrialID = NS_TrialID
End Property

Public Property Let NsSocket(Data As Winsock)
    Set wskNS = Data
End Property

Public Property Let SslSocket(Data As Winsock)
    Set SSL_Socket = Data
End Property

Private Sub Class_Initialize()
    NS_State = NsState_Disconnected
    NS_Server = "messenger.hotmail.com"
    NS_Port = 1863
    NS_Version = "6.1.0207"
End Sub

Private Sub wskNs_Close()
    wskNS.Close
    Call DisconnectClient
End Sub

Private Sub wskNs_Connect()
    ChangeNsState NsState_Connected
    RaiseEvent SignInProgress(20)
    NS_TrialID = 1
    NS_Send "VER " & NS_TrialID & " MSNP8 CVRO"
End Sub

Private Sub wskNs_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
    
    Dim Data As String, Commands() As String, command As String
    
    wskNS.GetData Data
    NS_Data = NS_Data & Data
    
    If xmlTag <> vbNullString Then
        If RemoveTag(xmlTag, NS_Data) Then
            xmlTag = vbNullString
        Else
            Exit Sub
        End If
    End If
    
    Commands = Split(NS_Data, vbCrLf)
        
    If Right$(Data, 2) = vbCrLf Then
        NS_Data = vbNullString
    Else
        NS_Data = Commands(UBound(Commands))
        If UBound(Commands) = 0 Then
            Exit Sub
        End If
    End If
    
    ReDim Preserve Commands(UBound(Commands) - 1)
    
    Dim CmdPtr As Integer, i As Integer, j As Integer, Temp As String, CmdParams() As String, aryTemp() As String, tmpCollection As Collection
    
    For CmdPtr = 0 To UBound(Commands)
    
        command = Commands(CmdPtr)
        
        Debug.Print "<- "; command
        
        If xmlTag <> vbNullString Then
            If RemoveTag(xmlTag, command) Then
                xmlTag = vbNullString
            End If
            command = vbNullString
        End If
        
        If Not command = vbNullString Then
        
            CmdParams = Split(command)
        
            Select Case CmdParams(0)
            Case "<NOTIFICATION"
                'if the notification crap is in one line we should remove it
                'put the remaining text in the commands array and move the
                'array pointer one step back so it can be re-processed
                xmlTag = "NOTIFICATION"
                If RemoveTag(xmlTag, command) Then
                    Commands(CmdPtr) = command
                    xmlTag = vbNullString
                    CmdPtr = CmdPtr - 1
                End If
            Case "VER"
                'VER 1 MSNP8
                ChangeNsState NsState_Negotiating
                RaiseEvent SignInProgress(40)
                NS_Send "CVR " & NS_TrialID & " 0x0413 winnt 5.2 i386 MSNMSGR " & NS_Version & " MSMSGS " & NS_Login
            
            Case "CVR"
                'CVR 2 6.0.0602 6.0.0602 5.0.0527 http://download.microsoft.com/download/4/d/e/4defe3d0-b2e3-4de2-ab23-8bd74be402ea/SETUPNT.EXE http://messenger.msn.com/nl
                RaiseEvent SignInProgress(60)
                NS_Send "USR " & NS_TrialID & " TWN I " & NS_Login
            
            Case "USR"
                Select Case CmdParams(2)
            
                Case "TWN"
                    'USR 3 TWN S lc=1033,id=507,tw=40,fs=1,ru=http%3A%2F%2Fmessenger%2Emsn%2Ecom,ct=1074818928,kpp=1,kv=5,ver=2.1.0173.1,tpf=b113ea206f97a760083b067ea71a892e
                    'SSL Authentication
                    ChangeNsState NsState_Authenticating
                    RaiseEvent SignInProgress(80)

                    Temp = GetNsTicket(CmdParams(4))

                    If Temp = "False" Then
                        Call wskNs_Close
                        RaiseEvent NsError("911")
                    ElseIf Temp = "Error" Then
                        Call wskNs_Close
                        Exit Sub
                    Else
                        NS_Send "USR " & NS_TrialID & " TWN S " & Temp
                    End If

                Case "OK"
                    'USR 4 OK ayeshali_86@hotmail.com Ayesha 1 0
                    RaiseEvent SignInProgress(100)
                    NS_Nick = MSN_Decode(CmdParams(4))
                End Select
                    
            Case "XFR"
                Select Case CmdParams(2)
            
                Case "NS"
                    wskNS.Close
                    NS_TrialID = 1
                    Temp = CmdParams(3)
                    NS_Server = Split(Temp, ":")(0)
                    NS_Port = Split(Temp, ":")(1)
                    ChangeNsState NsState_SwitchingServer
                    wskNS.Connect NS_Server, NS_Port
                Case "SB"
                    'Switch board server received
                    RaiseEvent SwitchboardReceived(CStr(Split(CmdParams(3), ":")(0)), Val(Split(CmdParams(3), ":")(1)), CmdParams(5))
                End Select
            
            Case "MSG"
                'MSG Hotmail Hotmail 461
                
                i = Val(CmdParams(3))
                Temp = vbNullString
                    
                For j = CmdPtr + 1 To UBound(Commands)
                    Temp = Temp & Commands(j) & vbCrLf
                    If Len(Temp) = i Then
                        ProcessNsMsg command & vbCrLf & Temp
                        Exit For
                    ElseIf Len(Temp) > i Then
                        Commands(j) = Mid$(Temp, i + 1, Len(Temp) - i - 2)
                        Temp = Left$(Temp, i)
                        ProcessNsMsg command & vbCrLf & Left$(Temp, i)
                        j = j - 2
                        Exit For
                    End If
                Next
                    
                If Len(Temp) < i Then
                    If Not NS_Data = vbNullString Then
                        Temp = Temp & NS_Data
                        NS_Data = vbNullString
                    End If
                    
                    If Len(Temp) = i Then
                        ProcessNsMsg command & vbCrLf & Temp
                    ElseIf Len(Temp) > i Then
                        NS_Data = Right$(Temp, Len(Temp) - i)
                        ProcessNsMsg command & vbCrLf & Left$(Temp, i)
                    Else
                        NS_Data = command & vbCrLf & Temp & NS_Data
                    End If
                End If
                    
                CmdPtr = j
            
            Case "SYN"
                'SYN 5 286 7 1
                If UBound(CmdParams) = 4 Then
                    NS_ContactCount = Val(CmdParams(3))
                End If
                RaiseEvent ContactListCacheID(Val(CmdParams(2)))
                If NS_ContactCount = 0 Then
                    RaiseEvent ListRetrievalComplete
                End If
            
            Case "GTC"
                'GTC A
                If UBound(CmdParams) = 1 Then
                    NS_GTC = CmdParams(1)
                Else
                    NS_GTC = CmdParams(3)
                End If
                RaiseEvent GTCChanged
        
            Case "BLP"
                'BLP BL
                If UBound(CmdParams) = 1 Then
                    NS_BLP = CmdParams(1)
                Else
                    NS_BLP = CmdParams(3)
                End If
                RaiseEvent BLPChanged
        
            Case "PRP"
                If UBound(CmdParams) = 2 Then
                    'PRP PHH 01%20234
                    RaiseEvent PropertyReceived(CmdParams(1), URL_Decode(CmdParams(2)))
                Else
                    'PRP 55 12183 PHH 555-1234
                End If
                
            Case "LSG"
                'LSG 0 Other%20Contacts 0
                RaiseEvent GroupReceived(Val(CmdParams(1)), URL_Decode(CmdParams(2)))
            
            Case "LST"
                'LST me_fatima87@hotmail.com tHe%20oNlY%20tHiNg%20tHAt%20lOOkS%20gOoD%20oN%20Me%20iS%20YOU%20;) 12
                NS_ContactCount = NS_ContactCount - 1
                NS_LastContact = CmdParams(1)
                Set tmpCollection = New Collection
            
                If UBound(CmdParams) > 3 Then
                    aryTemp = Split(CmdParams(4), ",")
                    For i = LBound(aryTemp) To UBound(aryTemp)
                        tmpCollection.Add aryTemp(i), "GRP " & aryTemp(i)
                    Next
                End If
            
                RaiseEvent ContactReceived(CmdParams(1), MSN_Decode(CmdParams(2)), Val(CmdParams(3)), tmpCollection)
            
                Erase aryTemp
                Set tmpCollection = Nothing
            
                If NS_ContactCount = 0 Then
                    RaiseEvent ListRetrievalComplete
                End If
            
            Case "BPR"
                If UBound(CmdParams) = 4 Then
                    RaiseEvent ContactPropertyReceived(CmdParams(2), CmdParams(3), URL_Decode(CmdParams(4)))
                    RaiseEvent ContactListCacheID(Val(CmdParams(1)))
                Else
                    'BPR PHH 01%20234
                    RaiseEvent ContactPropertyReceived(NS_LastContact, CmdParams(1), URL_Decode(CmdParams(2)))
                End If
            
            Case "ILN"
                'ILN 6 NLN grownupkid_87@hotmail.com grownupkid_87@hotmail.com 0
                RaiseEvent ContactInitialStatus(CmdParams(3), MSN_Decode(CmdParams(4)), Val(StatusCode(CmdParams(2))))
        
            Case "CHL"
                wskNS.SendData "QRY " & NS_TrialID & " msmsgs@msnmsgr.com 32" & vbCrLf & MD5Encrypt(CmdParams(2) & "Q1P7W2E4J9R8U3S5")
        
            Case "NLN"
                RaiseEvent ContactStatusChanged(CmdParams(2), MSN_Decode(CmdParams(3)), Val(StatusCode(CmdParams(1))))
        
            Case "FLN"
                RaiseEvent ContactOffline(CmdParams(1))
            
            Case "RNG"
                RaiseEvent ChatRequest(CmdParams(5), MSN_Decode(CmdParams(6)), Val(CmdParams(1)), CmdParams(4), CStr(Split(CmdParams(2), ":")(0)), Val(Split(CmdParams(2), ":")(1)))
        
            Case "REA"
                If CmdParams(3) = NS_Login Then
                    NS_Nick = MSN_Decode(CmdParams(4))
                    RaiseEvent NickChanged
                Else
                    RaiseEvent ContactRenamed(CmdParams(3), MSN_Decode(CmdParams(4)))
                    RaiseEvent ContactListCacheID(Val(CmdParams(2)))
                End If
            
            Case "REM"
                If UBound(CmdParams) = 4 Then
                    RaiseEvent ContactRemoved(CmdParams(4), Val(ListCode(CmdParams(2))), -1)
                Else
                    RaiseEvent ContactRemoved(CmdParams(4), Val(ListCode(CmdParams(2))), Val(CmdParams(5)))
                End If
            
                RaiseEvent ContactListCacheID(Val(CmdParams(3)))
        
            Case "ADD"
                If ListCode(CmdParams(2)) = msnList_Forward Then
                    If UBound(CmdParams) = 6 Then
                        RaiseEvent ContactAdded(CmdParams(4), MSN_Decode(CmdParams(5)), Val(ListCode(CmdParams(2))), Val(CmdParams(6)))
                    Else
                        RaiseEvent ContactAdded(CmdParams(4), MSN_Decode(CmdParams(5)), Val(ListCode(CmdParams(2))), 0)
                    End If
                Else
                    RaiseEvent ContactAdded(CmdParams(4), MSN_Decode(CmdParams(5)), Val(ListCode(CmdParams(2))), -1)
                End If
                RaiseEvent ContactListCacheID(Val(CmdParams(3)))
        
            Case "ADG"
                RaiseEvent GroupAdded(Val(CmdParams(4)), URL_Decode(CmdParams(3)))
                RaiseEvent ContactListCacheID(Val(CmdParams(2)))
            
            Case "RMG"
                RaiseEvent GroupRemoved(Val(CmdParams(3)))
                RaiseEvent ContactListCacheID(Val(CmdParams(2)))
            
            Case "REG"
                RaiseEvent GroupRenamed(Val(CmdParams(3)), URL_Decode(CmdParams(4)))
                RaiseEvent ContactListCacheID(Val(CmdParams(2)))
            
            Case "CHG"
                'CHG 6 NLN 0
                RaiseEvent StatusChanged(StatusCode(CmdParams(2)))
        
            Case "URL"
                RaiseEvent UrlReceived(CmdParams(2), CmdParams(3), Val(CmdParams(4)))
        
            Case "OUT"
                RaiseEvent NsError(command)
            
            Case Else
                If IsNumeric(CmdParams(0)) Then
                    RaiseEvent NsError(command)
                End If
            End Select
        End If
    Next
End Sub

Private Function RemoveTag(xmlTag As String, ByRef Data As String) As Boolean
    Dim pos As Integer
    Dim test As String
    test = "</" + xmlTag + ">"
    pos = InStr(Data, test)
    If pos > 0 Then
        Data = Mid$(Data, pos + Len(test))
        RemoveTag = True
    Else
        RemoveTag = False
    End If
End Function

Public Sub Connect()
    On Error GoTo Handler
    ChangeNsState NsState_Connecting
    NS_Data = vbNullString
    wskNS.Close
    wskNS.Connect NS_Server, NS_Port
    Exit Sub
Handler:
    Call DisconnectClient
End Sub

Private Sub wskNS_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    SetSocketError Description
End Sub

Private Function GetNsTicket(strChallenge As String) As String
    On Error Resume Next
    
    Dim Temp As String, SSL_Data As String
    
    Temp = Split(NS_Login, "@")(1)
    If Split(Temp, ".")(0) = "hotmail" Then
        Temp = "loginnet.passport.com"
    Else
        Select Case Temp
        Case "msn.com", "compaq.net", "webtv.net"
            Temp = "msnialogin.passport.com"
        Case Else
            Temp = "login.passport.com"
        End Select
    End If
    
    Set SecureSession = New clsSSL
    SecureSession.Socket = SSL_Socket
    
    SSL_Data = SecureSession.SSL_Get(Temp, 443, "login2.srf", "Authorization: Passport1.4 OrgVerb=GET,OrgURL=http%3A%2F%2Fmessenger%2Emsn%2Ecom,sign-in=" & URL_Encode(NS_Login) & ",pwd=" & URL_Encode(Password) & "," & strChallenge & _
                "User-Agent: MSMSGS" & vbCrLf & _
                "Host: loginnet.passport.com" & vbCrLf & _
                "Connection: Keep-Alive" & vbCrLf & _
                "Cache-Control: no-cache")
                
    If SSL_Data = vbNullString Then
        Call SetSocketError(vbNullString)
        Set SecureSession = Nothing
        Exit Function
    End If
    
    If InStr(SSL_Data, "HTTP/1.1 200 OK") > 0 Then
        If InStr(SSL_Data, "Authentication-Info") > 0 Then
            Temp = Split(Filter(Split(SSL_Data, vbCrLf), "Authentication-Info")(0), ": ")(1)
        Else
            Temp = Split(Filter(Split(SSL_Data, vbCrLf), "WWW-Authenticate")(0), ": ")(1)
        End If
        Temp = Filter(Split(Temp, ","), "from-PP=")(0)
        Temp = Filter(Split(Temp, "'"), "t=")(0)
        GetNsTicket = Temp
    Else
        GetNsTicket = "False"
    End If
    
    Set SecureSession = Nothing
End Function

Private Sub ChangeNsState(State As Integer)
    NS_State = State
    RaiseEvent StateChanged
End Sub

Private Sub ProcessNsMsg(Message As String)
    On Error Resume Next
    
    Dim MsgLines() As String, MsgType As String, i As Integer, MsgAttr() As String
    
    MsgLines = Split(Message, vbCrLf)
    
    MsgType = Split(Split(MsgLines(2))(1), ";")(0)
    
    Select Case MsgType
    Case "text/x-msmsgsprofile"
        'MSG Hotmail Hotmail 363
        'MIME-Version: 1.0
        'Content-Type: text/x-msmsgsprofile; charset=UTF-8
        'LoginTime: 1016941010
        'EmailEnabled: 1
        'MemberIdHigh: 41922
        'MemberIdLow: -619232012
        'lang_preference: 1033
        'preferredEmail: example@ passport.com
        'country: US
        'PostalCode: 12345
        'Gender: M
        'Kid: 0
        'Age:
        'sid: 507
        'kv: 2
        'MSPAuth: 2AAAAAAAADMoV8ORoz64BVwmjtksIg!kmR!Rj5tBBqEaW9hc4YnPHSOQ$$

        For i = 3 To UBound(MsgLines) - 2
            
            MsgAttr() = Split(MsgLines(i), ": ")
    
            Select Case MsgAttr(0)
            Case "LoginTime"
                NS_LoginTime = Val(MsgAttr(1))
            Case "EmailEnabled"
                NS_EmailEnabled = CBool(Val(MsgAttr(1)))
            Case "MemberIdHigh"
                NS_MemberIdHigh = Val(MsgAttr(1))
            Case "MemberIdLow"
                NS_MemberIdLow = Val(MsgAttr(1))
            Case "lang_preference"
                NS_lang_preference = Val(MsgAttr(1))
            Case "preferredEmail"
                NS_preferredEmail = MsgAttr(1)
            Case "country"
                NS_country = MsgAttr(1)
            Case "PostalCode"
                NS_PostalCode = Val(MsgAttr(1))
            Case "Gender"
                NS_Gender = MsgAttr(1)
            Case "Kid"
                NS_Kid = Val(MsgAttr(1))
            Case "Age"
                NS_Age = Val(MsgAttr(1))
            Case "BdayPre"
                NS_BDayPre = Val(MsgAttr(1))
            Case "Birthday"
                NS_Birthday = Val(MsgAttr(1))
            Case "Wallet"
                NS_Wallet = CBool(Val(MsgAttr(1)))
            Case "Flags"
                NS_Flags = Val(MsgAttr(1))
            Case "sid"
                NS_sid = Val(MsgAttr(1))
            Case "kv"
                NS_kv = Val(MsgAttr(1))
            Case "MSPAuth"
                NS_MSPAuth = MsgAttr(1)
            Case "ClientIP"
                NS_ClientIP = MsgAttr(1)
            Case "ClientPort"
                NS_ClientPort = Val(MsgAttr(1))
            End Select
                
        Next
        If Not (NS_State = NsState_SignedIn) Then
            NS_LoginStamp = Now()
            Call ChangeNsState(NsState_SignedIn)
        End If
            
    Case "text/x-msmsgsinitialemailnotification"
        'MSG Hotmail Hotmail 223
        'MIME-Version: 1.0
        'Content-Type: text/x-msmsgsinitialemailnotification; charset=UTF-8
        '
        'Inbox-Unread: 21
        'Folders-Unread: 33
        'Inbox-URL: /cgi-bin/HoTMaiL
        'Folders-URL: /cgi-bin/folders
        'Post-URL: http://www.hotmail.com
            
        RaiseEvent MailBoxStatus(Val(Split(Filter(MsgLines, "Inbox-Unread")(0))(1)), Val(Split(Filter(MsgLines, "Folders-Unread")(0))(1)))
        
    Case "text/x-msmsgsemailnotification"
        'MSG Hotmail Hotmail 355
        'MIME-Version: 1.0
        'Content-Type: text/x-msmsgsemailnotification; charset=UTF-8
        '
        'From: Mike Mintz
        'Message-URL: /cgi-bin/getmsg?msg=MSG1050451140.21?&start=2310&len=2059&curmbox=ACTIVE
        'Post-URL: https://loginnet.passport.com/ppsecure/?md5auth.srf?lc=1038
        'Subject: =?"us-ascii"?Q?newsubject?=
        'Dest-Folder: ACTIVE
        'From -Addr: example@ passport.com
        'id: 2

        RaiseEvent EmailNotification(CStr(Split(Filter(MsgLines, "From:")(0))(1)), CStr(Split(Filter(MsgLines, "From-Addr:")(0))(1)), CStr(Split(Filter(MsgLines, "Message-URL:")(0))(1)), CStr(Split(Filter(MsgLines, "Post-URL:")(0))(1)), CStr(Split(Filter(MsgLines, "Subject:")(0))(1)), CStr(Split(Filter(MsgLines, "Dest-Folder:")(0))(1)), Val(Split(Filter(MsgLines, "id:")(0))(1)))
    
    Case "text/x-msmsgsactivemailnotification"
        'MSG Hotmail Hotmail 145
        'MIME-Version: 1.0
        'Content-Type: text/x-msmsgsactivemailnotification; charset=UTF-8
        '
        'Src-Folder: ACTIVE
        'Dest-Folder: trAsH
        'Message-Delta: 2

        RaiseEvent MailBoxNotification(CStr(Split(Filter(MsgLines, "Src-Folder:")(0))(1)), CStr(Split(Filter(MsgLines, "Dest-Folder:")(0))(1)), Val(Split(Filter(MsgLines, "Message-Delta:")(0))(1)))
    End Select
End Sub

Private Sub NS_Send(Data As String)
    On Error Resume Next
    
    Debug.Print "-> "; Data
    Data = Data & vbCrLf
    Dim i As Integer
    For i = 1 To Len(Data) Step 1500
        wskNS.SendData Mid$(Data, i, 1500)
        DoEvents
    Next
    NS_TrialID = NS_TrialID + 1
End Sub

Private Sub SetSocketError(Description As String)
    SocketClosed = True
    NS_State = NsState_Disconnected
    RaiseEvent SocketError(Description)
    RaiseEvent StateChanged
End Sub

Public Sub ChangeStatus(Status As Integer)
    NS_Send "CHG " & NS_TrialID & " " & StatusCode(Status)
End Sub

Public Sub ChangeNick(Nick As String)
    If Not Nick = vbNullString Then
        NS_Send "REA " & NS_TrialID & " " & NS_Login & " " & MSN_Encode(Nick)
    End If
End Sub

Public Sub RequestContactList(CacheID As Integer)
    NS_Send "SYN " & NS_TrialID & " " & CacheID
End Sub

Public Sub AddContact(ByVal list As Integer, ByVal Email As String, Optional ByVal Nick As String, Optional Group As Integer = -1)
    Nick = IIf(Nick = vbNullString, Email, Nick)
    If Group = -1 Then
        NS_Send "ADD " & NS_TrialID & " " & ListCode(list) & " " & Email & " " & IIf(list = msnList_Forward, MSN_Encode(Nick), Email)
    Else
        NS_Send "ADD " & NS_TrialID & " " & ListCode(list) & " " & Email & " " & IIf(list = msnList_Forward, MSN_Encode(Nick), Email) & " " & Group
    End If
End Sub

Public Sub RemoveContact(list As Integer, Email As String, Optional Group As Integer = -1)
    If Group = -1 Then
        NS_Send "REM " & NS_TrialID & " " & ListCode(list) & " " & Email
    Else
        NS_Send "REM " & NS_TrialID & " " & ListCode(list) & " " & Email & " " & Group
    End If
End Sub

Public Sub RenameContact(Email As String, Nick As String)
    NS_Send "REA " & NS_TrialID & " " & Email & " " & MSN_Encode(IIf(Nick = vbNullString, Email, Nick))
End Sub

Public Sub RequestURL(Action As String)
    NS_Send "URL " & NS_TrialID & " " & Action
End Sub

Public Sub AddGroup(GroupName As String)
    NS_Send "ADG " & NS_TrialID & " " & URL_Encode(GroupName) & " 0"
End Sub

Public Sub RemoveGroup(GroupID As Integer)
    NS_Send "RMG " & NS_TrialID & " " & GroupID
End Sub

Public Sub RenameGroup(GroupID As Integer, NewGroupName As String)
    NS_Send "REG " & NS_TrialID & " " & GroupID & " " & URL_Encode(NewGroupName)
End Sub

Public Sub Disconnect()
    If NS_State = NsState_SignedIn Then
        NS_Send "OUT"
        DoEvents
    End If
    Call DisconnectClient
End Sub

Private Sub DisconnectClient()
    wskNS.Close
    ChangeNsState NsState_Disconnected
End Sub

Public Sub PingServer()
    On Error Resume Next
    
    wskNS.SendData "PNG" & vbCrLf
End Sub

Public Sub RequestSB()
    NS_Send "XFR " & NS_TrialID & " SB"
End Sub

Public Sub ChangeGTC(NewValue As String)
    NS_Send "GTC " & NS_TrialID & " " & NewValue
End Sub

Public Sub ChangeBLP(NewValue As String)
NS_Send "BLP " & NS_TrialID & " " & NewValue
End Sub

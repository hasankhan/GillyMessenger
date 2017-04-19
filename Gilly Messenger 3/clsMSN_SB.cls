VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMSN_SB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Sockets
    Private WithEvents wskSB As Winsock
Attribute wskSB.VB_VarHelpID = -1

'Enumerations
    Public Enum SbSessionType
        SbSession_Call
        SbSession_Ring
    End Enum

    Public Enum SbStateConstant
        SbState_Disconnected
        SbState_Connecting
        SbState_Connected
    End Enum
 
'Events
    Public Event SbError(Error As String)
    Public Event SocketError(Description As String)
    Public Event StateChanged()
    Public Event TypingUser(Email As String)
    Public Event MessageReceived(Email As String, Nick As String, FontName As String, FontColor As Long, FontBold As Boolean, FontItalic As Boolean, FontStrikethru As Boolean, FontUnderline As Boolean, Message As String, FakeNick As String)
    Public Event ContactJoined(Email As String, Nick As String)
    Public Event ContactLeft(Email As String)
    Public Event MessageFailure()
    Public Event InvitationReceived(AppName As String, AppGUID As String, Cookie As Double, Attributes As Collection)
    Public Event InvitationAccepted(Cookie As Double, Attributes As Collection)
    Public Event InvitationCancelled(Cookie As Double, CancelCode As String, Attributes As Collection)
    Public Event CustomMessageReceived(Email As String, MsgType As String, Header As String, Message As String)
    
'Properties
    'Server
    Private SB_Server As String
    Private SB_Port As Integer
    'Session
    Private SB_SessionID As Double
    Private SB_AuthCode As String
    Private SB_Contact As String
    Private SB_SessionType As Integer
    
'Server Vars
    Private SB_Data As String
    Private SB_TrialID As Double
    Private SB_State As Integer
    Private SB_Login As String
        
Public Property Get State() As Integer
    State = SB_State
End Property

Public Property Let Server(Data As String)
    SB_Server = Data
End Property

Public Property Get Server() As String
    Server = SB_Server
End Property

Public Property Let Port(Data As Integer)
    SB_Port = Data
End Property

Public Property Get Port() As Integer
    Port = SB_Port
End Property

Public Property Let SessionID(Data As Double)
    SB_SessionID = Data
End Property

Public Property Get SessionID() As Double
    SessionID = SB_SessionID
End Property

Public Property Let AuthCode(Data As String)
    SB_AuthCode = Data
End Property

Public Property Get AuthCode() As String
    AuthCode = SB_AuthCode
End Property

Public Property Let Contact(Data As String)
    SB_Contact = Data
End Property

Public Property Get Contact() As String
    Contact = SB_Contact
End Property

Public Property Let SessionType(Data As Integer)
    SB_SessionType = Data
End Property

Public Property Get SessionType() As Integer
    SessionType = SB_SessionType
End Property

Public Property Let Login(Data As String)
    SB_Login = Data
End Property

Public Property Get Login() As String
    Login = SB_Login
End Property

Public Property Get TrialID() As Double
    TrialID = SB_TrialID
End Property

Public Property Let Socket(Data As Winsock)
    Set wskSB = Data
End Property

Public Property Get Socket() As Winsock
    Set Socket = wskSB
End Property

Private Sub wskSb_Close()
    wskSB.Close
    Call ChangeSbState(SbState_Disconnected)
End Sub

Private Sub wskSb_Connect()
    If SB_SessionType = SbSession_Call Then
        SB_Send "USR " & SB_TrialID & " " & SB_Login & " " & SB_AuthCode
    Else
        SB_Send "ANS " & SB_TrialID & " " & SB_Login & " " & SB_AuthCode & " " & SB_SessionID
    End If
End Sub


Private Sub wskSb_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
    
    Dim Data As String, Commands() As String
    
    wskSB.GetData Data
    SB_Data = SB_Data & Data
    
    Commands = Split(SB_Data, vbCrLf)
    
    If Right$(Data, 2) = vbCrLf Then
        SB_Data = vbNullString
    Else
        SB_Data = Commands(UBound(Commands))
        If UBound(Commands) = 0 Then
            Exit Sub
        End If
    End If
    
    ReDim Preserve Commands(UBound(Commands) - 1)
    
    Dim CmdPtr As Integer, i As Integer, j As Integer, Temp As String, CmdParams() As String, aryTemp() As String, tmpCollection As Collection
    
    For CmdPtr = 0 To UBound(Commands)
    
        Debug.Print "<- "; Commands(CmdPtr)
        
        If Not Commands(CmdPtr) = vbNullString Then
        
            CmdParams = Split(Commands(CmdPtr))
        
            Select Case CmdParams(0)
                Case "ANS"
                    'ANS 1 OK
                    ChangeSbState SbState_Connected
                    
                Case "USR"
                    'USR 1 OK example@passport.com Example%20Name
                    ChangeSbState SbState_Connected
                    
                Case "IRO"
                    'IRO 1 2 2 myname@Sb.com My%20Name
                    RaiseEvent ContactJoined(CmdParams(4), MSN_Decode(CmdParams(5)))
                    
                Case "JOI"
                    'JOI name_123@hotmail.com Name_123
                    RaiseEvent ContactJoined(CmdParams(1), MSN_Decode(CmdParams(2)))
                
                Case "BYE"
                    'BYE example@passport.com
                    RaiseEvent ContactLeft(CmdParams(1))
                
                Case "MSG"
                    i = Val(CmdParams(3))
                    Temp = vbNullString
                    
                    For j = CmdPtr + 1 To UBound(Commands)
                        Temp = Temp & Commands(j) & vbCrLf
                        If Len(Temp) = i Then
                            ProcessSbMsg Commands(CmdPtr) & vbCrLf & Temp
                            Exit For
                        ElseIf Len(Temp) > i Then
                            Commands(j) = Mid$(Temp, i + 1, Len(Temp) - i - 2)
                            Temp = Left$(Temp, i)
                            ProcessSbMsg Commands(CmdPtr) & vbCrLf & Left$(Temp, i)
                            j = j - 2
                            Exit For
                        End If
                    Next
                    
                    If Len(Temp) < i Then
                        If Not SB_Data = vbNullString Then
                            Temp = Temp & SB_Data
                            SB_Data = vbNullString
                        End If
                    
                        If Len(Temp) = i Then
                            ProcessSbMsg Commands(CmdPtr) & vbCrLf & Temp
                        ElseIf Len(Temp) > i Then
                            SB_Data = Right$(Temp, Len(Temp) - i)
                            ProcessSbMsg Commands(CmdPtr) & vbCrLf & Left$(Temp, i)
                        Else
                           SB_Data = Commands(CmdPtr) & vbCrLf & Temp & SB_Data
                       End If
                    End If
                    
                    CmdPtr = j
                    
                Case "NAK"
                    'NAK 4
                    RaiseEvent MessageFailure
                    
                Case Else
                    If IsNumeric(CmdParams(0)) Then
                        RaiseEvent SbError(Commands(CmdPtr))
                    End If
            End Select
        End If
    Next
End Sub

Private Sub wskSb_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    SB_State = SbState_Disconnected
    RaiseEvent SocketError(Description)
    RaiseEvent StateChanged
End Sub

Private Sub ProcessSbMsg(Message As String)
    On Error GoTo Handler
    
    Dim MsgLines() As String, MsgType As String, i As Integer, MsgAttr() As String
    
    MsgLines = Split(Message, vbCrLf)
    
    MsgType = Split(Split(MsgLines(2))(1), ";")(0)
    
    Select Case MsgType
    Case "text/x-msmsgsinvite"
        Dim Attributes As Collection, InvitationCmd As String, InvitationCookie As Double
        Set Attributes = New Collection
        
        For i = 4 To UBound(MsgLines)
            MsgAttr = Split(MsgLines(i), ": ")
            If UBound(MsgAttr) = 1 Then
                Attributes.Add UTF8_Decode(MsgAttr(1)), MsgAttr(0)
            End If
        Next
                
        InvitationCmd = Attributes("Invitation-Command")
        InvitationCookie = Attributes("Invitation-Cookie")
        
        Attributes.Remove "Invitation-Command"
        Attributes.Remove "Invitation-Cookie"
        
        Select Case InvitationCmd
        Case "INVITE"
            Dim AppName As String, AppGUID As String
            AppName = Attributes("Application-Name")
            AppGUID = Attributes("Application-GUID")
            Attributes.Remove "Application-Name"
            Attributes.Remove "Application-GUID"
            RaiseEvent InvitationReceived(AppName, AppGUID, InvitationCookie, Attributes)
            
        Case "ACCEPT"
            RaiseEvent InvitationAccepted(InvitationCookie, Attributes)
            
        Case "CANCEL"
            Dim CancelCode As String
            If InCollection(Attributes, "Cancel-Code") Then
                CancelCode = Attributes("Cancel-Code")
                Attributes.Remove "Cancel-Code"
            End If
            RaiseEvent InvitationCancelled(InvitationCookie, CancelCode, Attributes)
        End Select
        
        Set Attributes = Nothing
            
    Case "text/x-msmsgscontrol"
        'MSG info@cracksoft.net CrackSoft 94
        'MIME-Version: 1.0
        'Content-Type: text/x-msmsgscontrol
        'TypingUser: info@ cracksoft.net
        
        RaiseEvent TypingUser(CStr(Split(MsgLines(3))(1)))
        
    Case "text/plain"
        'MSG info@cracksoft.net CrackSoft 94
        'MIME-Version: 1.0
        'Content-Type: text/plain; charset=UTF-8
        'X-MMS-IM-Format: FN=Tahoma; EF=B; CO=FF8000; CS=0; PF=0
        '
        'Hello
        Dim Temp As String
        Temp = Filter(MsgLines, "X-MMS-IM-Format")(0)
        MsgAttr = Split(Right$(Temp, Len(Temp) - 17), "; ")
        
        Dim MsgSubAttr() As String, FontName As String, FontColor As Long, FontBold As Boolean, FontItalic As Boolean, FontStrikethru As Boolean, FontUnderline As Boolean
        
        For i = 0 To UBound(MsgAttr)
            If Not MsgAttr(i) = vbNullString Then
                MsgSubAttr = Split(MsgAttr(i), "=")
                Select Case MsgSubAttr(0)
                Case "FN"
                    FontName = MsgSubAttr(1)
                Case "EF"
                    FontBold = InStr(MsgSubAttr(1), "B")
                    FontItalic = InStr(MsgSubAttr(1), "I")
                    FontStrikethru = InStr(MsgSubAttr(1), "S")
                    FontUnderline = InStr(MsgSubAttr(1), "U")
                Case "CO"
                    MsgSubAttr(1) = String$(6 - Len(MsgSubAttr(1)), "0") & MsgSubAttr(1)
                    FontColor = RGB(Val("&H" & Right$(MsgSubAttr(1), 2)), Val("&H" & Mid$(MsgSubAttr(1), 3, 2)), Val("&H" & Left$(MsgSubAttr(1), 2)))
                End Select
            End If
        Next
        
        Temp = Left$(Message, InStr(Message, vbCrLf & vbCrLf) - 4)
        If Not InStr(Temp, "P4-Context: ") = 0 Then
            Temp = Filter(Split(Temp, vbCrLf), "P4-Context: ")(0)
            Temp = Right$(Temp, Len(Temp) - 12)
        Else
            Temp = vbNullString
        End If
        i = InStr(Message, vbCrLf & vbCrLf) + 4
        RaiseEvent MessageReceived(CStr(Split(MsgLines(0))(1)), MSN_Decode(CStr(Split(MsgLines(0))(2))), URL_Decode(FontName), FontColor, FontBold, FontItalic, FontStrikethru, FontUnderline, UTF8_Decode(Mid$(Message, i, Len(Message) - i + 1)), Temp)
        
    Case Else
        Dim x As Integer
        For i = 1 To 3
            x = InStr(x + 1, Message, vbCrLf)
        Next
        i = InStr(Message, vbCrLf & vbCrLf)
        RaiseEvent CustomMessageReceived(CStr(Split(MsgLines(0))(1)), MsgType, Mid$(Message, x + 2, i - x - 2), UTF8_Decode(Mid$(Message, i + 4, Len(Message) - i + 5)))
    End Select
Handler:
End Sub

Public Sub Connect()
    On Error GoTo Handler
    
    ChangeSbState SbState_Connecting
    SB_Data = vbNullString
    SB_TrialID = 1
    wskSB.Close
    wskSB.Connect SB_Server, SB_Port
    Exit Sub
Handler:
    ChangeSbState SbState_Disconnected
End Sub

Private Sub SB_Send(Data As String)
    On Error Resume Next
    
    Debug.Print "-> "; Data
    Data = Data & vbCrLf
    Dim i As Integer
    For i = 1 To Len(Data) Step 1500
        wskSB.SendData Mid$(Data, i, 1500)
    Next
    SB_TrialID = SB_TrialID + 1
End Sub

Public Sub Disconnect()
    On Error Resume Next
    
    If SB_State = SbState_Connected Then
        wskSB.SendData "BYE"
        DoEvents
    End If
    wskSB.Close
    ChangeSbState SbState_Disconnected
End Sub

Private Sub ChangeSbState(State As Integer)
    SB_State = State
    RaiseEvent StateChanged
End Sub

Public Sub SendMessage(Message As String, Optional FontName As String, Optional FontColor As Long, Optional FontBold As Boolean, Optional FontItalic As Boolean, Optional FontStrikethru As Boolean, Optional FontUnderline As Boolean, Optional FakeNick As String)
    On Error Resume Next
    
    Dim SbMessage As String
    If Not Message = vbNullString Then
        Dim i As Integer
        Message = UTF8_Encode(Message)
        For i = 1 To Len(Message) Step 1500
            SbMessage = "MIME-Version: 1.0" & vbCrLf & _
            "Content-Type: text/plain; charset=UTF-8" & vbCrLf
            If Not FakeNick = vbNullString Then
                SbMessage = SbMessage & "P4-Context: " & FakeNick & vbCrLf
            End If
            SbMessage = SbMessage & "X-MMS-IM-Format: FN=" & URL_Encode(FontName) & "; EF=" & IIf(FontBold, "B", "") & IIf(FontItalic, "I", "") & IIf(FontBold, "B", "") & IIf(FontStrikethru, "S", "") & IIf(FontUnderline, "U", "") & "; CO=" & Hex(FontColor) & "; CS=0; PF=0" & vbCrLf & _
            vbCrLf & _
            Mid$(Message, i, 1500)
            wskSB.SendData "MSG " & SB_TrialID & " N " & Len(SbMessage) & vbCrLf & SbMessage
            SB_TrialID = SB_TrialID + 1
            DoEvents
        Next
    Else
        SbMessage = "MIME-Version: 1.0" & vbCrLf & _
        "Content-Type: text/plain; charset=UTF-8" & vbCrLf
        If Not FakeNick = vbNullString Then
            SbMessage = SbMessage & "P4-Context: " & FakeNick & vbCrLf
        End If
        SbMessage = SbMessage & "X-MMS-IM-Format: FN=" & URL_Encode(FontName) & "; EF=" & IIf(FontBold, "B", "") & IIf(FontItalic, "I", "") & IIf(FontBold, "B", "") & IIf(FontStrikethru, "S", "") & IIf(FontUnderline, "U", "") & "; CO=" & Hex(FontColor) & "; CS=0; PF=0" & vbCrLf & _
        vbCrLf
        wskSB.SendData "MSG " & SB_TrialID & " N " & Len(SbMessage) & vbCrLf & SbMessage
        SB_TrialID = SB_TrialID + 1
    End If
End Sub

Public Sub SendCustomMessage(MsgType As String, Header As String, Message As String)
    On Error Resume Next
    
    Dim SbMessage As String
    If Not Message = vbNullString Then
        Dim i As Integer
        Message = UTF8_Encode(Message)
        For i = 1 To Len(Message) Step 1500
            SbMessage = "MIME-Version: 1.0" & vbCrLf & _
            "Content-Type: " & MsgType & vbCrLf
            If Not Header = vbNullString Then
                SbMessage = SbMessage & Header & vbCrLf
            End If
            SbMessage = SbMessage & vbCrLf & Mid$(Message, i, 1500)
            wskSB.SendData "MSG " & SB_TrialID & " N " & Len(SbMessage) & vbCrLf & SbMessage
            SB_TrialID = SB_TrialID + 1
            DoEvents
        Next
    Else
        SbMessage = "MIME-Version: 1.0" & vbCrLf & _
        "Content-Type: " & MsgType & vbCrLf
        If Not Header = vbNullString Then
            SbMessage = SbMessage & Header & vbCrLf
        End If
        SbMessage = SbMessage & vbCrLf
        wskSB.SendData "MSG " & SB_TrialID & " N " & Len(SbMessage) & vbCrLf & SbMessage
        SB_TrialID = SB_TrialID + 1
    End If
End Sub

Public Sub InviteContact(Email As String)
    SB_Send "CAL " & SB_TrialID & " " & Email
End Sub

Public Sub SendTypingNotification()
    On Error Resume Next
    
    Dim SbMessage As String
    
    SbMessage = "MIME-Version: 1.0" & vbCrLf & _
    "Content-Type: text/x-msmsgscontrol" & vbCrLf & _
    "TypingUser: " & Login & vbCrLf & vbCrLf
    
    wskSB.SendData "MSG " & SB_TrialID & " U " & Len(SbMessage) & vbCrLf & SbMessage
    SB_TrialID = SB_TrialID + 1
End Sub

Public Sub AcceptInvitation(Cookie As Double, ParamArray Attributes())
    On Error Resume Next
    
    Dim i As Integer
    For i = 0 To UBound(Attributes)
        Attributes(i) = UTF8_Encode(Attributes(i))
    Next
    
    Dim SbMessage As String
    
    SbMessage = "MIME-Version: 1.0" & vbCrLf & _
    "Content-Type: text/x-msmsgsinvite; charset=UTF-8" & vbCrLf & vbCrLf & _
    "Invitation-Command: ACCEPT" & vbCrLf & _
    "Invitation-Cookie: " & Cookie
    
    If UBound(Attributes) > -1 Then
        SbMessage = SbMessage & vbCrLf & Join(Attributes, vbCrLf)
    End If
    
    SbMessage = SbMessage & vbCrLf
    
    wskSB.SendData "MSG " & SB_TrialID & " N " & Len(SbMessage) & vbCrLf & SbMessage
    SB_TrialID = SB_TrialID + 1
End Sub

Public Sub CancelInvitation(Cookie As Double, CancelCode As String, ParamArray Attributes())
    On Error Resume Next
    
    Dim i As Integer
    For i = 0 To UBound(Attributes)
        Attributes(i) = UTF8_Encode(Attributes(i))
    Next
    
    Dim SbMessage As String
    
    SbMessage = "MIME-Version: 1.0" & vbCrLf & _
    "Content-Type: text/x-msmsgsinvite; charset=UTF-8" & vbCrLf & vbCrLf & _
    "Invitation-Command: CANCEL" & vbCrLf & _
    "Invitation-Cookie: " & Cookie & vbCrLf & _
    "Cancel-Code: " & CancelCode
    
    If UBound(Attributes) > -1 Then
        SbMessage = SbMessage & vbCrLf & Join(Attributes, vbCrLf)
    End If
    
    SbMessage = SbMessage & vbCrLf
    
    wskSB.SendData "MSG " & SB_TrialID & " U " & Len(SbMessage) & vbCrLf & SbMessage
    SB_TrialID = SB_TrialID + 1
End Sub

Public Sub SendInvitation(AppName As String, AppGUID As String, Cookie As Double, ParamArray Attributes())
    On Error Resume Next
    
    Dim i As Integer
    For i = 0 To UBound(Attributes)
        Attributes(i) = UTF8_Encode(Attributes(i))
    Next
    
    Dim SbMessage As String
    
    SbMessage = "MIME-Version: 1.0" & vbCrLf & _
    "Content-Type: text/x-msmsgsinvite; charset=UTF-8" & vbCrLf & vbCrLf & _
    "Application-Name: " & AppName & vbCrLf & _
    "Application-GUID: " & AppGUID & vbCrLf & _
    "Invitation-Command: INVITE" & vbCrLf & _
    "Invitation-Cookie: " & Cookie & vbCrLf & _
    Join(Attributes, vbCrLf)
    
    wskSB.SendData "MSG " & SB_TrialID & " N " & Len(SbMessage) & vbCrLf & SbMessage
    SB_TrialID = SB_TrialID + 1
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMSN_FTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Socket
    Private WithEvents wskFTP As Winsock
Attribute wskFTP.VB_VarHelpID = -1
    
'Enumerations
    Public Enum FtpStateConstant
        FtpState_Disconnected
        FtpState_Connecting
        FtpState_Connected
        FtpState_Negotiating
        FtpState_Transfer
    End Enum

    Public Enum FtpTransferTypeConstant
        FtpTransferType_Send
        FtpTransferType_Receive
    End Enum

'Events
    Public Event FtpError(Error As String)
    Public Event SocketError(Description As String)
    Public Event Progress(PercentDone As Integer, BytesTransferred As Double, TransferRate As String)
    Public Event StateChanged()
    Public Event TransferComplete()

'Properties
    Private FTP_Protocol As String
    Private FTP_Login As String
    Private FTP_AuthCookie As Double
    Private FTP_State As Integer
    Private FTP_File As String
    Private FTP_FileSize As Double
    Private FTP_FilePath As String
    Private FTP_TransferType As Integer
'Other
    Private FTP_Data As String
    Private FTP_Layer As Integer
    Private BytesTransferred As Double
    Private TransferStamp As Date
    Private PacketSent As Boolean
    Private FileNum As Integer
    
Public Property Let Protocol(Data As String)
    FTP_Protocol = Data
End Property

Public Property Get Protocol() As String
    Protocol = FTP_Protocol
End Property

Public Property Let Login(Data As String)
    FTP_Login = Data
End Property

Public Property Get Login() As String
    Login = FTP_Login
End Property

Public Property Let AuthCookie(Data As Double)
    FTP_AuthCookie = Data
End Property

Public Property Get AuthCookie() As Double
    AuthCookie = FTP_AuthCookie
End Property

Public Property Let File(Data As String)
    FTP_File = Data
End Property

Public Property Get File() As String
    File = FTP_File
End Property

Public Property Let FileSize(Data As Double)
    FTP_FileSize = Data
    BytesTransferred = 0
End Property

Public Property Get FileSize() As Double
    FileSize = FTP_FileSize
End Property

Public Property Let FilePath(Data As String)
    FTP_FilePath = Data
End Property

Public Property Get FilePath() As String
    FilePath = FTP_FilePath
End Property

Public Property Let TransferType(Data As Integer)
    FTP_TransferType = Data
End Property

Public Property Get TransferType() As Integer
    TransferType = FTP_TransferType
End Property

Public Property Get State() As Integer
    State = FTP_State
End Property

Public Property Let Socket(Data As Winsock)
    Set wskFTP = Data
End Property

Private Sub Class_Initialize()
    FTP_Protocol = "MSNFTP"
End Sub

Public Sub Connect(Host As String, Port As Integer)
    On Error Resume Next
    
    ChangeFtpState FtpState_Connecting
    FTP_Layer = 0
    wskFTP.Close
    wskFTP.Connect Host, Port
End Sub

Public Sub Disconnect()
    On Error Resume Next
    
    Close #FileNum
    wskFTP.Close
    Call ChangeFtpState(FtpState_Disconnected)
End Sub

Public Function Listen(Optional Port As Integer) As Integer
    On Error Resume Next
    
    ChangeFtpState FtpState_Connecting
    FTP_Layer = 0
    wskFTP.Close
    If Port = 0 Then
        Port = 1862
    Else
        Port = Port - 1
    End If
    Do
        Err.Clear
        Port = Port + 1
        wskFTP.LocalPort = Port
        wskFTP.Listen
    Loop Until Not Err.Number = 10048
    Listen = Port
End Function

Private Sub wskFTP_Close()
    wskFTP.Close
    ChangeFtpState FtpState_Disconnected
End Sub

Private Sub wskFTP_Connect()
    ChangeFtpState FtpState_Connected
    If FTP_TransferType = FtpTransferType_Receive Then
        Call FTP_Send("VER " & FTP_Protocol)
    End If
End Sub

Private Sub wskFTP_ConnectionRequest(ByVal requestID As Long)
    wskFTP.Close
    wskFTP.Accept requestID
End Sub

Private Sub wskFTP_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
    
    Dim Data As String
    wskFTP.GetData Data
    FTP_Data = FTP_Data & Data
    
    If FTP_State <> FtpState_Transfer Or FTP_TransferType = FtpTransferType_Send Then
        Dim Commands() As String
    
        Commands = Split(FTP_Data, vbCrLf)
    
        If Right$(Data, 2) = vbCrLf Then
            FTP_Data = vbNullString
        Else
            FTP_Data = Commands(UBound(Commands))
            If UBound(Commands) = 0 Then
                Exit Sub
            End If
        End If
    
        ReDim Preserve Commands(UBound(Commands) - 1)
    
        Dim CmdPtr As Integer, i As Integer, CmdParams() As String
    
        For CmdPtr = 0 To UBound(Commands)
    
            Debug.Print "<- "; Commands(CmdPtr)
        
            If Not Commands(CmdPtr) = vbNullString Then
        
                CmdParams = Split(Commands(CmdPtr))
        
                Select Case CmdParams(0)
                Case "VER"
                    If InStr(Commands(CmdPtr), FTP_Protocol) = 0 Then
                        Call Disconnect
                    Else
                        ChangeFtpState FtpState_Negotiating
                        If FTP_TransferType = FtpTransferType_Send Then
                            FTP_Layer = 1
                            Call FTP_Send("VER " & FTP_Protocol)
                        Else
                            Call FTP_Send("USR " & FTP_Login & " " & FTP_AuthCookie)
                        End If
                    End If
                Case Else
                    Select Case FTP_Protocol
                    Case "MSNFTP"
                        Select Case CmdParams(0)
                        Case "USR"
                            If Not CmdParams(1) = FTP_Login Or Not CmdParams(2) = FTP_AuthCookie Then
                                Call Disconnect
                            Else
                                If FTP_Layer = 1 Then
                                    FTP_Layer = 2
                                    Call FTP_Send("FIL " & FTP_FileSize)
                                Else
                                    Call Disconnect
                                End If
                            End If
                        Case "FIL"
                            If Not Val(CmdParams(1)) = FTP_FileSize Then
                                Call Disconnect
                            Else
                                TransferStamp = Now()
                                ChangeFtpState FtpState_Transfer
                                FileNum = FreeFile
                                Kill FTP_FilePath
                                Open FTP_FilePath For Binary Access Write As #FileNum
                                Call FTP_Send("TFR")
                            End If
                        Case "TFR"
                            TransferStamp = Now()
                            ChangeFtpState FtpState_Transfer
                            If FTP_TransferType = FtpTransferType_Send Then
                                If FTP_Layer = 2 Then
                                    FTP_Layer = 3
                                    FileNum = FreeFile
                                    Open FTP_FilePath For Binary Access Read As #FileNum
                                    Dim j As Double, DataPacket As String
                                    j = FTP_FileSize
                                    Do
                                        If Not wskFTP.State = sckConnected Then
                                            Exit Do
                                        End If
                                        
                                        PacketSent = False
                                        If j < 2045 Then
                                            DataPacket = String(j, vbNullChar)
                                            Get #FileNum, , DataPacket
                                            wskFTP.SendData Chr$(0) & Chr$(j Mod 256) & Chr$(j \ 256) & DataPacket
                                            BytesTransferred = BytesTransferred + j
                                            j = 0
                                        Else
                                            DataPacket = String(2045, vbNullChar)
                                            Get #FileNum, , DataPacket
                                            wskFTP.SendData Chr$(0) & Chr$(2045 Mod 256) & Chr$(2045 \ 256) & DataPacket
                                            BytesTransferred = BytesTransferred + 2045
                                            j = j - 2045
                                        End If
                                        Do Until PacketSent = True
                                            DoEvents
                                        Loop
                                        RaiseEvent Progress((BytesTransferred * 100) / FTP_FileSize, BytesTransferred, ConvertBytes(BytesTransferred / DateDiff("s", TransferStamp, Now()), True))
                                    Loop Until j = 0
                                    Close #FileNum
                                    wskFTP.SendData Chr$(1) & Chr$(0) & Chr$(0)
                                Else
                                    Call Disconnect
                                End If
                            End If
                        Case "BYE"
                            If CmdParams(1) = "16777987" Or CmdParams(1) = "16777989" Then
                                RaiseEvent TransferComplete
                                Call Disconnect
                            Else
                                RaiseEvent FtpError(CmdParams(1))
                                Call Disconnect
                            End If
                        End Select
                    End Select
                End Select
            End If
        Next
    Else
        Dim PacketLen As Integer
        Do
            PacketLen = Asc(Mid$(FTP_Data, 2, 1)) + (256 * Asc(Mid$(FTP_Data, 3, 1)))
            If Len(FTP_Data) >= PacketLen + 3 Then
                Put #FileNum, , Mid$(FTP_Data, 4, PacketLen)
                BytesTransferred = BytesTransferred + PacketLen
                RaiseEvent Progress((BytesTransferred * 100) / FTP_FileSize, BytesTransferred, ConvertBytes(BytesTransferred / DateDiff("s", TransferStamp, Now()), True))
                FTP_Data = Right$(FTP_Data, Len(FTP_Data) - PacketLen - 3)
                If BytesTransferred >= FTP_FileSize Then
                    Close #FileNum
                    Call FTP_Send("BYE 16777989")
                    Do Until PacketSent
                        DoEvents
                    Loop
                    RaiseEvent TransferComplete
                    Call Disconnect
                    Exit Do
                ElseIf Len(FTP_Data) < 4 Then
                    Exit Do
                End If
            Else
                Exit Do
            End If
        Loop
    End If
End Sub

Private Sub wskFTP_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    RaiseEvent SocketError(Description)
    Call Disconnect
End Sub

Private Sub ChangeFtpState(State As Integer)
    FTP_State = State
    RaiseEvent StateChanged
End Sub

Private Sub FTP_Send(Data As String)
    On Error Resume Next
    
    Debug.Print "-> "; Data
    PacketSent = False
    wskFTP.SendData Data & vbCrLf
End Sub

Private Sub wskFTP_SendComplete()
    PacketSent = True
End Sub

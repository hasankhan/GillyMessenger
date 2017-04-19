Attribute VB_Name = "modRemoteControl"
Option Explicit
Option Compare Text

Public RemoteControl As Boolean
Public RC_Accounts As New Collection, RC_Sessions As New Collection
Private hFile As Integer, hFiles As New Collection, UserDir As New Collection, GMDir As String

Public Function RcProcess(IMWindow As frmChat, RcUser As String, RcLogin As String, RcCommand As String) As String
    On Error Resume Next
    If Not InCollection(hFiles, RcUser & "_" & RcLogin) Then
        Dim CmdParams() As String, Param As String, Temp As String, i As Integer
        CmdParams = Split(RcCommand)
        If UBound(CmdParams) > 0 Then
            Param = Right$(RcCommand, Len(RcCommand) - InStr(RcCommand, " "))
        End If
        GMDir = CurDir$
        If InCollection(UserDir, RcUser & "_" & RcLogin) Then
            Temp = UserDir(RcUser & "_" & RcLogin)
            ChDrive Left$(Temp, 2)
            ChDir Temp
        End If
        Select Case CmdParams(0)
        Case "dir"
            If RC_Accounts(RcLogin).Item("dirBrowsing") Then
                RcProcess = ListDir(Param)
            End If
        Case "cd", "chdir"
            If RC_Accounts(RcLogin).Item("dirBrowsing") Then
                If Not CurDir$ = Param Then
                    Temp = CurDir$
                    ChDir Param
                    If Not CurDir$ = Temp Then
                        RcProcess = "Directory changed to " & Param
                        SetCollectionItem UserDir, RcUser & "_" & RcLogin, CurDir$
                    Else
                        RcProcess = "Directory change failed."
                    End If
                Else
                    RcProcess = "Already in " & Param
                End If
            End If
        Case "get"
            If RC_Accounts(RcLogin).Item("dirBrowsing") Then
                If InStr(Param, ":\") = 0 Then
                    SendFile IMWindow, CurDir$ & "\" & Param
                Else
                    SendFile IMWindow, Param
                End If
            End If
        Case "copy"
            If RC_Accounts(RcLogin).Item("shellCommands") Then
                Dim FileSrc As String, FileDest As String
                If Param Like """*"" ""*""" Then
                    i = InStr(2, Param, """")
                    FileSrc = Mid$(Param, 2, i - 2)
                    FileDest = Mid$(Param, i + 3, Len(Param) - i - 3)
                ElseIf Param Like """*"" *" Then
                    i = InStr(2, Param, """")
                    FileSrc = Mid$(Param, 2, i - 2)
                    FileDest = Right$(Param, Len(Param) - i - 1)
                ElseIf Param Like "* ""*""" Then
                    i = InStr(Param, """")
                    FileSrc = Left$(Param, InStr(Param, " ") - 1)
                    FileDest = Mid$(Param, i + 3, Len(Param) - i - 4)
                Else
                    FileSrc = CmdParams(1)
                    FileDest = CmdParams(2)
                End If
                RcProcess = FileSrc & " copied to " & FileDest
            End If
        Case "del", "erase"
            If RC_Accounts(RcLogin).Item("shellCommands") Then
                Kill Param
                RcProcess = Param & " deleted."
            End If
        Case "ren", "rename"
            If RC_Accounts(RcLogin).Item("shellCommands") Then
                Name CmdParams(1) As CmdParams(2)
                RcProcess = CmdParams(1) & " renamed to " & CmdParams(2)
            End If
        Case "md", "mkdir"
            If RC_Accounts(RcLogin).Item("shellCommands") Then
                MkDir Param
                RcProcess = "Directory " & Param & " created."
            End If
        Case "rd", "rmdir"
            If RC_Accounts(RcLogin).Item("shellCommands") Then
                RmDir Param
                RcProcess = "Directory " & Param & " removed."
            End If
        Case "msgr"
            If RC_Accounts(RcLogin).Item("msgrControl") Then
                Call SendMsg(IMWindow, Param)
                RcProcess = Param & " command received."
            End If
        Case "execute"
            If RC_Accounts(RcLogin).Item("shellCommands") Then
                If ShellExecuteEx(Param) Then
                    RcProcess = Param & " executed."
                Else
                    RcProcess = Param & " execution failed."
                End If
            End If
        Case "list"
            If RC_Accounts(RcLogin).Item("msgrControl") Then
                Select Case CmdParams(1)
                Case "online"
                    RcProcess = OnlineList
                Case "chats"
                    RcProcess = ChatList
                End Select
            End If
        Case "dump"
            If RC_Accounts(RcLogin).Item("dirBrowsing") Then
                hFiles.Add FreeFile, RcUser & "_" & RcLogin
                Open Param For Output As hFiles(RcUser & "_" & RcLogin)
                RcProcess = Param & " opened for writting."
            End If
        Case "read"
            If RC_Accounts(RcLogin).Item("dirBrowsing") Then
                hFile = FreeFile
                Open Param For Binary As hFile
                Temp = Space$(LOF(hFile))
                Get #hFile, , Temp
                Close hFile
                RcProcess = Temp
            End If
        Case "logout"
            RC_Sessions.Remove RcUser
            hFiles.Remove RcUser & "_" & RcLogin
            UserDir.Remove RcUser & "_" & RcLogin
            RcProcess = "Logged out."
        Case "reboot"
            If RC_Accounts(RcLogin).Item("shellCommands") Then
                SendMsg IMWindow, "Rebooting the PC."
                DoEvents
                ExitWindowsEx EWX_FORCE Or EWX_REBOOT, 0
            End If
        Case "shutdown"
            If RC_Accounts(RcLogin).Item("shellCommands") Then
                SendMsg IMWindow, "Shutting down the PC."
                DoEvents
                ExitWindowsEx EWX_FORCE Or EWX_SHUTDOWN, 0
            End If
        Case Else
            If RC_Accounts(RcLogin).Item("dirBrowsing") Then
                If Len(RcCommand) = 2 And Right$(RcCommand, 1) = ":" Then
                    Temp = Left$(CurDir$, 2)
                    If Not RcCommand = Temp Then
                        ChDrive RcCommand
                        If Not Temp = Left$(CurDir$, 2) Then
                            RcProcess = "Drive changed to " & Left$(RcCommand, 1) & "."
                            SetCollectionItem UserDir, RcUser & "_" & RcLogin, CurDir$
                        Else
                            RcProcess = "Drive change failed."
                        End If
                    Else
                        RcProcess = "Already in " & Left$(Temp, 1) & " drive."
                    End If
                End If
            End If
        End Select
        ChDrive Left$(GMDir, 2)
        ChDir GMDir
    Else
        If Not RcCommand = "." Then
            Print #hFiles(RcUser & "_" & RcLogin), RcCommand
            RcProcess = """" & RcCommand & """ dumped."
        Else
            Close #hFiles(RcUser & "_" & RcLogin)
            hFiles.Remove RcUser & "_" & RcLogin
            RcProcess = "File dump complete."
        End If
    End If
End Function

Private Function ListDir(Parm As String) As String
    On Error Resume Next
    
    Dim WorkDir As String, strFile As String
    WorkDir = Replace$(CurDir$ & "\", "\\", "\")
    strFile = Dir$(WorkDir & Parm, vbDirectory Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive Or vbNormal Or vbVolume)
    Do Until strFile = vbNullString
        DoEvents
        If (GetAttr(WorkDir & strFile) And vbDirectory) = vbDirectory Then
            ListDir = ListDir & "[" & strFile & "]" & vbCrLf
        Else
            ListDir = ListDir & strFile & vbCrLf
        End If
        strFile = Dir$
    Loop
    If ListDir = vbNullString Then
        ListDir = "File Not Found."
    Else
        ListDir = Left$(ListDir, Len(ListDir) - 2)
    End If
End Function

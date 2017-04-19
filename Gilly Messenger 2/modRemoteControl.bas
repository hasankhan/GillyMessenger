Attribute VB_Name = "modRemoteControl"
Dim WorkDir As String, Parm As String, Strike As Integer, RcProcStart As Boolean, RcWindow As Long
Dim FileDump As Boolean

Public Function RcProcess(CMD As String, RcForm As Form) As String
On Error GoTo Reporter
If RcLoggedIn = True And RcForm.hwnd = RcWindow And FileDump = False Then
    If Left$(UCase$(CMD), 3) = "DIR" Then
        Parm = Right$(CMD, Len(CMD) - 3)
        Parm = Trim$(Parm)
        RcProcess = ListDir(Parm)
    ElseIf UCase$(Left$(CMD, 2)) = "CD" Then
        Parm = Right$(CMD, Len(CMD) - 2)
        Parm = Trim$(Parm)
        If Parm <> CurDir$ Then
            Temp = CurDir$
            ChDir Parm
            If CurDir$ <> Temp Then
                RcProcess = "Directory changed to " & Parm
            Else
                RcProcess = "Directory change failed."
            End If
        Else
            RcProcess = "Already in " & Parm
        End If
    ElseIf Len(CMD) = 2 And Right$(CMD, 1) = ":" Then
        CMD = UCase$(CMD)
        Temp = Left$(CurDir$, 2)
        If CMD <> Left$(CurDir$, 2) Then
            ChDrive CMD
            If Temp <> Left$(CurDir$, 2) Then
                RcProcess = "Drive changed to " & CMD
            Else
                RcProcess = "Drive change failed."
            End If
        Else
            RcProcess = "Already in " & CMD
        End If
    ElseIf UCase$(Left$(CMD, 5)) = "MSGR " Then
        Parm = Right$(CMD, Len(CMD) - 5)
        RcForm.txtMessage.Text = "'" & Parm & "' command received."
        RcForm.cmdSend.Value = 1
        RcProcess = Parm
    ElseIf UCase$(Left$(CMD, 8)) = "EXECUTE " Then
        CMD = Right$(CMD, Len(CMD) - 8)
        If Left$(CMD, 1) = Chr$(34) Then
            CMD = Right$(CMD, Len(CMD) - 1)
            X = InStr(CMD, Chr$(34))
        Else
            X = InStr(CMD, Chr$(vbKeySpace))
        End If
        If X > 0 Then
            Parm = Right$(CMD, Len(CMD) - X)
            CMD = Left$(CMD, X - 1)
        End If
        X = ShellExecute(RcWindow, vbNullString, CMD, Parm, vbNullString, 1)
        If X > 32 Then
            RcProcess = CMD & " executed."
        Else
            RcProcess = CMD & " execution failed."
        End If
    ElseIf UCase$(CMD) = "LIST ONLINE" Then
        RcProcess = ListOnline
    ElseIf UCase$(Left$(CMD, 5)) = "DUMP " Then
        CMD = Right$(CMD, Len(CMD) - 5)
        Open CMD For Output As #6
        RcProcess = "'" & CMD & "' opened."
        FileDump = True
    ElseIf UCase$(Left$(CMD, 5)) = "COPY " Then
        CMD = Right$(CMD, Len(CMD) - 5)
        X = InStr(CMD, " ")
        Parm = Left$(CMD, X - 1)
        Temp = Right$(CMD, Len(CMD) - X)
        FileCopy Parm, Temp
        RcProcess = "'" & Parm & "' copied to '" & Temp & "'"
    ElseIf UCase$(Left$(CMD, 4)) = "DEL " Then
        Parm = Right$(CMD, Len(CMD) - 4)
        Kill Parm
        RcProcess = "'" & Parm & "' deleted."
    ElseIf UCase$(Left$(CMD, 4)) = "REN " Then
        CMD = Right$(CMD, Len(CMD) - 4)
        X = InStr(CMD, " ")
        Parm = Left$(CMD, X - 1)
        Temp = Right$(CMD, Len(CMD) - X)
        Name Parm As Temp
        RcProcess = "'" & Parm & "' renamed to '" & Temp & "'"
    ElseIf UCase$(Left$(CMD, 3)) = "MD " Then
        Parm = Right$(CMD, Len(CMD) - 3)
        MkDir Parm
        RcProcess = "Directory '" & Parm & "' created."
    ElseIf UCase$(Left$(CMD, 3)) = "RD " Then
        Parm = Right$(CMD, Len(CMD) - 3)
        RmDir Parm
        RcProcess = "Directory '" & Parm & "' removed."
    ElseIf UCase$(Left$(CMD, 5)) = "READ " Then
        Parm = Right$(CMD, Len(CMD) - 5)
        Temp = Space$(FileLen(Parm))
        Open Parm For Binary As #7
        Get #7, , Temp
        Close #7
        RcProcess = Temp
    ElseIf UCase$(CMD) = "LOGOUT" Then
        RcProcess = "Logged Out."
        RcLoggedIn = False
        RcProcStart = False
        RcWindow = 0
    ElseIf UCase$(CMD) = "REBOOT" Then
        RcForm.txtMessage.Text = "Rebooting the PC."
        RcForm.cmdSend.Value = 1
        DoEvents
        ExitWindowsEx EWX_FORCE Or EWX_REBOOT, 0
    ElseIf UCase$(CMD) = "SHUTDOWN" Then
        RcForm.txtMessage.Text = "Shutting down the PC."
        RcForm.cmdSend.Value = 1
        DoEvents
        ExitWindowsEx EWX_FORCE Or EWX_SHUTDOWN, 0
    End If
ElseIf RcLoggedIn = True And RcForm.hwnd = RcWindow And FileDump = True Then
    If CMD = vbCrLf & "." Then
        Close #6
        FileDump = False
        RcProcess = "File dump complete."
    Else
        Print #6, CMD
        RcProcess = "'" & CMD & "' dumped."
    End If
Else
    If Right$(CMD, 25) = "GM Remote Control Request" And RcWindow = 0 Then
        RcProcess = "Welcome to GM Remote Control 1.0" & vbCrLf
        RcProcess = RcProcess & "Enter your Username and Password."
        RcProcStart = True
        RcWindow = RcForm.hwnd
    ElseIf RcProcStart = True And RcForm.hwnd = RcWindow Then
        If InStr(CMD, vbCrLf) > 0 Then
            If Left$(CMD, InStr(CMD, vbCrLf) - 1) = RcUsername And Right$(CMD, Len(CMD) - InStr(CMD, vbCrLf) - 1) = RcPassword Then
                RcProcess = "Logged In."
                RcLoggedIn = True
                RcUser = RcForm.lblStatus.Tag
            Else
                Strike = Strike + 1
                If Strike >= 3 Then
                    RcProcStart = False
                    RcWindow = 0
                    Strike = 0
                    RcForm.txtMessage.Text = "Failure Strike 3"
                    RcForm.cmdSend.Value = 1
                    DoEvents
                    Call cForm.Form_KeyPress(vbKeyEscape)
                Else
                    RcProcess = "Failure Strike " & Strike
                End If
            End If
        End If
    End If
End If
Exit Function
Reporter:
Resume Here
Here:
If Err.Number = 75 Then
    Resume Next
ElseIf Err.Number = 5 Then
    RcProcess = "Invalid command format."
Else
    RcProcess = Err.Description
End If
End Function

Private Function ListDir(Parm As String) As String
On Error Resume Next
WorkDir = Replace$(CurDir$ & "\", "\\", "\")
Temp = Dir$(WorkDir & Parm, vbDirectory Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive Or vbNormal Or vbVolume)
Do Until Temp = vbNullString
    DoEvents
    If (GetAttr(WorkDir & Temp) And vbDirectory) = vbDirectory Then
        ListDir = ListDir & "[" & Temp & "]" & vbCrLf
    Else
        ListDir = ListDir & Temp & vbCrLf
    End If
    Temp = Dir$
Loop
If ListDir = vbNullString Then
    ListDir = "File Not Found."
Else
    ListDir = Left$(ListDir, Len(ListDir) - 2)
End If
End Function

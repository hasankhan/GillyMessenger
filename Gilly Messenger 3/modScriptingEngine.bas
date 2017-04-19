Attribute VB_Name = "modScriptingEngine"
Option Explicit
Option Compare Text

Public GMScripts As New Collection, ScriptQue As New Collection

Public Sub LoadScript(File As String, Optional Params)
    On Error Resume Next
    
    If FileExists(File) Then
        Dim NewScript As Collection
        Set NewScript = New Collection
        
        NewScript.Add New Collection, "vars"
        NewScript.Add New Collection, "labels"
        NewScript.Add New Collection, "hfiles"
        NewScript.Add New Collection, "script_main": NewScript.Add 1, "pos_main"
        NewScript.Add 0, "eventcount"
        NewScript.Add 0.1, "sleep": NewScript.Add vbNullString, "lastexec"
        NewScript.Add File, "key"
        
        If Not IsMissing(Params) Then
            Call AddScriptVar(NewScript, "$args", UBound(Params) + 1)
            Dim i As Integer
            For i = 0 To UBound(Params)
                Call AddScriptVar(NewScript, "$arg[" & i + 1 & "]", CStr(Params(i)))
            Next
        Else
            Call AddScriptVar(NewScript, "$args", 0)
        End If
        
        Dim FileNum As Integer, Data As String, strEvent As String, strSub As String
        strEvent = "main"
        FileNum = FreeFile
        Open File For Input As #FileNum
        Do Until EOF(FileNum)
            Line Input #FileNum, Data
            Data = TrimX(Data)
            If Left$(Data, 1) <> "'" And Data <> vbNullString Then
                If strSub = vbNullString Then
                    If Left$(Data, 4) = "sub " Then
                        strSub = LCase$(Right$(Data, Len(Data) - InStr(Data, " ")))
                        NewScript.Add New Collection, "script_" & strSub
                        NewScript.Add 1, "pos_" & strSub
                    ElseIf Left$(Data, 6) = "event " Then
                        strEvent = LCase$(Right$(Data, Len(Data) - InStr(Data, " ")))
                        NewScript.Add New Collection, "script_" & strEvent
                        NewScript.Add 1, "pos_" & strEvent
                        SetCollectionItem NewScript, "eventcount", NewScript.Item("eventcount") + 1
                    ElseIf Data = "end event" Then
                        strEvent = "main"
                    ElseIf Data Like "[[]*[]]" Then
                        SetCollectionItem NewScript("labels"), SubString(Data, "[", "]"), NewScript.Item("script_" & strEvent).Count + 1
                    Else
                        NewScript.Item("script_" & strEvent).Add Data
                    End If
                Else
                    If Data = "end sub" Then
                        strSub = vbNullString
                    ElseIf Data Like "[[]*[]]" Then
                        SetCollectionItem NewScript("labels"), SubString(Data, "[", "]"), NewScript.Item("script_" & strSub).Count + 1
                    Else
                        NewScript.Item("script_" & strSub).Add Data
                    End If
                End If
            End If
        Loop
        Close #FileNum
        
        If NewScript("eventcount") = 0 Then
            frmMain.lblStatus.Caption = "Executing Script..."
        Else
            frmMain.lblStatus.Caption = "Script Loaded."
        End If
        
        If NewScript("script_main").Count > 0 And Not frmMain.tmrGMScript_Main.Enabled Then
            frmMain.tmrGMScript_Main.Enabled = True
        End If
        
        GMScripts.Add NewScript, File
        Set NewScript = Nothing
    End If
End Sub

Public Sub EndScript(Script As Collection, Optional Complete As Boolean, Optional EraseMenu As Boolean = True)
    On Error Resume Next
    
    Dim i As Integer
    For i = 1 To Script("hfiles").Count
        Close #Script("hfiles").Item(i)
    Next
    
    Dim strKey As String
    strKey = Script("key")
    With Script
        For i = 1 To .Count
            Script.Remove 1
        Next
    End With
    
    GMScripts.Remove strKey
    
    If EraseMenu Then
        For i = 1 To frmMain.mnuTools_GMScript_Script.UBound
            If frmMain.mnuTools_GMScript_Script(i).Tag = strKey And frmMain.mnuTools_GMScript_Script(i).Checked Then
                frmMain.mnuTools_GMScript_Script(i).Checked = False
                Exit For
            End If
        Next
        If i > frmMain.mnuTools_GMScript_Script.UBound Then
            For i = 1 To frmMain.mnuTools_GMScript_Other.UBound
                If frmMain.mnuTools_GMScript_Other(i).Tag = strKey Then
                    Unload frmMain.mnuTools_GMScript_Other(i)
                    Exit For
                End If
            Next
        End If
    End If
    
    frmMain.lblStatus.Caption = IIf(Complete, "Script Executed.", "Script stopped.")
End Sub

Public Function ParseScript(Window As Form, Script As Collection, strEventName As String, Optional EventParams, Optional strSubName As String, Optional SubParams) As Boolean
    On Error GoTo Handler
    
    Dim CmdParams() As String, Param As String, i As Integer, j As Integer, strTemp As String, frmTemp As Form
    
    If strSubName = vbNullString Then
        CmdParams = Split(Script("script_" & strEventName).Item(Script("pos_" & strEventName)))
    Else
        CmdParams = Split(Script("script_" & strSubName).Item(Script("pos_" & strSubName)))
    End If
    
    If UBound(CmdParams) > 0 Then
        If strSubName = vbNullString Then
            Param = Right$(Script("script_" & strEventName).Item(Script("pos_" & strEventName)), Len(Script("script_" & strEventName).Item(Script("pos_" & strEventName))) - InStr(Script("script_" & strEventName).Item(Script("pos_" & strEventName)), " "))
        Else
            Param = Right$(Script("script_" & strSubName).Item(Script("pos_" & strSubName)), Len(Script("script_" & strSubName).Item(Script("pos_" & strSubName))) - InStr(Script("script_" & strSubName).Item(Script("pos_" & strSubName)), " "))
        End If

        Select Case strEventName
        Case "statuschanged"
            Param = Replace$(Param, "$status", EventParams(0))
        Case "nickchanged"
            Param = Replace$(Param, "$nick", EventParams(0))
        Case "contactstatuschanged"
            Param = Replace$(Param, "$email", EventParams(0))
            Param = Replace$(Param, "$nick", EventParams(1))
            Param = Replace$(Param, "$status", EventParams(2))
        Case "contactnickchanged"
            Param = Replace$(Param, "$email", EventParams(0))
            Param = Replace$(Param, "$nick", EventParams(1))
        Case "imwindowopened"
            Param = Replace$(Param, "$email", EventParams(0))
            Param = Replace$(Param, "$nick", EventParams(1))
        Case "messagereceived"
            Param = Replace$(Param, "$email", EventParams(0))
            Param = Replace$(Param, "$nick", EventParams(1))
            Param = Replace$(Param, "$fontname", EventParams(2))
            Param = Replace$(Param, "$fontcolor", EventParams(3))
            Param = Replace$(Param, "$fontbold", EventParams(4))
            Param = Replace$(Param, "$fontitalic", EventParams(5))
            Param = Replace$(Param, "$fontstrikethru", EventParams(6))
            Param = Replace$(Param, "$fontunderline", EventParams(7))
            Param = Replace$(Param, "$message", EventParams(8))
        Case "messagesent"
            Param = Replace$(Param, "$email", EventParams(0))
            Param = Replace$(Param, "$fontname", EventParams(1))
            Param = Replace$(Param, "$fontcolor", EventParams(2))
            Param = Replace$(Param, "$fontbold", EventParams(3))
            Param = Replace$(Param, "$fontitalic", EventParams(4))
            Param = Replace$(Param, "$fontstrikethru", EventParams(5))
            Param = Replace$(Param, "$fontunderline", EventParams(6))
            Param = Replace$(Param, "$message", EventParams(7))
        Case "transferrequest"
            Param = Replace$(Param, "$email", EventParams(0))
            Param = Replace$(Param, "$filename", EventParams(1))
            Param = Replace$(Param, "$filesize", EventParams(2))
            Param = Replace$(Param, "$requestid", EventParams(3))
        Case "transferaccepted", "transfercancelled"
            Param = Replace$(Param, "$email", EventParams(0))
            Param = Replace$(Param, "$requestid", EventParams(1))
        Case "transfercomplete"
            Param = Replace$(Param, "$requestid", EventParams(0))
        Case "imwindowclosed"
            Param = Replace$(Param, "$email", EventParams(0))
        End Select
        
        If Window.Name = "frmChat" Then
            Param = Alias(Param, Window)
        Else
            Param = Alias(Param)
        End If
        
        If Not (CmdParams(0) = "set" Or CmdParams(0) = "destroy") Then
            If Not strSubName = vbNullString Then
                If Not IsMissing(SubParams) Then
                    Param = Replace$(Param, "$args", UBound(SubParams) + 1)
                    For i = 0 To UBound(SubParams)
                        Param = Replace$(Param, "$arg[" & i + 1 & "]", TrimX(CStr(SubParams(i))))
                    Next
                Else
                    Param = Replace$(Param, "$args", 0)
                End If
            End If
            Param = MixGmsVars(Script, Param)
        ElseIf CmdParams(0) = "set" Then
            Dim VarName As String, VarVal As String
            VarName = Trim$(Left$(Param, InStr(Param, "=") - 1))
            strTemp = SubString(VarName, "[", "]")
            If Not strTemp = vbNullString Then
                VarName = Replace$(VarName, strTemp, MixGmsVars(Script, strTemp))
            End If
            VarVal = GetString(Right$(Param, Len(Param) - InStr(Param, "=")))
            VarVal = MixGmsVars(Script, VarVal)
            If Not strSubName = vbNullString Then
                If Not IsMissing(SubParams) Then
                    VarVal = Replace$(VarVal, "$args", UBound(SubParams) + 1)
                    For i = 0 To UBound(SubParams)
                        VarVal = Replace$(VarVal, "$arg[" & i + 1 & "]", TrimX(CStr(SubParams(i))))
                    Next
                Else
                    VarVal = Replace$(VarVal, "$args", 0)
                End If
            End If
            Param = VarName & "=" & VarVal
        End If
        
        i = InStr(Param, "(inp)")
        Dim strInpParam As String, intInpParamCount As Integer, InpCount As Integer
        InpCount = WordCount(Param, "(inp)")
        intInpParamCount = 1
        While Not i = 0
            If CmdParams(0) = "set" Then
                If InpCount = 1 Then
                    strInpParam = InputBox("Enter the value for '" & Left$(Param, InStr(Param, "=") - 1) & "' variable.", "GM Scripting Engine")
                Else
                    strInpParam = InputBox("Enter the value no." & intInpParamCount & " for '" & Left$(Param, InStr(Param, "=") - 1) & "' variable.", "GM Scripting Engine")
                End If
            Else
                If InpCount = 1 Then
                    strInpParam = InputBox("Enter the value for '" & CmdParams(0) & "' command parameter.", "GM Scripting Engine")
                Else
                    strInpParam = InputBox("Enter the value no." & intInpParamCount & " for '" & CmdParams(0) & "' command parameter.", "GM Scripting Engine")
                End If
            End If
            Param = Replace$(Param, "(inp)", strInpParam, , 1)
            i = InStr(i + 1, Param, "(inp)")
        Wend
        CmdParams = Split(CmdParams(0) & " " & Param)
    End If
    
    Select Case CmdParams(0)
    Case "call"
        If InCollection(Script, "script_" & CmdParams(1)) Then
            Call SetCollectionItem(Script, "pos_" & CmdParams(1), 1)
            Do Until Script.Item("pos_" & CmdParams(1)) > Script.Item("script_" & CmdParams(1)).Count
                DoEvents
                If UBound(CmdParams) >= 2 Then
                    Call ParseScript(Window, Script, strEventName, EventParams, CmdParams(1), Split(Mid$(Param, InStr(Param, " ") + 1), ","))
                Else
                    Call ParseScript(Window, Script, strEventName, EventParams, CmdParams(1))
                End If
                If Script.Count = 0 Then 'To check if the script has not ended before the qued script statement is executed
                    Exit Do
                End If
            Loop
        Else
            Call ErrorAlert(Script, "Script Error", "Subroutine named " & CmdParams(1) & " not found!", Script("pos_" & strEventName) + 1)
        End If
    Case "if"
        If Not LogicCalc(Param) Then
            If strSubName = vbNullString Then
                Call SkipIf(Script, strEventName, Script("pos_" & strEventName))
            Else
                Call SkipIf(Script, strSubName, Script("pos_" & strSubName))
            End If
        End If
    Case "set"
        VarName = Trim$(Left$(Param, InStr(Param, "=") - 1))
        If VarVal Like "fopen(*,?)" Then
            strTemp = Mid$(VarVal, InStr(VarVal, "(") + 1)
            strTemp = Left$(strTemp, Len(strTemp) - 1)
            VarVal = FreeFile
            Script("hfiles").Add VarVal
            i = InStrRev(strTemp, ",")
            Select Case TrimX(Mid$(strTemp, i + 1))
            Case "i"
                Open Left$(strTemp, i - 1) For Input As #CInt(VarVal)
            Case "o"
                MakeSureDirectoryPathExists Left$(strTemp, i - 1)
                Open Left$(strTemp, i - 1) For Output As #CInt(VarVal)
            Case "a"
                MakeSureDirectoryPathExists Left$(strTemp, i - 1)
                Open Left$(strTemp, i - 1) For Append As #CInt(VarVal)
            Case "b"
                MakeSureDirectoryPathExists Left$(strTemp, i - 1)
                Open Left$(strTemp, i - 1) For Binary As #CInt(VarVal)
            Case "r"
                MakeSureDirectoryPathExists Left$(strTemp, i - 1)
                Open Left$(strTemp, i - 1) For Random As #CInt(VarVal)
            End Select
        ElseIf VarVal Like "fread(*,*,*)" Then
            Dim TempAry() As String
            TempAry = Split(Mid$(VarVal, 7, Len(VarVal) - 7), ",")
            strTemp = Space$(CInt(TempAry(2)))
            Get #CInt(TempAry(0)), TempAry(1), strTemp
            VarVal = strTemp
        ElseIf VarVal Like "fread(*)" Then
            i = InStr(VarVal, "(")
            Line Input #CInt(Mid$(VarVal, 7, Len(VarVal) - 7)), strTemp
            VarVal = strTemp
        Else
            Dim VarValParams() As String
            VarValParams = Split(VarVal)
            If UBound(VarValParams) >= 1 Then
                If VarValParams(0) = "sendfile" Then
                    If IsEmail(VarValParams(1)) And UBound(VarValParams) > 1 Then
                        VarVal = StartChat(VarValParams(1), , "|" & Join(SubArray(VarValParams, 2, UBound(VarValParams))))
                    ElseIf Window.Name = "frmChat" Then
                        VarVal = SendFile(Window, Mid$(VarVal, InStr(VarVal, " ") + 1))
                    End If
                ElseIf VarValParams(0) = "filelen" And IsNumeric(VarValParams(1)) And UBound(VarValParams) = 1 Then
                    VarVal = LOF(VarValParams(1))
                ElseIf VarValParams(0) = "mid" And IsNumeric(VarValParams(1)) And UBound(VarValParams) > 1 Then
                    If IsNumeric(VarValParams(2)) And UBound(VarValParams) > 2 Then
                        VarVal = Mid$(VarVal, 4 + Len(VarValParams(1)) + 1 + Len(VarValParams(2)) + 1 + Val(VarValParams(1)), Val(VarValParams(2)))
                    Else
                        VarVal = Mid$(VarVal, 4 + Len(VarValParams(1)) + 1 + Val(VarValParams(1)))
                    End If
                ElseIf VarValParams(0) = "left" And IsNumeric(VarValParams(1)) And UBound(VarValParams) > 1 Then
                    VarVal = Left$(Right$(VarVal, Len(VarVal) - 5 - Len(VarValParams(1)) - 1), Val(VarValParams(1)))
                ElseIf VarValParams(0) = "right" And IsNumeric(VarValParams(1)) And UBound(VarValParams) > 1 Then
                    VarVal = Right$(Right$(VarVal, Len(VarVal) - 6 - Len(VarValParams(1)) - 1), Val(VarValParams(1)))
                ElseIf VarValParams(0) = "len" Then
                    VarVal = Len(VarVal) - 4
                ElseIf VarValParams(0) = "comment" And IsEmail(VarValParams(1)) And UBound(VarValParams) = 1 Then
                    VarVal = GetBuddyComment(VarValParams(1))
                ElseIf VarValParams(0) = "customnick" And IsEmail(VarValParams(1)) And UBound(VarValParams) = 1 Then
                    VarVal = GetBuddyCustomNick(VarValParams(1))
                ElseIf VarValParams(0) = "nick" And IsEmail(VarValParams(1)) And UBound(VarValParams) = 1 Then
                    VarVal = GetContactAttr(VarValParams(1), "nick")
                ElseIf VarValParams(0) = "status" And IsEmail(VarValParams(1)) And UBound(VarValParams) = 1 Then
                    VarVal = GetContactAttr(VarValParams(1), "status")
                ElseIf VarValParams(0) = "rand" And IsNumeric(VarValParams(1)) And UBound(VarValParams) = 1 Then
                    Randomize Timer
                    VarVal = Fix(Rnd * Val(VarValParams(1))) + 1
                ElseIf VarValParams(0) = "inlist" And IsEmail(VarValParams(1)) And UBound(VarValParams) = 1 Then
                    If InCollection(ContactList, VarValParams(1)) Then
                        VarVal = InList(GetContactAttr(VarValParams(1), "lists"), msnList_Forward)
                    Else
                        VarVal = False
                    End If
                ElseIf VarValParams(0) = "inlist" And UBound(VarValParams) = 2 Then
                    If IsEmail(VarValParams(2)) Then
                        If InCollection(ContactList, VarValParams(2)) Then
                            VarVal = InList(GetContactAttr(VarValParams(2), "lists"), Val(ListCode(VarValParams(1))))
                        Else
                            VarVal = False
                        End If
                    End If
                Else
                    VarVal = GetString(Right$(Param, Len(Param) - InStr(Param, "=")))
                    VarVal = MathCalc(VarVal)
                End If
            Else
                VarVal = GetString(Right$(Param, Len(Param) - InStr(Param, "=")))
                VarVal = MathCalc(VarVal)
            End If
        End If
        
        Call AddScriptVar(Script, VarName, VarVal)
        
    Case "fwrite"
        If UBound(CmdParams) >= 3 Then
            If IsNumeric(CmdParams(1)) And IsNumeric(CmdParams(2)) Then
                j = 0
                For i = 1 To 2
                    j = InStr(j + 1, Param, " ")
                Next
                Put #CInt(CmdParams(1)), CmdParams(2), Right$(Param, Len(Param) - j)
            Else
                Print #CInt(CmdParams(1)), Right$(Param, Len(Param) - InStr(Param, " "));
            End If
        Else
            Print #CInt(CmdParams(1)), Right$(Param, Len(Param) - InStr(Param, " "));
        End If
    Case "fclose"
        Close #CInt(CmdParams(1))
    Case "destroy"
        If InCollection(Script("vars"), Param) Then
            Script("vars").Remove Param
        End If
    Case "accept"
        If Not Param = vbNullString Then
            For Each frmTemp In Forms
                If frmTemp.Name = "frmTransfer" Then
                    If frmTemp.Cookie = CmdParams(1) Then
                        If UBound(CmdParams) >= 2 Then
                            frmTemp.objMSN_FTP.FilePath = Replace$(Mid$(Param, InStr(Param, " ") + 1) & "\", "\\", "\") & frmTemp.objMSN_FTP.File
                        End If
                        Call frmTemp.cmdAccept_Click
                    End If
                    Exit For
                End If
            Next
        End If
    Case "cancel"
        If Not Param = vbNullString Then
            For Each frmTemp In Forms
                If frmTemp.Name = "frmTransfer" Then
                    If frmTemp.Cookie = CmdParams(1) Then
                        Call frmTemp.cmdCancel_Click
                    End If
                End If
            Next
        End If
    Case "playsound"
        Call PlaySound(Param, "script")
    Case "changenick"
        Call frmMain.objMSN_NS.ChangeNick(Param)
    Case "goto"
        If IsNumeric(Param) Then
            If strSubName = vbNullString Then
                SetCollectionItem Script, "pos_" & strEventName, Val(Param) - 1
            Else
                SetCollectionItem Script, "pos_" & strSubName, Val(Param) - 1
            End If
        ElseIf InCollection(Script("labels"), Param) Then
            If strSubName = vbNullString Then
                SetCollectionItem Script, "pos_" & strEventName, Script("labels").Item(Param) - 1
            Else
                SetCollectionItem Script, "pos_" & strEventName, Script("labels").Item(Param) - 1
            End If
        Else
            If strSubName = vbNullString Then
                Call ErrorAlert(Script, "Script Error", "Undefined label used in Goto statement", Script("script_" & strEventName).Item(Script("pos_" & strEventName)))
            Else
                Call ErrorAlert(Script, "Script Error", "Undefined label used in Goto statement", Script("script_" & strSubName).Item(Script("pos_" & strSubName)))
            End If
        End If
    Case "changestatus"
        Call frmMain.objMSN_NS.ChangeStatus(Val(CmdParams(1)))
    Case "signout"
        Call Signout
    Case "signin"
        Call SignIn(CmdParams(1), CmdParams(2))
    Case "sleep"
        If strEventName = "main" Then
            SetCollectionItem Script, "sleep", Val(Param) / 1000
        End If
    Case "chat"
        Call StartChat(CmdParams(1))
    Case "msg"
        Call StartChat(CmdParams(1), Right$(Param, Len(Param) - InStr(Param, " ")))
    Case "msgex"
        If Window.Name = "frmChat" Then
            SendMsg Window, Param
        End If
    Case "msgbox"
        MsgBox Param, , "GM Script Message"
    Case "sendfile"
        If IsEmail(CmdParams(1)) And UBound(CmdParams) > 1 Then
            Call StartChat(CmdParams(1), , "|" & Mid$(Param, InStr(Param, " ") + 1))
        ElseIf Window.Name = "frmChat" Then
            Call SendFile(Window, "|" & Param)
        End If
    Case "msgall"
        Call MessageAll(Param)
    Case "ignore"
        Call IgnoreContact(CmdParams(1))
    Case "unignore"
        Call UnignoreContact(CmdParams(1))
    Case "block"
        Call BlockContact(CmdParams(1))
    Case "unblock"
        Call UnblockContact(CmdParams(1))
    Case "addcontact"
        If IsEmail(CmdParams(1)) Then
            Call AddContact(CmdParams(1))
        ElseIf UBound(CmdParams) = 2 Then
            If IsEmail(CmdParams(2)) Then
                Call frmMain.objMSN_NS.AddContact(Val(ListCode(CmdParams(1))), CmdParams(2))
            End If
        End If
    Case "delcontact"
        If IsEmail(CmdParams(1)) Then
            Call frmMain.objMSN_NS.RemoveContact(msnList_Forward, CmdParams(1))
        ElseIf UBound(CmdParams) = 2 Then
            If IsEmail(CmdParams(2)) Then
                Call frmMain.objMSN_NS.RemoveContact(Val(ListCode(CmdParams(1))), CmdParams(2))
            End If
        End If
    Case "comment"
        If UBound(CmdParams) > 1 Then
            Call SetBuddyComment(CmdParams(1), Right$(Param, Len(Param) - InStr(Param, " ")))
        ElseIf UBound(CmdParams) = 1 Then
            Call SetBuddyComment(CmdParams(1), vbNullString)
        End If
    Case "customnick"
        If UBound(CmdParams) > 1 Then
            Call SetBuddyCustomNick(CmdParams(1), Right$(Param, Len(Param) - InStr(Param, " ")))
        ElseIf UBound(CmdParams) = 1 Then
            Call SetBuddyCustomNick(CmdParams(1), vbNullString)
        End If
    Case "loadscript"
        Dim ScriptParams() As String
        ScriptParams = Split(Param, ",")
        For i = 0 To UBound(ScriptParams)
            ScriptParams(i) = TrimX(ScriptParams(i))
        Next
        If UBound(ScriptParams) = 0 Then
            Call ToggleScript(Param, True)
        Else
            Call ToggleScript(ScriptParams(0), True, SubArray(ScriptParams, 1, UBound(ScriptParams)))
        End If
    Case "loadbot"
        Call LoadBot(Param)
    Case "execute"
        Call ShellExecuteEx(Param)
    Case "end"
        If UBound(CmdParams) = 0 Then
            Call EndScript(Script, True)
            Exit Function
        End If
    Case "exit"
        Call TerminateGM
    Case "else"
        Dim ElseCount As Integer
        If strSubName = vbNullString Then
            For i = Script("pos_" & strEventName) + 1 To Script("script_" & strEventName).Count
                If Script("script_" & strEventName).Item(i) = "end if" Then
                    ElseCount = ElseCount - 1
                ElseIf Split(Script("script_" & strEventName).Item(i))(0) = "if" Then
                    ElseCount = ElseCount + 1
                End If
                If ElseCount = -1 Then
                    Exit For
                End If
            Next
            SetCollectionItem Script, "pos_" & strEventName, i - 1
        Else
            For i = Script("pos_" & strSubName) + 1 To Script("script_" & strSubName).Count
                If Script("script_" & strSubName).Item(i) = "end if" Then
                    ElseCount = ElseCount - 1
                ElseIf Split(Script("script_" & strSubName).Item(i))(0) = "if" Then
                    ElseCount = ElseCount + 1
                End If
                If ElseCount = -1 Then
                    Exit For
                End If
            Next
            SetCollectionItem Script, "pos_" & strSubName, i - 1
        End If
    Case Else
        If strSubName = vbNullString Then
            Call ErrorAlert(Script, "Script Error", "Unrecognized statement.", Script("script_" & strEventName).Item(Script("pos_" & strEventName)))
        Else
            Call ErrorAlert(Script, "Script Error", "Unrecognized statement.", Script("script_" & strSubName).Item(Script("pos_" & strSubName)))
        End If
        Exit Function
    End Select
    
    If Not Script.Count = 0 Then 'To check if the script has not ended before the qued script statement is executed
        If strSubName = vbNullString Then
            SetCollectionItem Script, "pos_" & strEventName, Script("pos_" & strEventName) + 1
        Else
            SetCollectionItem Script, "pos_" & strSubName, Script("pos_" & strSubName) + 1
        End If
    End If
    
    ParseScript = True
    Exit Function
Handler:
    Call ErrorAlert(Script, "Unexpected Error.", Err.Description)
End Function

Private Sub SkipIf(Script As Collection, strEventName As String, GMSPos As Integer)
    Dim i As Integer, IfCount As Integer
    For i = GMSPos + 1 To Script("script_" & strEventName).Count
        If Script("script_" & strEventName).Item(i) = "end if" Or (Script("script_" & strEventName).Item(i) = "else" And IfCount = 0) Then
            IfCount = IfCount - 1
        ElseIf Split(Script("script_" & strEventName).Item(i))(0) = "if" Then
            IfCount = IfCount + 1
        End If
        If IfCount = -1 Then
            Exit For
        End If
    Next
    If i <= Script("script_" & strEventName).Count Then
        If Script("script_" & strEventName).Item(i) = "else" Then
            SetCollectionItem Script, "pos_" & strEventName, i
        Else
            SetCollectionItem Script, "pos_" & strEventName, i - 1
        End If
    Else
        SetCollectionItem Script, "pos_" & strEventName, i - 1
    End If
End Sub

Private Function MixGmsVars(Script As Collection, Text As String) As String
    On Error Resume Next

    Dim Temp As String
    Temp = SubString(Text, "[", "]")
    If Temp = vbNullString Then
        MixGmsVars = Text
    Else
        MixGmsVars = Replace$(Text, Temp, MixGmsVars(Script, Temp))
    End If
    
    Temp = Script("key")
    MixGmsVars = Replace$(MixGmsVars, "$installdir", App.Path)
    MixGmsVars = Replace$(MixGmsVars, "$scriptpath", Left$(Temp, InStrRev(Temp, "\") - 1))
    MixGmsVars = Replace$(MixGmsVars, "$scriptname", Right$(Temp, Len(Temp) - InStrRev(Temp, "\")))
    MixGmsVars = Replace$(MixGmsVars, "$idletime", GetIdleTime)
    Dim i As Integer
    With Script("vars")
        For i = 1 To .Count
            MixGmsVars = Replace$(MixGmsVars, .Item(i).Item("name"), .Item(i).Item("value"))
        Next
    End With
    Temp = SubString(MixGmsVars, "eof(", ")")
    If Not Temp = vbNullString Then
        MixGmsVars = Replace$(MixGmsVars, "eof(" & Temp & ")", EOF(Val(Temp)))
    End If
    MixGmsVars = Replace$(MixGmsVars, "$null", vbNullString)
    MixGmsVars = Replace$(MixGmsVars, "$listcontacts", ContactList.Count)
    If IsNumeric(SubString(MixGmsVars, "$listcontact[", "]")) Then
        For i = 1 To ContactList.Count
            MixGmsVars = Replace$(MixGmsVars, "$listcontact[" & i & "]", ContactList(i).Item("email"))
        Next
    End If
End Function

Private Sub ErrorAlert(Script As Collection, ErrorType As String, Description As String, Optional Statement As String)
    If Not Statement = vbNullString Then
        MsgBox "Error: " & ErrorType & vbCrLf & vbCrLf & Description & vbCrLf & vbCrLf & "Statement: " & Statement, vbExclamation, "GM Scripting Engine"
    Else
        MsgBox "Error: " & ErrorType & vbCrLf & vbCrLf & Description, vbExclamation, "GM Scripting Engine"
    End If
    Call EndScript(Script)
End Sub

Public Function MathCalc(Expression As String) As String
    If Expression = vbNullString Then
        Exit Function
    Else
        MathCalc = Expression
    End If

    Dim SubExpression As String
    SubExpression = SubString(Expression, "(", ")")
    If Not SubExpression = vbNullString Then
        Expression = Replace$(Expression, "(" & SubExpression & ")", MathCalc(SubExpression))
    End If
    
    Dim Operand1 As String, Operand2 As String
    Dim Operator As String, OperatorPos As Integer, OperatorLenDiff As Integer
    
    OperatorPos = InStr(Expression, "%")
    If OperatorPos = 0 Then
        OperatorPos = InStr(Expression, "-")
        If OperatorPos = 0 Then
            OperatorPos = InStr(Expression, "+")
            If OperatorPos = 0 Then
                OperatorPos = InStr(Expression, "/")
                If OperatorPos = 0 Then
                    OperatorPos = InStr(Expression, "*")
                    If OperatorPos = 0 Then
                        OperatorPos = InStr(Expression, "^")
                        If OperatorPos = 0 Then
                            Exit Function
                        Else
                            Operator = "^"
                        End If
                    Else
                        Operator = "*"
                    End If
                Else
                    Operator = "/"
                End If
            Else
                Operator = "+"
            End If
        Else
            Operator = "-"
        End If
    Else
        Operator = "%"
    End If
    
    Operand1 = MathCalc(GetString(Left$(Expression, OperatorPos - 1)))
    Operand2 = MathCalc(GetString(Right$(Expression, Len(Expression) - OperatorPos - OperatorLenDiff)))
    
    If (IsNumeric(Operand1) Or (Operand1 = vbNullString And (Operator = "-" Or Operator = "+"))) And IsNumeric(Operand2) Then
        Select Case Operator
        Case "^"
            MathCalc = Val(Operand1) ^ Val(Operand2)
        Case "*"
            MathCalc = Val(Operand1) * Val(Operand2)
        Case "/"
            MathCalc = Val(Operand1) / Val(Operand2)
        Case "+"
            MathCalc = Val(Operand1) + Val(Operand2)
        Case "-"
            MathCalc = Val(Operand1) - Val(Operand2)
        Case "%"
            MathCalc = Val(Operand1) Mod Val(Operand2)
        End Select
    End If
End Function

Public Function LogicCalc(Expression) As Boolean
    If Expression = vbNullString Then
        Exit Function
    End If
    
    Dim Operand1 As String, Operand2 As String
    Dim Operator As String, OperatorPos As Integer, OperatorLenDiff As Integer
    
    OperatorPos = InStr(Expression, "<>")
    If OperatorPos = 0 Then
        OperatorPos = InStr(Expression, ">=")
        If OperatorPos = 0 Then
            OperatorPos = InStr(Expression, "<=")
            If OperatorPos = 0 Then
                OperatorPos = InStr(Expression, "=")
                If OperatorPos = 0 Then
                    OperatorPos = InStr(Expression, ">")
                    If OperatorPos = 0 Then
                        OperatorPos = InStr(Expression, "<")
                        If OperatorPos = 0 Then
                            OperatorPos = InStr(Expression, "like")
                            If OperatorPos = 0 Then
                                LogicCalc = Expression
                                Exit Function
                            Else
                                Operator = "like"
                                OperatorLenDiff = 3
                            End If
                        Else
                            Operator = "<"
                        End If
                    Else
                        Operator = ">"
                    End If
                Else
                    Operator = "="
                End If
            Else
                Operator = "<="
                OperatorLenDiff = 1
            End If
        Else
            Operator = ">="
            OperatorLenDiff = 1
        End If
    Else
        Operator = "<>"
        OperatorLenDiff = 1
    End If
    
    Operand1 = MathCalc(GetString(Left$(Expression, OperatorPos - 1)))
    Operand2 = MathCalc(GetString(Right$(Expression, Len(Expression) - OperatorPos - OperatorLenDiff)))
    
    Select Case Operator
    Case "="
        LogicCalc = (Operand1 = Operand2)
    Case "<>"
        LogicCalc = (Operand1 <> Operand2)
    Case ">="
        If IsNumeric(Operand1) And IsNumeric(Operand2) Then
            LogicCalc = (Val(Operand1) >= Val(Operand2))
        Else
            LogicCalc = (Operand1 >= Operand2)
        End If
    Case "<="
        If IsNumeric(Operand1) And IsNumeric(Operand2) Then
            LogicCalc = (Val(Operand1) <= Val(Operand2))
        Else
            LogicCalc = (Operand1 <= Operand2)
        End If
    Case ">"
        If IsNumeric(Operand1) And IsNumeric(Operand2) Then
            LogicCalc = (Val(Operand1) > Val(Operand2))
        Else
            LogicCalc = (Operand1 > Operand2)
        End If
    Case "<"
        If IsNumeric(Operand1) And IsNumeric(Operand2) Then
            LogicCalc = (Val(Operand1) < Val(Operand2))
        Else
            LogicCalc = (Operand1 < Operand2)
        End If
    Case "like"
        LogicCalc = (Operand1 Like Operand2)
    End Select
End Function

Public Sub ToggleScript(ScriptPath As String, Optional Reload As Boolean, Optional Params)
    Dim strScript As String, i As Integer
    strScript = LCase$(ScriptPath)
    
    For i = 1 To frmMain.mnuTools_GMScript_Script.UBound
        If LCase$(frmMain.mnuTools_GMScript_Script(i).Caption) = strScript Or frmMain.mnuTools_GMScript_Script(i).Tag = strScript Then
            frmMain.mnuTools_GMScript_Script(i).Checked = Not frmMain.mnuTools_GMScript_Script(i).Checked
            If frmMain.mnuTools_GMScript_Script(i).Checked Then
                If IsMissing(Params) Then
                    Call LoadScript(frmMain.mnuTools_GMScript_Script(i).Tag)
                Else
                    Call LoadScript(frmMain.mnuTools_GMScript_Script(i).Tag, Params)
                End If
            Else
                Call EndScript(GMScripts(frmMain.mnuTools_GMScript_Script(i).Tag))
                If Reload Then
                    If IsMissing(Params) Then
                        Call LoadScript(frmMain.mnuTools_GMScript_Script(i).Tag)
                    Else
                        Call LoadScript(frmMain.mnuTools_GMScript_Script(i).Tag, Params)
                    End If
                End If
            End If
            Exit Sub
        End If
    Next
    
    For i = 1 To frmMain.mnuTools_GMScript_Other.UBound
        If LCase$(frmMain.mnuTools_GMScript_Other(i).Caption) = strScript Or frmMain.mnuTools_GMScript_Other(i).Tag = strScript Then
            Dim ScriptFile As String
            ScriptFile = frmMain.mnuTools_GMScript_Other(i).Tag
            If Reload Then
                Call EndScript(GMScripts(ScriptFile), , False)
                If IsMissing(Params) Then
                    Call LoadScript(ScriptFile)
                Else
                    Call LoadScript(ScriptFile, Params)
                End If
            Else
                Call EndScript(GMScripts(ScriptFile))
            End If
            Exit Sub
        End If
    Next
    
    If FileExists(strScript) Then
        Dim strFileTitle As String
        strFileTitle = Right$(strScript, Len(strScript) - InStrRev(strScript, "\"))
        If InStr(strFileTitle, ".") > 0 Then
            strFileTitle = Left$(strFileTitle, InStr(strFileTitle, ".") - 1)
        End If
        Load frmMain.mnuTools_GMScript_Other(frmMain.mnuTools_GMScript_Other.UBound + 1)
        frmMain.mnuTools_GMScript_Other(frmMain.mnuTools_GMScript_Other.UBound).Tag = strScript
        frmMain.mnuTools_GMScript_Other(frmMain.mnuTools_GMScript_Other.UBound).Caption = strFileTitle
        frmMain.mnuTools_GMScript_Other(frmMain.mnuTools_GMScript_Other.UBound).Checked = True
        If IsMissing(Params) Then
            Call LoadScript(strScript)
        Else
            Call LoadScript(strScript, Params)
        End If
    End If
End Sub

Public Sub QueScript(Source As Form, strEvent As String, Optional Params)
    On Error GoTo Handler
    
    Dim PrevCount As Integer
    PrevCount = ScriptQue.Count
    Dim ScriptIndex As Integer, ScriptLet As Collection, i As Integer
    For ScriptIndex = 1 To GMScripts.Count
        If InCollection(GMScripts(ScriptIndex), "script_" & strEvent) Then
            With GMScripts(ScriptIndex)
                If .Item("script_" & strEvent).Count > 0 Then
                    Set ScriptLet = New Collection
                    ScriptLet.Add Source, "source"
                    ScriptLet.Add ScriptIndex, "script"
                    ScriptLet.Add strEvent, "event"
                    ScriptLet.Add Params, "params"
                    ScriptQue.Add ScriptLet
                    Set ScriptLet = Nothing
                End If
            End With
        End If
    Next
    If ScriptQue.Count > 0 And PrevCount = 0 And Not frmMain.tmrGMScript_Events.Enabled Then
        frmMain.tmrGMScript_Events.Enabled = True
    End If
Handler:
End Sub

Private Sub AddScriptVar(Script As Collection, VarName As String, VarVal As String)
    'Adding the variable to script variable table
    
    Dim NewVar As Collection
    Set NewVar = New Collection
    NewVar.Add VarName, "name"
    NewVar.Add VarVal, "value"
    
    SetCollectionItem Script("vars"), VarName, NewVar
    
    'Sorting script variable table
    
    Dim i As Integer, j As Integer
    If Script("vars").Count > 1 Then
        Dim colSorted As Collection
        Set colSorted = New Collection
        Dim SortedKey As String
        For i = 1 To Script("vars").Count
            SortedKey = Script("vars")(1).Item("name")
            For j = 2 To Script("vars").Count
                If SortedKey < Script("vars")(j).Item("name") Then
                    SortedKey = Script("vars")(j).Item("name")
                End If
            Next
            colSorted.Add Script("vars").Item(SortedKey), SortedKey
            Script("vars").Remove SortedKey
        Next
        SetCollectionItem Script, "vars", colSorted
        Set colSorted = Nothing
    End If
End Sub

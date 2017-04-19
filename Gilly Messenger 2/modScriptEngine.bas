Attribute VB_Name = "modScriptEngine"
Dim Scripts() As New Collection, ScriptMap0() As Integer, ScriptMap1() As Integer, TmpCount As Integer

Public Sub ExecuteScript()
frmMain.tmrGMScript.Enabled = False
frmMain.tmrGMScript.Tag = "Running"
Call ProcessGMS
GMSCount = GMSCount + 1
If GMSCount > UBound(ScriptMap0) Then
    Call EndScript
ElseIf frmMain.tmrGMScript.Tag <> vbNullString Then
    frmMain.tmrGMScript.Enabled = True
End If
End Sub

Public Sub ProcessGMS()
On Error GoTo Handler
Dim CMD As String, Parm As String, ScriptLine As String
ScriptLine = Scripts(ScriptMap0(GMSCount)).Item(ScriptMap1(GMSCount))
For i = 1 To Len(ScriptLine)
    X = InStr(i, ScriptLine, "VAL:", vbTextCompare)
    If X > 0 Then
        Temp = Split(Right(ScriptLine, Len(ScriptLine) - X + 1))(0)
        Temp = Right(Temp, Len(Temp) - 4)
        ScriptLine = Replace(ScriptLine, "VAL:" & Temp, GetVal(Temp), , , vbTextCompare)
    End If
Next
If UCase$(ScriptLine) Like "IF * IS *" Then
    If UCase$(Left$(ScriptLine, 3)) = "IF " And UCase$(Split(ScriptLine)(2)) = "IS" Then
        Dim IfType As String, IfVal As String, ifNot As Boolean
        IfVal = Split(ScriptLine)(1)
        If UCase$(Split(ScriptLine)(3)) <> "NOT" Then
            IfType = UCase$(Split(ScriptLine)(3))
            ifNot = False
        Else
            IfType = UCase$(Split(ScriptLine)(4))
            ifNot = True
        End If
        Select Case IfType
        Case "ONLINE"
            If ifNot = False Then
                If IsOnline(IfVal) = False Then
                    PassScript ScriptMap0(GMSCount)
                End If
            Else
                If Not (IsOnline(IfVal) = False) Then
                    PassScript ScriptMap0(GMSCount)
                End If
            End If
        Case "BLOCKED"
            If ifNot = False Then
                If GetBuddyBlock(IfVal) = "" Then
                    PassScript ScriptMap0(GMSCount)
                End If
            Else
                If Not (GetBuddyBlock(IfVal) = "") Then
                    PassScript ScriptMap0(GMSCount)
                End If
            End If
        Case "IGNORED"
            If ifNot = False Then
                If IsIgnored(IfVal) = False Then
                    PassScript ScriptMap0(GMSCount)
                End If
            Else
                If Not (IsIgnored(IfVal) = False) Then
                    PassScript ScriptMap0(GMSCount)
                End If
            End If
        Case Else
            PassScript ScriptMap0(GMSCount)
        End Select
        Exit Sub
    End If
End If
CMD = Split(ScriptLine)(0)
CMD = UCase$(CMD)
If Len(ScriptLine) > Len(CMD) + 1 Then
    Parm = Right$(ScriptLine, Len(ScriptLine) - Len(CMD) - 1)
End If
If UCase$(Parm) = "(INP)" Then
    Parm = InputBox("Enter the parameters for " & CMD & " command.", "GM Script Engine")
ElseIf UCase$(Left$(Parm, 4)) = "VAL:" Then
    Parm = Right$(Parm, Len(Parm) - 4)
    Parm = GetVal(Parm)
End If
Parm = Alias(Parm)
Select Case CMD
Case "VAR"
    If UCase$(Right$(Parm, 5)) = "(INP)" Then
        Dim TmpVar As String
        TmpVar = InputBox("Enter the value for " & Left$(Parm, InStr(Parm, ":") - 1))
        If Parm = vbNullString Then GoTo Handler
        SetVar Left$(Parm, InStr(Parm, ":") - 1), TmpVar
    ElseIf UCase$(Right$(Parm, 6)) = "(NICK)" Then
        SetVar Left$(Parm, InStr(Parm, ":") - 1), Nick
    Else
        SetVar Left$(Parm, InStr(Parm, ":") - 1), Right$(Parm, Len(Parm) - InStr(Parm, ":"))
    End If
Case "SND"
    mciSendString "close GM_Sound", vbNullString, 0, 0
    mciSendString "open " & Chr(34) & Parm & Chr(34) & " alias GM_Sound", vbNullString, 0, 0
    mciSendString "play GM_Sound", vbNullString, 0, 0
Case "REN"
    ChangeNick Parm
Case "GTO"
    GMSCount = Val(Parm) - 2
Case "CST"
    ChangeStatus Val(Parm)
Case "SOT"
    Call Signout
Case "SIN"
    Call SignIn(Left$(Parm, InStr(Parm, " ") - 1), Right$(Parm, InStr(Parm, " ")))
Case "SLP"
    frmMain.tmrGMScript.Interval = Val(Parm)
Case "AMG"
    SetAutoMsg Parm
Case "MGA"
    MessageAll Parm
Case "CLG"
    Call frmMain.mnuChatLogger_Click
Case "MSG"
    Call MessageUser(Left$(Parm, InStr(Parm, " ") - 1), Right$(Parm, Len(Parm) - InStr(Parm, " ")))
Case "IGR"
    Call Ignore(Parm)
Case "UGR"
    Call Unignore(Parm)
Case "BLK"
    Call Block(Parm)
Case "UBK"
    Call UnBlock(Parm)
Case "END"
    If Parm = vbNullString Then
        Call EndScript
    End If
Case "EXT"
    Call frmMain.mnuExit_Click
Case Else
    GoTo Handler
End Select
Exit Sub
Handler:
If Err.Number <> 0 Then
    MsgBox "Error in GM Script.", vbCritical, "Error!"
End If
Call EndScript
End Sub

Public Sub EndScript()
frmMain.tmrGMScript.Enabled = False
frmMain.tmrGMScript.Tag = vbNullString
frmMain.tmrGMScript.Interval = 500
ResetCollection GMSVars
frmMain.mnuStopScript.Visible = False
frmMain.mnuGMScript.Visible = True
If frmMain.lblStatus.Caption = "Executing Script..." Then
    frmMain.lblStatus.Caption = "GM Script executed."
End If
End Sub

Public Sub LoadScript(File As String)
On Error GoTo Handler
Dim Data As String
Open File For Binary As #1
Data = String(LOF(1), vbNullChar)
Get #1, , Data
Close #1

Dim Lines() As String, CurScript As Integer, X As Integer
TmpCount = 0
Lines = Split(Data, vbCrLf)
ReDim Scripts(0)
For X = 0 To UBound(Lines)
    Lines(X) = LTrim$(Lines(X))
    Lines(X) = Replace$(Lines(X), Chr$(vbKeyTab), vbNullString)
    If Lines(X) <> vbNullString And Left$(Lines(X), 1) <> "'" Then
        If UCase$(Left$(Lines(X), 2)) = "IF" Then
            AddScript CurScript, Lines(X)
            ReDim Preserve Scripts(UBound(Scripts) + 1)
            CurScript = CurScript + 1
        ElseIf UCase$(Lines(X)) = "END IF" Then
            CurScript = CurScript - 1
            AddScript CurScript, Lines(X)
        Else
            AddScript CurScript, Lines(X)
        End If
    End If
Next

GMSCount = 0
frmMain.tmrGMScript.Enabled = True
Exit Sub
Handler:
Resume Here
Here:
MsgBox "Invalid GM Script file.", vbCritical, "Error!"
Call EndScript
End Sub

Public Sub SetVar(VarName As String, VarValue As Variant)
On Error Resume Next
GMSVars.Remove VarName
GMSVars.Add VarValue, VarName
End Sub

Public Function GetVal(VarName As String) As String
On Error Resume Next
GetVal = GMSVars(VarName)
End Function

Private Sub PassScript(PassToScript As Integer)
Do Until ScriptMap0(GMSCount + 1) = PassToScript
GMSCount = GMSCount + 1
If GMSCount > UBound(ScriptMap0) Then Exit Do
Loop
End Sub

Private Sub AddScript(DesScript As Integer, Line As String)
Scripts(DesScript).Add Line
AddScriptMap DesScript, Scripts(DesScript).Count
End Sub

Private Sub AddScriptMap(DesScript As Integer, DesScriptPos As Integer)
ReDim Preserve ScriptMap0(TmpCount)
ReDim Preserve ScriptMap1(TmpCount)
ScriptMap0(TmpCount) = DesScript
ScriptMap1(TmpCount) = DesScriptPos
TmpCount = TmpCount + 1
End Sub

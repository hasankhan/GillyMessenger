Attribute VB_Name = "modChatBot"
Option Compare Text
Option Explicit

Public ChatBot() As String

Public Sub LoadBot(File As String)
    On Error Resume Next
    
    If IsNumeric(frmMain.mnuTools_ChatBot.Tag) Then
        frmMain.mnuTools_ChatBot_Bot(Val(frmMain.mnuTools_ChatBot.Tag)).Checked = False
    Else
        frmMain.mnuTools_ChatBot_Other.Checked = False
    End If
    frmMain.mnuTools_ChatBot.Tag = vbNullString
    
    Erase ChatBot
    
    If Not File = vbNullString Then
        Dim i As Integer
        For i = 1 To frmMain.mnuTools_ChatBot_Bot.UBound
            If frmMain.mnuTools_ChatBot_Bot(i).Tag = File Or frmMain.mnuTools_ChatBot_Bot(i).Caption = File Then
                File = frmMain.mnuTools_ChatBot_Bot(i).Tag
                frmMain.mnuTools_ChatBot_Bot(i).Checked = True
                frmMain.mnuTools_ChatBot.Tag = i
                Exit For
            End If
        Next
        If i > frmMain.mnuTools_ChatBot_Bot.UBound Then
            frmMain.mnuTools_ChatBot_Other.Checked = True
            frmMain.mnuTools_ChatBot.Tag = "Other"
        End If
        
        Dim FileNum As Integer, Data As String
        FileNum = FreeFile
        Open File For Input As #FileNum
        Do Until EOF(FileNum)
            Line Input #FileNum, Data
            If Left$(Data, 1) <> "'" Then
                ReDim Preserve ChatBot(ArraySize(ChatBot) + 1)
                ChatBot(UBound(ChatBot)) = Data
            End If
        Loop
        Close #FileNum
        Call ChangeGMStatus("Bot loaded.")
        Dim FileName As String
        FileName = Right$(File, Len(File) - InStrRev(File, "\"))
        If InStr(FileName, ".") > 0 Then
            FileName = Left$(FileName, InStr(FileName, ".") - 1)
        End If
    Else
        Call ChangeGMStatus("Bot unloaded.")
    End If
End Sub

Public Function BotReply(Message As String) As String
    On Error Resume Next
    
    Dim i As Integer
    For i = 0 To ArraySize(ChatBot)
        If Message Like Left$(ChatBot(i), InStr(ChatBot(i), "=") - 1) Then
            BotReply = Right$(ChatBot(i), Len(ChatBot(i)) - InStr(ChatBot(i), "="))
            Dim Words() As String, j As Integer, Num As Integer
            Words = Split(BotReply)
            For j = 0 To ArraySize(Words)
                If Left$(Words(j), 1) = "$" And IsNumeric(Right$(Words(j), Len(Words(j)) - 1)) Then
                    Num = Val(Right$(Words(j), Len(Words(j)) - 1)) - 1
                    If Num <= UBound(Split(Message)) Then
                        BotReply = Replace$(BotReply, Words(j), Split(Message)(Num))
                    End If
                End If
            Next
            Exit For
        End If
    Next
End Function

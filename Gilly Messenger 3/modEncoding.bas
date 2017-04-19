Attribute VB_Name = "modEncoding"
Option Explicit

'UTF-8 Encoding
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Const CP_UTF8 = 65001

'Base64 Encoding
Private Const Base64Chars As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

Public Function UTF8_Encode(ByVal Text As String) As String
    On Error Resume Next
    
    Dim sBuffer As String
    Dim lLength As Long

    lLength = WideCharToMultiByte(CP_UTF8, 0, StrPtr(Text), -1, 0, 0, 0, 0)
    sBuffer = Space$(lLength)
    lLength = WideCharToMultiByte(CP_UTF8, 0, StrPtr(Text), -1, StrPtr(sBuffer), Len(sBuffer), 0, 0)
    sBuffer = StrConv(sBuffer, vbUnicode)
    UTF8_Encode = Left$(sBuffer, lLength - 1)
End Function

Public Function UTF8_Decode(ByVal Text As String) As String
    On Error Resume Next
    
    Dim lLength As Long
    Dim sBuffer As String

    Text = StrConv(Text, vbFromUnicode)
    lLength = MultiByteToWideChar(CP_UTF8, 0, StrPtr(Text), -1, 0, 0)
    sBuffer = Space$(lLength)
    lLength = MultiByteToWideChar(CP_UTF8, 0, StrPtr(Text), -1, StrPtr(sBuffer), Len(sBuffer))
    UTF8_Decode = Left$(sBuffer, lLength - 1)
End Function

Public Function URL_Encode(ByVal sText As String) As String
    On Error Resume Next
    
    Dim i As Long, Char As String
    For i = 1 To Len(sText)
        Char = Mid$(sText, i, 1)
        If Char Like "[a-zA-Z0-9$-_.+!*'(),;/?:@=&]" Then
            URL_Encode = URL_Encode & Char
        Else
            Char = Hex(Asc(Char))
            If Len(Char) < 2 Then
                Char = "0" & Char
            End If
            URL_Encode = URL_Encode & "%" & Char
        End If
    Next
End Function

Public Function URL_Decode(ByVal sText As String) As String
    On Error Resume Next
    
    Dim i As Long, Char As String
    For i = 1 To Len(sText)
        Char = Mid$(sText, i, 1)
        If Char = "%" Then
            URL_Decode = URL_Decode & Chr$(Val("&H" & Mid$(sText, i + 1, 2)))
            i = i + 2
        Else
            URL_Decode = URL_Decode & Char
        End If
    Next
End Function

Public Function Base64_Encode(ByVal NormalString As String, Optional ByVal Break As Integer = 0) As String
    Dim i As Long, Bin1Len8 As String, Bin2Len8 As String, Bin3Len8 As String
    Dim Bin1Len6 As String, Bin2Len6 As String, Bin3Len6 As String, Bin4Len6 As String
    
    If NormalString = vbNullString Then
        Exit Function
    End If
    
    For i = 1 To Len(NormalString) - 3 Step 3
        Bin1Len8 = Base10ToBinary(Asc(Mid$(NormalString, i, 1)))
        Bin2Len8 = Base10ToBinary(Asc(Mid$(NormalString, i + 1, 1)))
        Bin3Len8 = Base10ToBinary(Asc(Mid$(NormalString, i + 2, 1)))
        Call Bin3x8To4x6(Bin1Len8, Bin2Len8, Bin3Len8, Bin1Len6, Bin2Len6, Bin3Len6, Bin4Len6)
        Base64_Encode = Base64_Encode & Mid$(Base64Chars, BinaryToBase10(Bin1Len6) + 1, 1)
        Base64_Encode = Base64_Encode & Mid$(Base64Chars, BinaryToBase10(Bin2Len6) + 1, 1)
        Base64_Encode = Base64_Encode & Mid$(Base64Chars, BinaryToBase10(Bin3Len6) + 1, 1)
        Base64_Encode = Base64_Encode & Mid$(Base64Chars, BinaryToBase10(Bin4Len6) + 1, 1)
        DoEvents
    Next
    
    NormalString = Right$(NormalString, Len(NormalString) - IIf(Len(NormalString) / 3 = Int(Len(NormalString) / 3), Len(NormalString) - 3, Int(Len(NormalString) / 3) * 3))
    Bin1Len8 = Base10ToBinary(Asc(Left$(NormalString, 1)))

    If Len(NormalString) >= 2 Then
        Bin2Len8 = Base10ToBinary(Asc(Mid$(NormalString, 2, 1)))
    Else
        Bin2Len8 = "0"
    End If
    
    If Len(NormalString) = 3 Then
        Bin3Len8 = Base10ToBinary(Asc(Right$(NormalString, 1)))
    Else
        Bin3Len8 = "0"
    End If
    
    Call Bin3x8To4x6(Bin1Len8, Bin2Len8, Bin3Len8, Bin1Len6, Bin2Len6, Bin3Len6, Bin4Len6)
    
    Base64_Encode = Base64_Encode & Mid$(Base64Chars, BinaryToBase10(Bin1Len6) + 1, 1)
    Base64_Encode = Base64_Encode & Mid$(Base64Chars, BinaryToBase10(Bin2Len6) + 1, 1)
    Base64_Encode = Base64_Encode & IIf(Len(NormalString) >= 2, Mid(Base64Chars, BinaryToBase10(Bin3Len6) + 1, 1), "=")
    Base64_Encode = Base64_Encode & IIf(Len(NormalString) = 3, Mid(Base64Chars, BinaryToBase10(Bin4Len6) + 1, 1), "=")
    
    If Break > 0 Then
        i = Break + 1
        While i < Len(Base64_Encode)
            Base64_Encode = Left$(Base64_Encode, i - 1) & vbCrLf & Mid$(Base64_Encode, i)
            i = i + Break + 2
        Wend
    End If
End Function

Public Function Base64_Decode(ByVal Base64String As String) As String
    Dim i As Long, Bin1Len8 As String, Bin2Len8 As String, Bin3Len8 As String
    Dim Bin1Len6 As String, Bin2Len6 As String, Bin3Len6 As String, Bin4Len6 As String

    Base64String = Replace$(Base64String, vbCr, "")
    Base64String = Replace$(Base64String, vbLf, "")

    If Base64String = vbNullString Then
        Exit Function
    End If
    
    For i = 0 To 255
        If InStr(Base64String, Chr$(i)) > 0 And Not ((InStr(Base64Chars, Chr$(i)) > 0) Or (i = Asc("="))) Then
            Exit Function
        End If
    Next
    
    If Not Len(Base64String) / 4 = Len(Base64String) \ 4 Then
        Exit Function
    End If

    For i = 1 To Len(Base64String) Step 4
        Bin1Len6 = Base10ToBinary(InStr(Base64Chars, Mid$(Base64String, i, 1)) - 1)
        Bin2Len6 = Base10ToBinary(InStr(Base64Chars, Mid$(Base64String, i + 1, 1)) - 1)
        
        If Mid$(Base64String, i + 2, 1) = "=" Then
            Bin3Len6 = "0"
        Else
            Bin3Len6 = Base10ToBinary(InStr(Base64Chars, Mid$(Base64String, i + 2, 1)) - 1)
        End If
            
        If Mid$(Base64String, i + 3, 1) = "=" Then
            Bin4Len6 = "0"
        Else
            Bin4Len6 = Base10ToBinary(InStr(Base64Chars, Mid$(Base64String, i + 3, 1)) - 1)
        End If
            
        Call Bin4x6To3x8(Bin1Len6, Bin2Len6, Bin3Len6, Bin4Len6, Bin1Len8, Bin2Len8, Bin3Len8)

        Base64_Decode = Base64_Decode & Chr(BinaryToBase10(Bin1Len8))
        If Not Mid$(Base64String, i + 2, 1) = "=" Then
            Base64_Decode = Base64_Decode & Chr$(BinaryToBase10(Bin2Len8))
        End If
        If Not Mid$(Base64String, i + 3, 1) = "=" Then
            Base64_Decode = Base64_Decode & Chr$(BinaryToBase10(Bin3Len8))
        End If
        DoEvents
    Next i
End Function

Private Function Base10ToBinary(ByVal Base10 As Long) As String
    Dim PrevResult As Integer, CurResult As Integer
    If Base10 = 0 Then
        Base10ToBinary = "0"
        Exit Function
    End If
    Do
        CurResult = Int(Log(Base10) / Log(2))
        If PrevResult = 0 Then PrevResult = CurResult + 1
        Base10ToBinary = Base10ToBinary & String$(PrevResult - CurResult - 1, "0") & "1"
        Base10 = Base10 - 2 ^ CurResult
        PrevResult = CurResult
    Loop Until Base10 = 0
    Base10ToBinary = Base10ToBinary & String$(CurResult, "0")
End Function

Private Function BinaryToBase10(ByVal Binary As String) As Long
    Dim i As Integer
    For i = Len(Binary) To 1 Step -1
        BinaryToBase10 = BinaryToBase10 + Val(Mid$(Binary, i, 1)) * 2 ^ (Len(Binary) - i)
    Next
End Function

Private Sub Bin3x8To4x6(ByVal Bin1Len8 As String, ByVal Bin2Len8 As String, ByVal Bin3Len8 As String, ByRef Bin1Len6 As String, ByRef Bin2Len6 As String, ByRef Bin3Len6 As String, ByRef Bin4Len6 As String)
    Bin1Len8 = Right$("0000000" & Bin1Len8, 8)
    Bin2Len8 = Right$("0000000" & Bin2Len8, 8)
    Bin3Len8 = Right$("0000000" & Bin3Len8, 8)
    Bin1Len6 = Left$(Bin1Len8, 6)
    Bin2Len6 = Right$(Bin1Len8, 2) & Left$(Bin2Len8, 4)
    Bin3Len6 = Right$(Bin2Len8, 4) & Left$(Bin3Len8, 2)
    Bin4Len6 = Right$(Bin3Len8, 6)
End Sub

Private Sub Bin4x6To3x8(ByVal Bin1Len6 As String, ByVal Bin2Len6 As String, ByVal Bin3Len6 As String, ByVal Bin4Len6 As String, ByRef Bin1Len8 As String, ByRef Bin2Len8 As String, ByRef Bin3Len8 As String)
    Bin1Len6 = Right$("00000" & Bin1Len6, 6)
    Bin2Len6 = Right$("00000" & Bin2Len6, 6)
    Bin3Len6 = Right$("00000" & Bin3Len6, 6)
    Bin4Len6 = Right$("00000" & Bin4Len6, 6)
    Bin1Len8 = Bin1Len6 & Left$(Bin2Len6, 2)
    Bin2Len8 = Right$(Bin2Len6, 4) & Left$(Bin3Len6, 4)
    Bin3Len8 = Right$(Bin3Len6, 2) & Bin4Len6
End Sub

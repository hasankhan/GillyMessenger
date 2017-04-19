VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSSL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------
' Copyright (C) 2003 Jason K. Resch
'
' This program is free software; you can redistribute it and/or
' modify it under the terms of the GNU General Private License
' as published by the Free Software Foundation; either version 2
' of the License, or (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Private License for more details.
'
' You should have received a copy of the GNU General Private License
' along with this program; if not, write to the Free Software
' Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
'
' Author: Jason K. Resch
' URL: http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=43694&lngWId=1
'-----------------------------------------------

Option Explicit 'Declare All Variables

'CryptoAPI Functions
Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hSessionKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, ByVal pbData As String, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal hHash As Long, ByVal dwParam As Long, ByVal pbData As String, ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Private Declare Function CryptDeriveKey Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hBaseData As Long, ByVal dwFlags As Long, ByRef hSessionKey As Long) As Long
Private Declare Function CryptEncrypt Lib "advapi32.dll" (ByVal hSessionKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, ByRef pdwDataLen As Long, ByVal dwBufLen As Long) As Long
Private Declare Function CryptDecrypt Lib "advapi32.dll" (ByVal hSessionKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, ByRef pdwDataLen As Long) As Long
Private Declare Function CryptDestroyKey Lib "advapi32.dll" (ByVal hSessionKey As Long) As Long
Private Declare Function CryptImportKey Lib "advapi32.dll" (ByVal hProv As Long, ByVal pbData As String, ByVal dwDataLen As Long, ByVal hPubKey As Long, ByVal dwFlags As Long, ByRef phKey As Long) As Long
Private Declare Function CryptExportKey Lib "advapi32.dll" (ByVal hSessionKey As Long, ByVal hExpKey As Long, ByVal dwBlobType As Long, ByVal dwFlags As Long, ByVal pbData As String, ByRef pdwDataLen As Long) As Long
Private Declare Function CryptGenRandom Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwLen As Long, ByVal pbBuffer As String) As Long

'CryptoAPI Constant
Private Const SERVICE_PROVIDER As String = "Microsoft Enhanced Cryptographic Provider v1.0" & vbNullChar
Private Const KEY_CONTAINER As String = "GCN SSL Container" & vbNullChar
Private Const PROV_RSA_FULL As Long = 1
Private Const CRYPT_NEWKEYSET As Long = 8
Private Const CRYPT_EXPORTABLE As Long = 1
Private Const CALG_MD5 As Long = 32771
Private Const CALG_RC4 As Long = 26625
Private Const HP_HASHVAL As Long = 2
Private Const SIMPLEBLOB As Long = 1
Private Const GEN_KEY_BITS As Long = &H800000

'Class Variables
Dim hCryptProv As Long
Dim hClientWriteKey As Long
Dim hClientReadKey As Long
Dim hMasterKey As Long
Dim lngType As Long

'Variables for Parsing
Private Layer As Integer
Private SSL_Buffer As String
Private SeekLen As Integer

'Encryption Keys
Private MASTER_KEY As String
Private CLIENT_READ_KEY As String
Private CLIENT_WRITE_KEY As String

'Server Attributes
Private Private_KEY As String
Private ENCODED_CERT As String
Private CONNECTION_ID As String

'Counters
Private SEND_SEQUENCE_NUMBER As Double
Private RECV_SEQUENCE_NUMBER As Double

'Hand Shake Variables
Private CLIENT_HELLO As String
Private CHALLENGE_DATA As String

'SSL Socket
Private WithEvents wskSSL As Winsock
Attribute wskSSL.VB_VarHelpID = -1
Private SocketClosed As Boolean

Private Sub CertToPrivateKey()

    'Create CryptoAPI Blob from Certificate
    Const lPbkLen As Long = 1024
    Dim lOffset As Long
    Dim lStart As Long
    Dim sBlkLen As String
    Dim sRevKey As String
    Dim ASNStart As Long
    Dim ASNKEY As String

    lOffset = CLng(lPbkLen \ 8)
    lStart = 5 + (lOffset \ 128) * 2

    ASNStart = InStr(1, ENCODED_CERT, Chr(48) & Chr(129) & Chr(137) & Chr(2) & Chr(129) & Chr(129) & Chr(0)) + lStart
    ASNKEY = Mid(ENCODED_CERT, ASNStart, 128)

    sRevKey = StrReverse(ASNKEY)

    sBlkLen = CStr(Hex(lPbkLen \ 256))
    If Len(sBlkLen) = 1 Then sBlkLen = "0" & sBlkLen

    Private_KEY = (HexToBin( _
            "06020000" & _
            "00A40000" & _
            "52534131" & _
            "00" & sBlkLen & "0000" & _
            "01000100") & sRevKey)

End Sub

Private Function VerifyMAC(ByVal DecryptedRecord As String) As Boolean

    'Verify the Message Authentication Code
    Dim PrependedMAC As String
    Dim RecordData As String
    Dim CalculatedMAC As String
    
    PrependedMAC = Mid(DecryptedRecord, 1, 16)
    RecordData = Mid(DecryptedRecord, 17)
    
    CalculatedMAC = MD5_Hash(CLIENT_READ_KEY & RecordData & RecvSequence)
    
    Call IncrementRecv

    If CalculatedMAC = PrependedMAC Then
        VerifyMAC = True
    Else
        VerifyMAC = False
    End If

End Function

Private Function SendSequence() As String

    'Convert Send Counter to a String
    Dim TempString As String
    Dim TempSequence As Double
    Dim TempByte As Double
    Dim i As Integer
    
    TempSequence = SEND_SEQUENCE_NUMBER
    
    For i = 1 To 4
        TempByte = 256 * ((TempSequence / 256) - Int(TempSequence / 256))
        TempSequence = Int(TempSequence / 256)
        TempString = Chr(TempByte) & TempString
    Next
    
    SendSequence = TempString

End Function

Private Function RecvSequence() As String

    'Convert Receive Counter to a String
    Dim TempString As String
    Dim TempSequence As Double
    Dim TempByte As Double
    Dim i As Integer
    
    TempSequence = RECV_SEQUENCE_NUMBER
    
    For i = 1 To 4
        TempByte = 256 * ((TempSequence / 256) - Int(TempSequence / 256))
        TempSequence = Int(TempSequence / 256)
        TempString = Chr(TempByte) & TempString
    Next
    
    RecvSequence = TempString

End Function

Private Sub SendClientHello()
    On Error Resume Next
    
    'Send Client Hello
    Layer = 0
    
    Call GenerateRandomBytes(16, CHALLENGE_DATA)
    
    SEND_SEQUENCE_NUMBER = 0
    RECV_SEQUENCE_NUMBER = 0
    
    CLIENT_HELLO = Chr(1) & _
                    Chr(0) & Chr(2) & _
                    Chr(0) & Chr(3) & _
                    Chr(0) & Chr(0) & _
                    Chr(0) & Chr(Len(CHALLENGE_DATA)) & _
                    Chr(1) & Chr(0) & Chr(128) & _
                    CHALLENGE_DATA

    If wskSSL.State = 7 Then wskSSL.SendData AddRecordHeader(CLIENT_HELLO)

End Sub

Private Sub SendMasterKey()
    On Error Resume Next
    
    'Send Master Key
    Layer = 1
    
    Call GenerateRandomBytes(32, MASTER_KEY)

    Call CertToPrivateKey

    wskSSL.SendData AddRecordHeader(Chr(2) & _
                                    Chr(1) & Chr(0) & Chr(128) & _
                                    Chr(0) & Chr(0) & _
                                    Chr(0) & Chr(128) & _
                                    Chr(0) & Chr(0) & _
                                    ExportKeyBlob(MASTER_KEY, CLIENT_READ_KEY, CLIENT_WRITE_KEY, CHALLENGE_DATA, CONNECTION_ID, Private_KEY))

End Sub

Private Sub SendClientFinish()

    'Send ClientFinished Message
    Layer = 2
    Call SSLSend(Chr(3) & CONNECTION_ID)

End Sub

Private Sub SSLSend(ByVal Plaintext As String)
    On Error Resume Next
    
    'Send Plaintext as an Encrypted SSL Record
    Dim SSLRecord As String
    Dim OtherPart As String
    Dim SendAnother As Boolean
    
    If Len(Plaintext) > 32751 Then
        SendAnother = True
        Plaintext = Mid(Plaintext, 1, 32751)
        OtherPart = Mid(Plaintext, 32752)
    Else
        SendAnother = False
    End If
    
    SSLRecord = AddMACData(Plaintext)
    SSLRecord = RC4_Encrypt(SSLRecord)
    SSLRecord = AddRecordHeader(SSLRecord)
    
    wskSSL.SendData SSLRecord
    
    If SendAnother Then
        Call SSLSend(OtherPart)
    End If

End Sub

Private Function AddMACData(ByVal Plaintext As String) As String

    'Prepend MAC Data to the Plaintext
    AddMACData = MD5_Hash(CLIENT_WRITE_KEY & Plaintext & SendSequence) & Plaintext

End Function

Private Function AddRecordHeader(ByVal RecordData As String) As String

    'Prepend SLL Record Header to the Data Record
    Dim FirstChar As String
    Dim LastChar As String
    Dim TheLen As Long
        
    TheLen = Len(RecordData)
    
    FirstChar = Chr(128 + (TheLen \ 256))
    LastChar = Chr(TheLen Mod 256)

    AddRecordHeader = FirstChar & LastChar & RecordData
    
    Call IncrementSend

End Function

Private Sub IncrementSend()

    'Increment Counter for Each Record Sent
    SEND_SEQUENCE_NUMBER = SEND_SEQUENCE_NUMBER + 1
    If SEND_SEQUENCE_NUMBER = 4294967296# Then SEND_SEQUENCE_NUMBER = 0

End Sub

Private Sub IncrementRecv()

    'Increment Counter for Each Record Received
    RECV_SEQUENCE_NUMBER = RECV_SEQUENCE_NUMBER + 1
    If RECV_SEQUENCE_NUMBER = 4294967296# Then RECV_SEQUENCE_NUMBER = 0

End Sub

Private Function BytesToLen(ByVal TwoBytes As String) As Long
    On Error Resume Next
    
    'Convert Byte Pair to Packet Length
    Dim FirstByteVal As Long
    FirstByteVal = Asc(Left(TwoBytes, 1))
    If FirstByteVal >= 128 Then FirstByteVal = FirstByteVal - 128
    
    BytesToLen = 256 * FirstByteVal + Asc(Right(TwoBytes, 1))

End Function

Private Function HexToBin(ByVal HexString As String) As String

    'Convert a Hexadecimal String to characters
    Dim BinString As String, i As Integer
    For i = 1 To Len(HexString) Step 2
        BinString = BinString & Chr(Val("&H" & Mid(HexString, i, 2)))
    Next i
    HexToBin = BinString

End Function

Private Function ExportKeyBlob(ByRef StrMasterKey As String, ByRef StrReadKey As String, ByRef StrWriteKey As String, ByVal strChallenge As String, ByVal StrConnectionID As String, ByVal StrPrivateKey As String) As String

    'Create Keys and Return PKCS Block
    Dim lngReturnValue As Long
    Dim lngLength As Long
    Dim rgbBlob As String
    Dim hPrivateKey As Long
    
    Call CreateKey(hMasterKey, StrMasterKey)
    StrMasterKey = MD5_Hash(StrMasterKey)
    
    Call CreateKey(hClientReadKey, StrMasterKey & "0" & strChallenge & StrConnectionID)
    Call CreateKey(hClientWriteKey, StrMasterKey & "1" & strChallenge & StrConnectionID)
    
    StrReadKey = MD5_Hash(StrMasterKey & "0" & strChallenge & StrConnectionID)
    StrWriteKey = MD5_Hash(StrMasterKey & "1" & strChallenge & StrConnectionID)

    lngReturnValue = CryptImportKey(hCryptProv, StrPrivateKey, Len(StrPrivateKey), 0, 0, hPrivateKey)

    lngReturnValue = CryptExportKey(hMasterKey, hPrivateKey, SIMPLEBLOB, 0, vbNull, lngLength)
    rgbBlob = String(lngLength, 0)
    lngReturnValue = CryptExportKey(hMasterKey, hPrivateKey, SIMPLEBLOB, 0, rgbBlob, lngLength)
    
    If Not hPrivateKey = 0 Then CryptDestroyKey hPrivateKey
    If Not hMasterKey = 0 Then CryptDestroyKey hMasterKey

    ExportKeyBlob = StrReverse(Right(rgbBlob, 128))

End Function

Private Sub CreateKey(ByRef KeyName As Long, ByVal HashData As String)

    'Create a Session Key from a Hash
    Dim lngParams As Long
    Dim lngReturnValue As Long
    Dim lngHashLen As Long
    Dim hHash As Long
    
    lngReturnValue = CryptCreateHash(hCryptProv, CALG_MD5, 0, 0, hHash)
    If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "Could not create a Hash Object (CryptCreateHash API)"
    
    lngReturnValue = CryptHashData(hHash, HashData, Len(HashData), 0)
    If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "Could not calculate a Hash Value (CryptHashData API)"
    
    lngParams = GEN_KEY_BITS Or CRYPT_EXPORTABLE
    lngReturnValue = CryptDeriveKey(hCryptProv, CALG_RC4, hHash, lngParams, KeyName)
    If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "Could not create a session key (CryptDeriveKey API)"
    
    If Not hHash = 0 Then CryptDestroyHash hHash
    
End Sub

Private Function RC4_Encrypt(ByVal Plaintext As String) As String

    'Encrypt with Client Write Key
    Dim lngLength As Long
    Dim lngReturnValue As Long
    
    lngLength = Len(Plaintext)
    lngReturnValue = CryptEncrypt(hClientWriteKey, 0, False, 0, Plaintext, lngLength, lngLength)

    RC4_Encrypt = Plaintext

End Function

Private Function RC4_Decrypt(ByVal Ciphertext As String) As String

    'Decrypt with Client Read Key
    Dim lngLength As Long
    Dim lngReturnValue As Long
    
    lngLength = Len(Ciphertext)
    lngReturnValue = CryptDecrypt(hClientReadKey, 0, False, 0, Ciphertext, lngLength)

    RC4_Decrypt = Ciphertext

End Function

Public Property Let Socket(Data As Winsock)
    Set wskSSL = Data
End Property


Private Sub Class_Initialize()

    'Initiate Secure Session
    Dim lngReturnValue As Long
    Dim TheAnswer As Long
    
    lngReturnValue = CryptAcquireContext(hCryptProv, KEY_CONTAINER, SERVICE_PROVIDER, PROV_RSA_FULL, CRYPT_NEWKEYSET) 'try to make a new key container
    
    If lngReturnValue = 0 Then
        lngReturnValue = CryptAcquireContext(hCryptProv, KEY_CONTAINER, SERVICE_PROVIDER, PROV_RSA_FULL, 0) 'try to get a handle to a key container that already exists, and if it fails...
        If lngReturnValue = 0 Then
            TheAnswer = MsgBox("You do not have the required High Encryption Pack installed." & vbCrLf & "Would like to download this pack from Microsoft's website?", vbCritical Or vbYesNo)
        End If
    End If
    
    If TheAnswer = vbYes Then
        ShellExecute 0, "open", "http://www.microsoft.com/windows/ie/ie6/downloads/recommended/128bit/default.mspx", vbNullString, vbNullString, 1
    End If
End Sub


Private Sub Class_Terminate()

    'Free up Memory
    If Not hClientWriteKey = 0 Then CryptDestroyKey hClientWriteKey
    If Not hClientReadKey = 0 Then CryptDestroyKey hClientReadKey
    If Not hCryptProv = 0 Then CryptReleaseContext hCryptProv, 0

End Sub

Private Function GenerateRandomBytes(ByVal Length As Long, ByRef TheString As String) As Boolean

    'Generate Random Bytes
    Dim i As Integer

    Randomize
    TheString = ""
    For i = 1 To Length
        TheString = TheString & Chr(Int(Rnd * 256))
    Next
    
    GenerateRandomBytes = CryptGenRandom(hCryptProv, Length, TheString)

End Function

Private Function MD5_Hash(ByVal TheString As String) As String

    'Digest a String using MD5
    Dim lngReturnValue As Long
    Dim strHash As String
    Dim hHash As Long
    Dim lngHashLen As Long
    
    lngReturnValue = CryptCreateHash(hCryptProv, CALG_MD5, 0, 0, hHash)
    lngReturnValue = CryptHashData(hHash, TheString, Len(TheString), 0)
    lngReturnValue = CryptGetHashParam(hHash, HP_HASHVAL, vbNull, lngHashLen, 0)
    strHash = String(lngHashLen, vbNullChar)
    lngReturnValue = CryptGetHashParam(hHash, HP_HASHVAL, strHash, lngHashLen, 0)
    
    If Not hHash = 0 Then CryptDestroyHash hHash
    
    MD5_Hash = strHash

End Function

Private Sub wskSSL_Close()
    wskSSL.Close
    SocketClosed = True
End Sub

Private Sub wskSSL_Connect()
    SocketClosed = False
    Call SendClientHello
End Sub

Private Sub wskSSL_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
    
    Dim TheData As String
    Dim ReachLen As Long
    Do
        If SeekLen = 0 Then
            If bytesTotal >= 2 Then
                wskSSL.GetData TheData, vbString, 2
                SeekLen = BytesToLen(TheData)
                bytesTotal = bytesTotal - 2
            Else
                Exit Sub
            End If
        End If
        
        If bytesTotal >= SeekLen Then
            wskSSL.GetData TheData, vbString, SeekLen
            bytesTotal = bytesTotal - SeekLen
        Else
            Exit Sub
        End If
        
        Select Case Layer
            Case 0:
                ENCODED_CERT = Mid(TheData, 12, BytesToLen(Mid(TheData, 6, 2)))
                CONNECTION_ID = Right(TheData, BytesToLen(Mid(TheData, 10, 2)))
                Call IncrementRecv
                Call SendMasterKey
            Case 1:
                TheData = RC4_Decrypt(TheData)
                If Right(TheData, Len(CHALLENGE_DATA)) = CHALLENGE_DATA Then
                    If VerifyMAC(TheData) Then
                        Call SendClientFinish
                    End If
                Else
                    wskSSL.Close
                End If
             Case 2:
                TheData = RC4_Decrypt(TheData)
                If VerifyMAC(TheData) = False Then
                    wskSSL.Close
                End If
                Layer = 3
                
             Case 3:
                TheData = RC4_Decrypt(TheData)
                If VerifyMAC(TheData) Then
                    SSL_Buffer = SSL_Buffer & Mid(TheData, 17)
                End If
        End Select
    
        SeekLen = 0
        
    Loop Until bytesTotal = 0
End Sub

Private Sub wskSSL_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    SocketClosed = True
End Sub

Public Function SSL_Get(RemoteHost As String, RemotePort As Integer, Page As String, Optional Headers As String) As String
    wskSSL.Close
    SeekLen = 0
    SSL_Buffer = vbNullString
    Layer = 0
    SocketClosed = False
    
    wskSSL.Connect RemoteHost, RemotePort
    
    Do Until Layer = 3
        If SocketClosed Then
            Exit Function
        End If
        DoEvents
        Sleep 1
    Loop
    
    If Headers = vbNullString Then
        SSLSend "GET /" & Page & " HTTP/1.1" & vbCrLf & vbCrLf
    Else
        SSLSend "GET /" & Page & " HTTP/1.1" & vbCrLf & Headers & vbCrLf & vbCrLf
    End If
    
    Do Until Right$(SSL_Buffer, 4) = vbCrLf & vbCrLf
        If SocketClosed Then
            Exit Function
        End If
        DoEvents
        Sleep 1
    Loop
    
    SSL_Get = SSL_Buffer
End Function

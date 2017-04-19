Attribute VB_Name = "modSSL"
'-----------------------------------------------
' Copyright (C) 2003 Jason K. Resch
'
' This program is free software; you can redistribute it and/or
' modify it under the terms of the GNU General Public License
' as published by the Free Software Foundation; either version 2
' of the License, or (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program; if not, write to the Free Software
' Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
'
' Author: Jason K. Resch
' URL: http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=43694&lngWId=1
'-----------------------------------------------


'Encryption Object
Public SecureSession As clsCrypto

'Variables for Parsing
Public Layer As Integer
Public strBuffer As String
Public Processing As Boolean
Public SeekLen As Integer

'Encryption Keys
Public MASTER_KEY As String
Public CLIENT_READ_KEY As String
Public CLIENT_WRITE_KEY As String

'Server Attributes
Public PUBLIC_KEY As String
Public ENCODED_CERT As String
Public CONNECTION_ID As String

'Counters
Public SEND_SEQUENCE_NUMBER As Double
Public RECV_SEQUENCE_NUMBER As Double

'Hand Shake Variables
Public CLIENT_HELLO As String
Public CHALLENGE_DATA As String


Private Sub CertToPublicKey()

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

    sRevKey = ReverseString(ASNKEY)

    sBlkLen = CStr(Hex(lPbkLen \ 256))
    If Len(sBlkLen) = 1 Then sBlkLen = "0" & sBlkLen

    PUBLIC_KEY = (HexToBin( _
            "06020000" & _
            "00A40000" & _
            "52534131" & _
            "00" & sBlkLen & "0000" & _
            "01000100") & sRevKey)

End Sub

Public Function VerifyMAC(ByVal DecryptedRecord As String) As Boolean

    'Verify the Message Authentication Code
    Dim PrependedMAC As String
    Dim RecordData As String
    Dim CalculatedMAC As String
    
    PrependedMAC = Mid(DecryptedRecord, 1, 16)
    RecordData = Mid(DecryptedRecord, 17)
    
    CalculatedMAC = SecureSession.MD5_Hash(CLIENT_READ_KEY & RecordData & RecvSequence)
    
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
    
    TempSequence = RECV_SEQUENCE_NUMBER
    
    For i = 1 To 4
        TempByte = 256 * ((TempSequence / 256) - Int(TempSequence / 256))
        TempSequence = Int(TempSequence / 256)
        TempString = Chr(TempByte) & TempString
    Next
    
    RecvSequence = TempString

End Function

Public Sub SendClientHello(ByRef Socket As Winsock)

    'Send Client Hello
    Layer = 0
    
    Call SecureSession.GenerateRandomBytes(16, CHALLENGE_DATA)
    
    SEND_SEQUENCE_NUMBER = 0
    RECV_SEQUENCE_NUMBER = 0
    
    CLIENT_HELLO = Chr(1) & _
                    Chr(0) & Chr(2) & _
                    Chr(0) & Chr(3) & _
                    Chr(0) & Chr(0) & _
                    Chr(0) & Chr(Len(CHALLENGE_DATA)) & _
                    Chr(1) & Chr(0) & Chr(128) & _
                    CHALLENGE_DATA

    If Socket.State = 7 Then Socket.SendData AddRecordHeader(CLIENT_HELLO)

End Sub

Public Sub SendMasterKey(ByRef Socket As Winsock)

    'Send Master Key
    Layer = 1
    
    Call SecureSession.GenerateRandomBytes(32, MASTER_KEY)

    Call CertToPublicKey

    Socket.SendData AddRecordHeader(Chr(2) & _
                                    Chr(1) & Chr(0) & Chr(128) & _
                                    Chr(0) & Chr(0) & _
                                    Chr(0) & Chr(128) & _
                                    Chr(0) & Chr(0) & _
                                    SecureSession.ExportKeyBlob(MASTER_KEY, CLIENT_READ_KEY, CLIENT_WRITE_KEY, CHALLENGE_DATA, CONNECTION_ID, PUBLIC_KEY))

End Sub

Public Sub SendClientFinish(ByRef Socket As Winsock)

    'Send ClientFinished Message
    Layer = 2
    Call SSLSend(Socket, Chr(3) & CONNECTION_ID)

End Sub

Public Sub SSLSend(ByRef Socket As Winsock, ByVal Plaintext As String)
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
    SSLRecord = SecureSession.RC4_Encrypt(SSLRecord)
    SSLRecord = AddRecordHeader(SSLRecord)
    
    Socket.SendData SSLRecord
    
    If SendAnother = True Then
        Call SSLSend(Socket, OtherPart)
    End If

End Sub

Private Function AddMACData(ByVal Plaintext As String) As String

    'Prepend MAC Data to the Plaintext
    AddMACData = SecureSession.MD5_Hash(CLIENT_WRITE_KEY & Plaintext & SendSequence) & Plaintext

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

Public Sub IncrementSend()

    'Increment Counter for Each Record Sent
    SEND_SEQUENCE_NUMBER = SEND_SEQUENCE_NUMBER + 1
    If SEND_SEQUENCE_NUMBER = 4294967296# Then SEND_SEQUENCE_NUMBER = 0

End Sub

Public Sub IncrementRecv()

    'Increment Counter for Each Record Received
    RECV_SEQUENCE_NUMBER = RECV_SEQUENCE_NUMBER + 1
    If RECV_SEQUENCE_NUMBER = 4294967296# Then RECV_SEQUENCE_NUMBER = 0

End Sub

Public Function BytesToLen(ByVal TwoBytes As String) As Long

    'Convert Byte Pair to Packet Length
    Dim FirstByteVal As Long
    FirstByteVal = Asc(Left(TwoBytes, 1))
    If FirstByteVal >= 128 Then FirstByteVal = FirstByteVal - 128
    
    BytesToLen = 256 * FirstByteVal + Asc(Right(TwoBytes, 1))

End Function

Private Function HexToBin(ByVal HexString As String) As String

    'Convert a Hexadecimal String to characters
    Dim BinString As String
    For i = 1 To Len(HexString) Step 2
        BinString = BinString & Chr(Val("&H" & Mid(HexString, i, 2)))
    Next i
    HexToBin = BinString

End Function

Public Function ReverseString(ByVal TheString As String) As String

    'Reverse String
    Dim Reversed As String
    For i = Len(TheString) To 1 Step -1
        Reversed = Reversed & Mid(TheString, i, 1)
    Next i
    ReverseString = Reversed

End Function






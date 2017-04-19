VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCrypto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

'CryptoAPI Constants
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


Public Function ExportKeyBlob(ByRef StrMasterKey As String, ByRef StrReadKey As String, ByRef StrWriteKey As String, ByVal strChallenge As String, ByVal StrConnectionID As String, ByVal StrPublicKey As String) As String

    'Create Keys and Return PKCS Block
    Dim lngReturnValue As Long
    Dim lngLength As Long
    Dim rgbBlob As String
    Dim hPublicKey As Long
    
    Call CreateKey(hMasterKey, StrMasterKey)
    StrMasterKey = MD5_Hash(StrMasterKey)
    
    Call CreateKey(hClientReadKey, StrMasterKey & "0" & strChallenge & StrConnectionID)
    Call CreateKey(hClientWriteKey, StrMasterKey & "1" & strChallenge & StrConnectionID)
    
    StrReadKey = MD5_Hash(StrMasterKey & "0" & strChallenge & StrConnectionID)
    StrWriteKey = MD5_Hash(StrMasterKey & "1" & strChallenge & StrConnectionID)

    lngReturnValue = CryptImportKey(hCryptProv, StrPublicKey, Len(StrPublicKey), 0, 0, hPublicKey)

    lngReturnValue = CryptExportKey(hMasterKey, hPublicKey, SIMPLEBLOB, 0, vbNull, lngLength)
    rgbBlob = String(lngLength, 0)
    lngReturnValue = CryptExportKey(hMasterKey, hPublicKey, SIMPLEBLOB, 0, rgbBlob, lngLength)
    
    If hPublicKey <> 0 Then CryptDestroyKey hPublicKey
    If hMasterKey <> 0 Then CryptDestroyKey hMasterKey

    ExportKeyBlob = ReverseString(Right(rgbBlob, 128))

End Function

Public Sub CreateKey(ByRef KeyName As Long, ByVal HashData As String)

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
    
    If hHash <> 0 Then CryptDestroyHash hHash
    
End Sub

Function RC4_Encrypt(ByVal Plaintext As String) As String

    'Encrypt with Client Write Key
    Dim lngLength As Long
    Dim lngReturnValue As Long
    
    lngLength = Len(Plaintext)
    lngReturnValue = CryptEncrypt(hClientWriteKey, 0, False, 0, Plaintext, lngLength, lngLength)

    RC4_Encrypt = Plaintext

End Function

Function RC4_Decrypt(ByVal Ciphertext As String) As String

    'Decrypt with Client Read Key
    Dim lngLength As Long
    Dim lngReturnValue As Long
    
    lngLength = Len(Ciphertext)
    lngReturnValue = CryptDecrypt(hClientReadKey, 0, False, 0, Ciphertext, lngLength)

    RC4_Decrypt = Ciphertext

End Function


Private Sub Class_Initialize()

    'Initiate Secure Session
    Dim lngReturnValue As Long
    Dim TheAnswer As Long
    
    lngReturnValue = CryptAcquireContext(hCryptProv, KEY_CONTAINER, SERVICE_PROVIDER, PROV_RSA_FULL, CRYPT_NEWKEYSET) 'try to make a new key container
    
    If lngReturnValue = 0 Then
        lngReturnValue = CryptAcquireContext(hCryptProv, KEY_CONTAINER, SERVICE_PROVIDER, PROV_RSA_FULL, 0) 'try to get a handle to a key container that already exists, and if it fails...
        If lngReturnValue = 0 Then TheAnswer = MsgBox("GCN has detected that you do not have the required High Encryption Pack installed." & vbCrLf & "Would like to download this pack from Microsoft's website?", vbCritical Or vbYesNo)
    End If
    
    If TheAnswer = vbYes Then
        ShellExecute frmMain.hwnd, vbNullString, "http://www.microsoft.com/windows/ie/downloads/recommended/128bit/default.asp", vbNullString, vbNullString, 1
        frmMain.wskSSL_Close
    End If
    
    If TheAnswer = vbNo Then
        frmMain.wskSSL_Close
    End If

End Sub


Private Sub Class_Terminate()

    'Free up Memory
    If hClientWriteKey <> 0 Then CryptDestroyKey hClientWriteKey
    If hClientReadKey <> 0 Then CryptDestroyKey hClientReadKey
    If hCryptProv <> 0 Then CryptReleaseContext hCryptProv, 0

End Sub

Public Function GenerateRandomBytes(ByVal Length As Long, ByRef TheString As String) As Boolean

    'Generate Random Bytes
    Dim i As Integer

    Randomize
    TheString = ""
    For i = 1 To Length
        TheString = TheString & Chr(Int(Rnd * 256))
    Next
    
    GenerateRandomBytes = CryptGenRandom(hCryptProv, Length, TheString)

End Function

Public Function MD5_Hash(ByVal TheString As String) As String

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
    
    If hHash <> 0 Then CryptDestroyHash hHash
    
    MD5_Hash = strHash

End Function



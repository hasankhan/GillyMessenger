Attribute VB_Name = "modRegistry"
Option Explicit

'This program needs 3 buttons
Private Const REG_SZ = 1 ' Unicode nul terminated string
Private Const REG_BINARY = 3 ' Free form binary
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
Private Const HKEY_CURRENT_CONFIG = &H80000005
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Private Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String) As String
    Dim lResult As Long, lValueType As Long, strBuf As String, lDataBufSize As Long
    'retrieve nformation about the key
    lResult = RegQueryValueEx(hKey, strValueName, 0, lValueType, ByVal 0, lDataBufSize)
    If lResult = 0 Then
        If lValueType = REG_SZ Then
            'Create a buffer
            strBuf = String(lDataBufSize, Chr$(0))
            'retrieve the key's content
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, ByVal strBuf, lDataBufSize)
            If lResult = 0 Then
                'Remove the unnecessary chr$(0)'s
                RegQueryStringValue = Left$(strBuf, InStr(1, strBuf, Chr$(0)) - 1)
            End If
        ElseIf lValueType = REG_BINARY Then
            Dim strData As Integer
            'retrieve the key's value
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, strData, lDataBufSize)
            If lResult = 0 Then
                RegQueryStringValue = strData
            End If
        End If
    End If
End Function

Public Function ReadRegKey(strKey As String)
    Dim Ret, hKey As Long, strPath As String, strValue As String
    hKey = BaseKey(Left$(strKey, InStr(strKey, "\") - 1))
    strPath = KeyPath(strKey)
    strValue = Right$(strKey, Len(strKey) - InStrRev(strKey, "\"))
    'Open the key
    RegOpenKey hKey, strPath, Ret
    'Get the key's content
    ReadRegKey = RegQueryStringValue(Ret, strValue)
    'Close the key
    RegCloseKey Ret
End Function

Public Sub WriteRegKey(strKey As String, strData As String)
    Dim Ret, hKey As Long, strPath As String, strValue As String
    hKey = BaseKey(Left$(strKey, InStr(strKey, "\") - 1))
    strPath = KeyPath(strKey)
    strValue = Right$(strKey, Len(strKey) - InStrRev(strKey, "\"))
    'Create a new key
    RegCreateKey hKey, strPath, Ret
    'Save a string to the key
    RegSetValueEx Ret, strValue, 0, REG_SZ, ByVal strData, Len(strData)
    'close the key
    RegCloseKey Ret
End Sub

Public Sub DeleteRegKey(strKey As String)
    Dim Ret, hKey As Long, strPath As String, strValue As String
    hKey = BaseKey(Left$(strKey, InStr(strKey, "\") - 1))
    strPath = KeyPath(strKey)
    strValue = Right$(strKey, Len(strKey) - InStrRev(strKey, "\"))
    'Create a new key
    RegCreateKey hKey, strPath, Ret
    'Delete the key's value
    RegDeleteValue Ret, strValue
    'close the key
    RegCloseKey Ret
End Sub

Private Function BaseKey(strKey As String) As Long
    Select Case LCase(strKey)
    Case "hkey_classes_root"
        BaseKey = HKEY_CLASSES_ROOT
    Case "hkey_current_user"
        BaseKey = HKEY_CURRENT_USER
    Case "hkey_local_machine"
        BaseKey = HKEY_LOCAL_MACHINE
    Case "hkey_users"
        BaseKey = HKEY_USERS
    Case "hkey_current_config"
        BaseKey = HKEY_CURRENT_CONFIG
    End Select
End Function

Private Function KeyPath(strKey As String) As String
    Dim i As Integer
    i = InStr(strKey, "\")
    KeyPath = Mid$(strKey, i + 1, InStrRev(strKey, "\") - i)
End Function

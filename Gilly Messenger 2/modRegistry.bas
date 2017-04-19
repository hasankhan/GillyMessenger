Attribute VB_Name = "modRegistry"
'Read a registry key
Public Function ReadRegKey(RegKey As String) As String
    On Error Resume Next
    Dim RegObj
    Set RegObj = CreateObject("WScript.Shell")
    ReadRegKey = RegObj.RegRead(RegKey)
    If ReadRegKey = "" Then
        ReadRegKey = "Key does not exist."
    End If
    Set RegObj = Nothing
End Function
   
'Write a registry key
Public Function WriteRegKey(RegKey As String, Value As String)
    On Error Resume Next
    Dim RegObj
    Set RegObj = CreateObject("WScript.Shell")
    RegObj.RegWrite RegKey, Value
    Set RegObj = Nothing
End Function
   
'Delete a registry key
Public Function DeleteRegKey(RegKey As String)
    On Error Resume Next
    Dim RegObj
    Set RegObj = CreateObject("WScript.Shell")
    RegObj.RegDelete (RegKey)
    Set RegObj = Nothing
End Function

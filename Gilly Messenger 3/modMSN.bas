Attribute VB_Name = "modMSN"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function StatusCode(Status As Variant) As Variant
    Select Case Status
    Case "NLN"
        StatusCode = msnStatus_Online
    Case msnStatus_Online
        StatusCode = "NLN"
    Case "BSY"
        StatusCode = msnStatus_Busy
    Case msnStatus_Busy
        StatusCode = "BSY"
    Case "BRB"
        StatusCode = msnStatus_BeRightBack
    Case msnStatus_BeRightBack
        StatusCode = "BRB"
    Case "AWY"
        StatusCode = msnStatus_Away
    Case msnStatus_Away
        StatusCode = "AWY"
    Case "PHN"
        StatusCode = msnStatus_OneThePhone
    Case msnStatus_OneThePhone
        StatusCode = "PHN"
    Case "LUN"
        StatusCode = msnStatus_OutToLunch
    Case msnStatus_OutToLunch
        StatusCode = "LUN"
    Case "IDL"
        StatusCode = msnStatus_Idle
    Case msnStatus_Idle
        StatusCode = "IDL"
    Case msnStatus_Offline
        StatusCode = "HDN"
    Case "HDN"
        StatusCode = msnStatus_Offline
    End Select
End Function

Public Function ListCode(list As Variant) As Variant
    Select Case list
    Case msnList_Allow
        ListCode = "AL"
    Case "AL"
        ListCode = msnList_Allow
    Case msnList_Block
        ListCode = "BL"
    Case "BL"
        ListCode = msnList_Block
    Case msnList_Forward
        ListCode = "FL"
    Case "FL"
        ListCode = msnList_Forward
    Case msnList_Reverse
        ListCode = "RL"
    Case "RL"
        ListCode = msnList_Reverse
    End Select
End Function

Public Function StatusName(Status As Integer) As String
    StatusName = Choose(Status + 1, "Online", "Busy", "Be Right Back", "Away", "On The Phone", "Out To Lunch", "Idle", "Offline", "Unknown")
End Function

Public Function MSN_Encode(Data As String) As String
    MSN_Encode = URL_Encode(UTF8_Encode(Data))
End Function

Public Function MSN_Decode(Data As String) As String
    MSN_Decode = URL_Decode(UTF8_Decode(Data))
End Function

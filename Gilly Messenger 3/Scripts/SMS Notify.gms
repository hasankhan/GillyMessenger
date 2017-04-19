Set $CellNo=(INP)
If $CellNo=NULL
    End
End If
Set $Counter=1 
Set $Contact=$ListContact[$Counter]
Set $PrevStatus[$Contact]=Status $Contact
Set $Counter=$Counter+1 
If $Counter<=$ListContacts 
    Goto 6
End If
ChangeStatus 7
Event ContactStatusChanged
    If $PrevStatus[$Email]=7
        If $Status<>7
            Execute C:\Program Files\SMS Messenger\SmsMessenger.exe $CellNo GM $Email has signed in.
        End If
    End If
    Set $PrevStatus[$Email]=$Status
End Event
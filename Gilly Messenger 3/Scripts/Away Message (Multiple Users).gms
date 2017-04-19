'------------------------------------------------------------------------------------- 
'Created by: Maverick 
'Website: www.mavetech.tk 
'Forum: www.cracksoft.net.pk/forums 
'Comment: This script will allow you to send a message to multiple contacts while you are away. 
'         All what you have to do is input the no. of messages, email of the contacts and messages. 
'         When these contacts change status, their message will be sent.
'------------------------------------------------------------------------------------- 
Set $NumberOfMessages = (INP)
Set $i = 0
If $NumberOfMessages = $i
    GoTo 14
Else
    Set $Contact = (INP)
    Set $AwayMessage=(INP)
    Set $Group[$i,0] = $Contact
    Set $Group[$i,1] = $AwayMessage
    Set $Group[$i,2] = False
    Set $i = $i + 1
    GoTo 3
End If
Set $Text="Hey i am away right now but i left this message for you"
ChangeStatus 3
Event ContactStatusChanged
    Set $j = 0
    If $Group[$j,0] = $Email
        If $Group[$j,2] <> True
            If $Status <> 7           
                Msg $Group[$j,0] $Text
                Msg $Group[$j,0] $Group[$j,1]
                Set $Group[$j,2] = True
            Else
                Set $Group[$j,2] = False
            End If
        End If
    Else
        Set $j = $j + 1
        GoTo 17
    End If
End Event
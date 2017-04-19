'-------------------------------------------------------------------------------------
'Created by: Maverick
'Website: www.mavetech.tk
'Forum: www.cracksoft.net.pk/forums
'Comment: This script will allow you to sEnd a message to someone while you are away.
'         All what you have to do is input the email of the contact and the message.
'         When this contact changes status, the message will be sent.
'-------------------------------------------------------------------------------------
Set $Contact=(INP)
Set $Text="Hey i am away right now but i left this message for you"
Set $YourAwayMessage=(INP)
Set $SendOnce=False
ChangeStatus 3
Event ContactStatusChanged
    If $SendOnce = True
        End
    End If
    If $Status<>7
        If $Email = $Contact
            Msg $Email $Text
            Msg $Email $YourAwayMessage
            Set $SendOnce=True
        End If
    End If
End Event
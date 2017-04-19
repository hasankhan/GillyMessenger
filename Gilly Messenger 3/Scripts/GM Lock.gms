'-------------------------------------------------------------------------------------
'Created by: Dilated
'Comment: You input a password, and it'll keep prompting you for that password until
'         you get it right, otherwise you can't do anything else
'-------------------------------------------------------------------------------------
Set $Password=(INP)
Set $Lockdown=True
If $Lockdown=True
   Set $Unlock=(INP)
   If $Unlock=$Password
      Set $Lockdown=False
   End If
Else
   End
End If
GoTo 3
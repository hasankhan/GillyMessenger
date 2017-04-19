' Send message to anyone as soon as he comes online.
'
' Get the email of that person
VAR EmailOfTarget:(INP)
' Get the message for that person
VAR Message:(INP)
' Check to see if the buddy is online
if VAL:EmailOfTarget is online
	' IF block starts from here
	' Message the buddy
	msg val:EmailOfTarget val:Message
	' End the script
	end
' IF block ends here
end if
' Go back to 3rd line.
gto 3
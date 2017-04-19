' Script to send nick popups to all online contacts + those who have blocked you
'
' Save your nickname into a variable called nick
VAR Nick:(Nick)
' Save the first message into a variable called Message_1
VAR Message_1:(INP)
' Save the second message into a variable called Message_1
VAR Message_2:(INP)
' Save the third message into a variable called Message_1
VAR Message_3:(INP)
' Save the fourth message into a variable called Message_1
VAR Message_4:(INP)
' Change your status to appear offline
CST 7
' Change your nickname to the 4th message
REN VAL:Message_4
' Change your status to online to send the first popup
CST 0
' Change your status to appear offline
CST 7
' Change your nickname to the 3rd message
REN VAL:Message_3
' Change your status to online to send the second popup
CST 0
' Change your status to appear offline
CST 7
' Change your nickname to the 2nd message
REN VAL:Message_2
' Change your status to online to send the third popup
CST 0
' Change your status to appear offline
CST 7
' Change your nickname to the 1st message
REN VAL:Message_1
' Change your status to online to send the fourth popup
CST 0
' Change your nick back to the original one
REN VAL:Nick
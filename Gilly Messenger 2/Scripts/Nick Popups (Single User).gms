' Script to send popups to single contact
'
' Get the email of the person whom you want to send popups into variable EmailOfTarget
Var EmailOfTarget:(INP)
' Get your current nick into a variable Nick
VAR Nick:(Nick)
' Get the popup messages (For detail check Nick Popups.gms)
VAR Message_1:(INP)
VAR Message_2:(INP)
VAR Message_3:(INP)
VAR Message_4:(INP)
' Block the person
BLK VAL:EmailOfTarget
' Change your nick
REN VAL:Message_4
' Unblock the person to send the first popup
UBK VAL:EmailOfTarget
BLK VAL:EmailOfTarget
REN VAL:Message_3
UBK VAL:EmailOfTarget
BLK VAL:EmailOfTarget
REN VAL:Message_2
UBK VAL:EmailOfTarget
BLK VAL:EmailOfTarget
REN VAL:Message_1
UBK VAL:EmailOfTarget
' Change your nick back to the original one
REN VAL:Nick
' Annoy anyone by regularly blocking and unblocking him.
'
' Get email of the person who we want to annoy into a variable called EmailOfTarget
VAR EmailOfTarget:(INP)
' Block the the person by taking his email from the variable
BLK VAL:EmailOfTarget
' Unblock the the person by taking his email from the variable
UBK VAL:EmailOfTarget
' Pause the execution of script for 5 seconds otherwise MSN will detect a flood
SLP 10000
' Goto back to second statement
GTO 2
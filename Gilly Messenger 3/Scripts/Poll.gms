'-------------------------------------------------------------------------------------
'					PollScript
'Created by:	Millenium
'Topic:		http://www.cracksoft.net.pk/forums/viewtopic.php?t=290
'Comment:	Poll script is a Gilly Messenger script that allows you to quickly set
'		up polls and let your contacts vote. The messageboxes are pretty
'		self explanatory. This script also has a powerful logging feature.
'-------------------------------------------------------------------------------------

Set $Title=PollScript by Millenium - (EMAIL)
Msgbox Welcome to the poll setup wizard! If you feel any needs of stopping this wizard, just leave the inputboxes blank or hit cancel!(CRLF)Please open the windows with the people you'd like to invite to vote.(CRLF)Please send !stop to one of your contacts to stop voting and logging.
Msgbox A log path will now be asked. If you don't want to log the poll, just leave the box blank or click cancel.
Set $LogNr=$Null
Set $LogPath=(INP)
If $LogPath<>$Null
    Set $LogNr=fOpen($LogPath,o)
End If
Msgbox $Choices is the number of poll options. Please enter a number between 2 and 15.(CRLF)$Question is the question that will be asked to your contacts.(CRLF)(CRLf)Click OK to start the wizard.
Set $Choices=2
Set $Question=(INP)
If $Question=$Null
    End
End If
Set $Choices=(INP)
If $Choices=$null
    End
End If
If $Choices>15
    Set $Choices=15
End If
If $Choices<2
    Set $Choices=2
End If

Set $Counter=1
Set $PollText[$Counter]=(INP)
Set $PollValue[$Counter]=0
If $PollText[$Counter]=$Null
    End
End If
Set $PollValue[$Counter]=0
Set $Counter=$Counter+1
If $Counter>$Choices
    Goto 37
End If
Goto 26

Set $Counter=1
Set $Msg=$Null
Set $Msg=$Msg(CRLF)$Counter.$PollText[$Counter] - Type !vote$Counter to vote for this.
Set $Counter=$Counter+1
If $Counter>$Choices
    Goto 45
End If
Goto 39

Set $TotalMessage=$Title(CRLF)(CRLF)Question: $Question(CRLF)$Msg(CRLf)(CRLF)Type !results to show results.

If $LogNr<>$Null
    Set $Msg=--- Possible answers:
    Set $Counter=1
    Set $Msg=$Msg(CRLF)--- $Counter.$PollText[$Counter]
    Set $Counter=$Counter+1
    If $Counter>$Choices
        Goto 55
    End If
    Goto 49
    Set $data=--- $Title(CRLF)(CRLF)--- Question: $Question(CRLF)(CRLF)$Msg(CRLF)(CRLF)--- [(DATE) (TIME)] Poll logging started.(CRLF)(CRLF)
    fWrite $LogNr $data
End If

MsgAll $TotalMessage

Event MessageReceived

    If $TotalMessage=$Null
        Goto 43
    End If

    If $Choices=$NUll Then
        Goto 43
    End If

    Set $Counter=1
    If $Message=!vote$Counter
        If $Voted[$Email]<>yes
            Set $PollValue[$Counter]=$PollValue[$Counter]+1
            Set $Voted[$Email]=yes
            If $LogNr<>$Null
                Set $data=--- [(DATE) (TIME)] $Email voted for $PollText[$Counter](CRLF)(CRLF)
                fWrite $LogNr $data
            End If
            MsgEx Thank you for casting your vote!
        Else
            MsgEx You have already voted in this poll.
        End If
    End If
    Set $Counter=$Counter+1
    If $Counter>$Choices
        Goto 26
    End If
    Goto 8

    Set $ResultsMsg=$Question(CRLF)
    Set $Counter=1
    Set $ResultsMsg=$ResultsMsg(CRLF)$Counter.$PollText[$Counter]: $PollValue[$Counter] votes.
    Set $Counter=$Counter+1
    If $Counter>$Choices
        Goto 34
    End If
    Goto 28

    If $Message=!results
        MsgEx $ResultsMsg
    End If

    If $Message=!poll
        MsgEx $TotalMessage
    End If

    If $Message=!vote
        MsgEx You must type !vote followed directly by the number (no space)
    End If

    Set $Trash=$Null
End Event

Event MessageSent

    Set $ResultsMsg=$Question(CRLF)
    Set $Counter=1
    Set $ResultsMsg=$ResultsMsg(CRLF)$Counter.$PollText[$Counter]: $PollValue[$Counter] votes.
    Set $Counter=$Counter+1
    If $Counter>$Choices
        Goto 9
    End If
    Goto 3

    If $Message=!invite
        MsgEx $TotalMessage
    End If

    If $Message=!results
        MsgEx $ResultsMsg
    End If

    Set $ResultsBuffer=--- [(DATE) (TIME)] Poll terminated.(CRLF)(CRLF)--- Results:
    Set $Counter=1
    Set $ResultsBuffer=$ResultsBuffer(CRLF)--- $Counter.$PollText[$Counter]: $PollValue[$Counter] votes.
    Set $Counter=$Counter+1
    If $Counter>$Choices
        Goto 23
    End If
    Goto 17

    If $Message=!stop
        If $LogNr<>$Null
            fWrite $LogNr $ResultsBuffer
            Sleep 250
            fClose $LogNr
        End If
        MsgEx Poll stopped. Results:(CRLF)(CRLF)$ResultsMsg
        End
    End If
End Event
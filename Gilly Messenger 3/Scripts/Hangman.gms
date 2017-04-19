'-------------------------------------------------------------------------------------
'Created by: Dilated
'Hotmail: ttristann@hotmail.com
'Comment: This is hangman, the letter guessing game. You enable the script and anyone
'         who wants to play just types '!hangman'. Next, you're prompted for a word
'         and the game begins!
'-------------------------------------------------------------------------------------
Event MessageReceived
    If $Message = !giveup
        Set $GameEnabled[$Email] = FALSE
        MsgEx You Lose!(CRLF)(CRLF)The Correct word was $Word[$Email]
    End If
    If $GameEnabled[$Email] = TRUE
        Set $Msglen[$Email] = Len $Message
        Set $Output[$Email] = $Null
        Set $Info[$Email] = $Null
        Set $Wrong[$Email] = TRUE
        If $Msglen[$Email] > 1
            MsgEx Only enter one letter, or '!giveup'
        Else
            Set $Guessed[$Email] = $Guessed[$Email] $Message,
            Set $int = 1
            If $int <= $Length[$Email]
                Set $Var[$Email] = Mid $int 1 $Word[$Email]
                If $Var[$Email] =  
                    Set $Output[$Email] = $Output[$Email]   x
                    Set $Holder = Len $Output[$Email]
                    Set $Holder = $Holder - 1
                    Set $Output[$Email] = Left $Holder $Output[$Email]
                    Set $int = $int + 1
                    Goto 15
                End If
                If $Let[$int][$Email] = FALSE
                    If $Message = $Var[$Email]
                        Set $Let[$int][$Email] = TRUE
                        Set $Wrong[$Email] = FALSE
                        Goto 15
                    Else
                        Set $Output[$Email] = $Output[$Email] _
                    End If
                Else
                    Set $Output[$Email] = $Output[$Email] $Var[$Email]
                End If
                Set $int = $int + 1
                Goto 15
            End If
            If $Wrong[$Email] = TRUE
                Set $Guesses[$Email] = $Guesses[$Email] + 1
                Set $Info[$Email] = Sorry, there are no $Message's(CRLF)(CRLF)
             Else
                Set $Info[$Email] = $Null
            End If
            Set $Blah = 1
            If $Guesses[$Email] = 0
                Set $Hanger[$Email] = ____(CRLF)||   I(CRLF)||(CRLF)||(CRLF)||(CRLF)||_____    [
            End If
            If $Guesses[$Email] = 1
                Set $Hanger[$Email] = ____(CRLF)||   I(CRLF)||   O(CRLF)||(CRLF)||(CRLF)||_____    [
            End If
            If $Guesses[$Email] = 2
                Set $Hanger[$Email] = ____(CRLF)||   I(CRLF)||   O(CRLF)||   |(CRLF)||(CRLF)||_____    [
            End If
            If $Guesses[$Email] = 3
                Set $Hanger[$Email] = ____(CRLF)||   I(CRLF)||   O(CRLF)||  \|(CRLF)||(CRLF)||_____    [
            End If
            If $Guesses[$Email] = 4
                Set $Hanger[$Email] = ____(CRLF)||   I(CRLF)||   O(CRLF)||  \|/(CRLF)||(CRLF)||_____    [
            End If
            If $Guesses[$Email] = 5
                Set $Hanger[$Email] = ____(CRLF)||   I(CRLF)||   O(CRLF)||  -|-(CRLF)||  / (CRLF)||_____    [
            End If
            If $Guesses[$Email] = 6
                Set $Hanger[$Email] = ____(CRLF)||   I(CRLF)||   X(CRLF)||  /|\(CRLF)||  / \(CRLF)||_____    [
                Set $Info[$Email] = You lose! The correct word was $Word[$Email](CRLF)
    		Set $GameEnabled[$Email] = FALSE
            End If
            If $Blah <= $Length[$Email]
              Set $Var[$Email] = Mid $Blah 1 $Word[$Email]
                 If $Var[$Email] =  
                   Set $Blah = $Blah + 1
                   Goto 69
                End If
                If $Let[$Blah][$Email] = FALSE
                    Goto 87
                Else
                    If $Blah = $Length[$Email]
		        MsgEx $Info[$Email]$Hanger[$Email]$Output[$Email]](CRLF)You figured it out! Congragulations!
			Set $GameEnabled[$Email] = FALSE
			Goto 117
                    End If
                    Set $Blah = $Blah + 1
                    Goto 69
                End If
            End If
	    MsgEx $Info[$Email]$Hanger[$Email]$Output[$Email]](CRLF)You have had $Guesses[$Email]/6 incorrect guesses so far!(CRLF)$Guessed[$Email]
        End If
    End If
    If $Message = !hangman
        If $GameEnabled[$Email] = TRUE
        Else
            Set $Word[$Email] = (INP)
            Set $Guesses[$Email] = 0
            Set $Guessed[$Email] = Guessed:
            Set $Length[$Email] = Len $Word[$Email]
            Set $Output[$Email] = ____(CRLF)||   I(CRLF)||(CRLF)||(CRLF)||(CRLF)||_____    [
            Set $int = 1
            If $int <= $Length[$Email]
                Set $Var[$Email] = Mid $int 1 $Word[$Email]
                If $Var[$Email] =  
                    Set $Output[$Email] = $Output[$Email]   x
                    Set $Holder = Len $Output[$Email]
                    Set $Holder = $Holder - 1
                    Set $Output[$Email] = Left $Holder $Output[$Email]
                Else
                    Set $Output[$Email] = $Output[$Email] _
                End If
                Set $Let[$int][$Email] = FALSE
                Set $int = $int + 1
                Goto 99
            End If
            MsgEx $Output[$Email]]
        End If
        Set $GameEnabled[$Email] = TRUE
    End If
End Event
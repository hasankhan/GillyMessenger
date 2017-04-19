Set $MyWinningCombs[1] = XXX??????
Set $MyWinningCombs[2] = ???XXX???
Set $MyWinningCombs[3] = ??????XXX
Set $MyWinningCombs[4] = X??X??X??
Set $MyWinningCombs[5] = ?X??X??X?
Set $MyWinningCombs[6] = ??X??X??X
Set $MyWinningCombs[7] = X???X???X
Set $MyWinningCombs[8] = ??X?X?X??
Set $BuddyWinningCombs[1] = OOO??????
Set $BuddyWinningCombs[2] = ???OOO???
Set $BuddyWinningCombs[3] = ??????OOO
Set $BuddyWinningCombs[4] = O??O??O??
Set $BuddyWinningCombs[5] = ?O??O??O?
Set $BuddyWinningCombs[6] = ??O??O??O
Set $BuddyWinningCombs[7] = O???O???O
Set $BuddyWinningCombs[8] = ??O?O?O??

Event MessageSent
    If $GameEnabled[$Email] <> True
        If $Message = !tic tac toe
            Set $GameMap[$Email] = 123456789
            Set $GameEnabled[$Email] = True
            Set $Toss = Rand 2
            If $Toss = 2
                Set $GameTurn[$Email] = (Email)
            Else
                Set $GameTurn[$Email] = $Email
            End If
            MsgEx 1|2|3(CrLf)----------(CrLf)4|5|6(CrLf)----------(CrLf)7|8|9(CrLf)(CrLf)$GameTurn[$Email] will go first.
        End If
    Else
        If $Message like [1-9]
            If $GameTurn[$Email] = (Email)
                Set $GameChar = Mid $Message 1 $GameMap[$Email]
                If $GameChar = X
                    MsgEx You have already marked this position.
                Else
                    If $GameChar = O
                        MsgEx $Email has already marked this position.
                    Else
                        Set $Temp = $Message - 1
                        Set $LeftMap = Left $Temp $GameMap[$Email]
                        Set $Temp = $Message + 1
                        Set $RightMap = Mid $Temp $GameMap[$Email]
                        Set $GameMap[$Email] = $LeftMapX$RightMap
                        Set $Counter = 1
                        Set $GameDraw = True
                        If $Counter <= 9
                            Set $GameChar = Mid $Counter 1 $GameMap[$Email]
                            If $GameChar like #
                                Set $GameDraw = False
                            End If
                            Set $Counter = $Counter + 1
                            Goto 30
                        End If
                        Set $Counter = 1
                        If $Counter <= 8
			    If $GameMap[$Email] like $MyWinningCombs[$Counter]
                                Set $GameEnabled[$Email] = False
                            End If
                            Set $Counter = $Counter + 1
                            Goto 39
                        End If
                        Set $Char11 = Mid 1 1 $GameMap[$Email]
                        Set $Char12 = Mid 2 1 $GameMap[$Email]
                        Set $Char13 = Mid 3 1 $GameMap[$Email]
                        Set $Char21 = Mid 4 1 $GameMap[$Email]
                        Set $Char22 = Mid 5 1 $GameMap[$Email]
                        Set $Char23 = Mid 6 1 $GameMap[$Email]
                        Set $Char31 = Mid 7 1 $GameMap[$Email]
                        Set $Char32 = Mid 8 1 $GameMap[$Email]
                        Set $Char33 = Mid 9 1 $GameMap[$Email]
                        If $GameEnabled[$Email] = False
                            Set $Info = (CrLf)(CrLf)(Email) has won tic tac toe.
                        Else
                            If $GameDraw = True
                                Set $Info = (CrLf)(CrLf)Game is draw.
                                Set $GameEnabled[$Email] = False
                            Else
                                Set $GameTurn[$Email] = $Email
				Set $Info = $Null
                            End If
                        End If
			MsgEx $Char11|$Char12|$Char13(CrLf)----------(CrLf)$Char21|$Char22|$Char23(CrLf)----------(CrLf)$Char31|$Char32|$Char33$Info
                    End If
                End If
            Else
                MsgEx Its $Email's turn.
            End If    
        Else
            If $Message = 0
                Set $Char11 = Mid 1 1 $GameMap[$Email]
                Set $Char12 = Mid 2 1 $GameMap[$Email]
                Set $Char13 = Mid 3 1 $GameMap[$Email]
                Set $Char21 = Mid 4 1 $GameMap[$Email]
                Set $Char22 = Mid 5 1 $GameMap[$Email]
                Set $Char23 = Mid 6 1 $GameMap[$Email]
                Set $Char31 = Mid 7 1 $GameMap[$Email]
                Set $Char32 = Mid 8 1 $GameMap[$Email]
                Set $Char33 = Mid 9 1 $GameMap[$Email]
                MsgEx $Char11|$Char12|$Char13(CrLf)----------(CrLf)$Char21|$Char22|$Char23(CrLf)----------(CrLf)$Char31|$Char32|$Char33
            Else
                If $Message = !tic tac toe help
                    MsgEx Enter a number to place your sign on its place.(Crlf)You can enter 0 to view the game status.
                End If
            End If
        End If
    End If
End Event

Event MessageReceived
    If $GameEnabled[$Email] <> True
        If $Message = !tic tac toe
            Set $GameMap[$Email] = 123456789
            Set $GameEnabled[$Email] = True
            Set $Toss = Rand 2
            If $Toss = 2
                Set $GameTurn[$Email] = $Email
            Else
                Set $GameTurn[$Email] = (Email)
            End If
            MsgEx 1|2|3(CrLf)----------(CrLf)4|5|6(CrLf)----------(CrLf)7|8|9(CrLf)(CrLf)$GameTurn[$Email] will go first.
        End If
    Else
        If $Message like [1-9]
            If $GameTurn[$Email] = $Email
                Set $GameChar = Mid $Message 1 $GameMap[$Email]
                If $GameChar = O
                    MsgEx You have already marked this position.
                Else
                    If $GameChar = X
                        MsgEx $Email has already marked this position.
                    Else
                        Set $Temp = $Message - 1
                        Set $LeftMap = Left $Temp $GameMap[$Email]
                        Set $Temp = $Message + 1
                        Set $RightMap = Mid $Temp $GameMap[$Email]
                        Set $GameMap[$Email] = $LeftMapO$RightMap
                        Set $Counter = 1
                        Set $GameDraw = True
                        If $Counter <= 9
                            Set $GameChar = Mid $Counter 1 $GameMap[$Email]
                            If $GameChar like #
                                Set $GameDraw = False
                            End If
                            Set $Counter = $Counter + 1
                            Goto 30
                        End If
                        Set $Counter = 1
                        If $Counter <= 8
                            If $GameMap[$Email] like $BuddyWinningCombs[$Counter]
                                Set $GameEnabled[$Email] = False
                            End If
                            Set $Counter = $Counter + 1
                            Goto 39
                        End If
                        Set $Char11 = Mid 1 1 $GameMap[$Email]
                        Set $Char12 = Mid 2 1 $GameMap[$Email]
                        Set $Char13 = Mid 3 1 $GameMap[$Email]
                        Set $Char21 = Mid 4 1 $GameMap[$Email]
                        Set $Char22 = Mid 5 1 $GameMap[$Email]
                        Set $Char23 = Mid 6 1 $GameMap[$Email]
                        Set $Char31 = Mid 7 1 $GameMap[$Email]
                        Set $Char32 = Mid 8 1 $GameMap[$Email]
                        Set $Char33 = Mid 9 1 $GameMap[$Email]
			If $GameEnabled[$Email] = False
                            Set $Info = (CrLf)(CrLf)$Email has won tic tac toe.
                        Else
                            If $GameDraw = True
                                Set $Info = (CrLf)(CrLf)Game is draw.
                                Set $GameEnabled[$Email] = False
                            Else
                                Set $GameTurn[$Email] = (Email)
				Set $Info = $Null
                            End If
                        End If
			MsgEx $Char11|$Char12|$Char13(CrLf)----------(CrLf)$Char21|$Char22|$Char23(CrLf)----------(CrLf)$Char31|$Char32|$Char33$Info
                    End If
                End If
            Else
                MsgEx Its (Email)'s turn.
            End If
        Else
            If $Message = 0
                Set $Char11 = Mid 1 1 $GameMap[$Email]
                Set $Char12 = Mid 2 1 $GameMap[$Email]
                Set $Char13 = Mid 3 1 $GameMap[$Email]
                Set $Char21 = Mid 4 1 $GameMap[$Email]
                Set $Char22 = Mid 5 1 $GameMap[$Email]
                Set $Char23 = Mid 6 1 $GameMap[$Email]
                Set $Char31 = Mid 7 1 $GameMap[$Email]
                Set $Char32 = Mid 8 1 $GameMap[$Email]
                Set $Char33 = Mid 9 1 $GameMap[$Email]
                MsgEx $Char11|$Char12|$Char13(CrLf)----------(CrLf)$Char21|$Char22|$Char23(CrLf)----------(CrLf)$Char31|$Char32|$Char33
            Else
                If $Message = !tic tac toe help
                    MsgEx Enter a number to place your sign on its place.(Crlf)You can enter 0 to view the game status.
                End If
            End If
        End If
    End If
End Event
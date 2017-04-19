'------------------------------------------------------------------------------------- 
'Created by: Maverick 
'Website: www.mavetech.tk 
'Forum: www.cracksoft.net.pk/forums 
'Comment: This script will allow you to enter multiple sound file paths.. These sound files will be played in 
'   a sequential order everytime a person changes his/her status to anything :) 
'------------------------------------------------------------------------------------- 
Set $NumberOfSoundTracks=(INP)
Set $i=1
If $i<=$NumberOfSoundTracks
   Set $FileDirs[$i]=(INP)
   Set $i=$i+1
   GoTo 3
End If
Set $i=1
Event ContactStatusChanged
   If $i>=$NumberOfTracks
      Set $i=1
   End If
   PlaySound $FileDirs[$i]
   Set $i=$i+1
End Event
VAR EmailOfTarget:(INP)
VAR PathOfSound:(INP)
if VAL:EmailOfTarget is online
	SND VAL:PathOfSound
	end
end if
gto 3
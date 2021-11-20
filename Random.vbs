Function RndNum(Seed, MinValue, MaxValue)
    Num = (((Hour(Now) + 1) * (Minute(Now) + 1) * ((30 - Second(Now)) + 1) * Right((Seed * Right(Timer, 5)), 5)) Mod MaxValue) 
	RndNum = Abs(Num * Seed) Mod MaxValue
	if RndNum < MinValue then
	RndNum = RndNum + MinValue
	end if
End Function

Do
    counter = counter + 1
    Numbers = Numbers & RndNum(counter,50,800) & ", "
Loop While counter < 50

MsgBox Numbers
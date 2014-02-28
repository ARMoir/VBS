Info = Array("Test","Test","Test","Test","New","New")
x = 0
z = ubound(Info)
Do
x = x+1
Do
z = z-1
If x=z Then
Info(x) = Info(z)
ElseIf Info(x) = Info(z) Then
Info(x) = ""
End If
Loop Until z=0
z = ubound(Info)
Loop Until x=ubound(Info)
for each x in Info 
If x <> "" Then
Unique = Unique &" "& x
End If
next
Wscript.Echo Unique

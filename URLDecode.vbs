Function URLDecode(x)
 Dim a,b,c,n,s
 b=1
 s=Space(1)
 n=CInt(Len(x))
 Do Until (b -1)=n
	c=Mid(x,b,1)
	if c="+" Or c="%2B" Or c="%2b" Then
		a = a & s
	ElseIf	 c="%" Then
		a = a & Chr(CLng("&h" & Mid(x,b+1,2)))
		b = b+2
	Else
		a = a & c
	End if
	b=b+1
 Loop
 URLDecode=a
End Function

Dim test,test_enc
test="$ & + , / : ; = ? @"
test_enc="%24+%26+%2B+%2C+%2F+%3A+%3B+%3D+%3F+%40"
Response.Write test&"<br />"& URLDecode(test_enc)

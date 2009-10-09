<html>
	<head>
	</head>
	<body>
		1 + "2" = <%=1 + "2"%><BR/>
		"1" + "2" = <%="1" + "2"%><BR/>	
		1 + empty = <%=1 + empty%><BR/>
		1 + null = <%=1 + null%><BR/>
		Now + 1 = <%=Now + 1%><BR/>
		1 + Now = <%=1 + Now%><BR/>
		Now + Now = <%=Now + Now%><BR/>
		True + True = <%=True + True%><BR/>
		<BR/>
		"2" - 1 = <%="2" - 1%><BR/>
		"2" - "1" = <%="2" - "1"%><BR/>
		Now - 1 = <%=Now - 1%><BR/>
		1 - Now = <%=1 - Now%><BR/>
		Now - (Now-7) = <%=Now - (Now-7)%><BR/>
		True - True = <%=True - True%><BR/>
		<BR/>
		"2"/"1" = <%="2"/"1"%><BR/>
		True/True = <%=True/True%><BR/>
		Date/Date = <%=Date/Date%><BR/>
		<BR/>
		"True" and "True"=<%="True" and "True"%><BR/>
		"true" and "True"=<%="true" and "True"%><BR/>
		True and "True"=<%=True and "True"%><BR/>
		True and "1"=<%=True and "1"%><BR/>
		True and "-1"=<%=True and "-1"%><BR/>
		"Yes" and "Yes"=Type mismatch<BR/>
		1 and "True" = <%=1 and "True"%><BR/>
		1.1 and "True" = <%=1.1 and "True"%><BR/>
		<BR/>
		2 <= "1" = <%=2 <= "1"%><BR/>
		"1" <= "2" = <%="1" <= "2"%><BR/>
		Empty <= 1 = <%=Empty < 1%><BR/>
		Empty <= "1" = <%=Empty < "1"%><BR/>
		Empty = Empty = <%=Empty = Empty%><BR/>
	</body>
</html>
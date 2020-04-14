<html>
	<head>
	</head>
	<body>
		Vartype of 123456 is <%=VarType(123456) %><BR/>	
		Vartype of .23456 is <%=VarType(.23456) %><BR/>	
		VarType of <%=Server.HtmlEncode("&hFFFFFF")%> is <%=VarType(&hFFFFFF) %><BR/>	
		VarType of 1.23e-3 is <%=VarType(1.23e-3)%><BR/>
		VarType of #12/21/1999# is <%=VarType(#12/21/1999#)%><BR/>
	</body>
</html>
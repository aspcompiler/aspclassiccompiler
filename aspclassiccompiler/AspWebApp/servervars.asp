<html>
	<head>
	</head>
	<body>
	<%
	url = Request.ServerVariables("HTTP_URL")
x = instr(1, url, "url=", 1)
if x > 0 then
	url = Mid(url, x + 4)
end if
	%>
		<%=url%>
	</body>
</html>
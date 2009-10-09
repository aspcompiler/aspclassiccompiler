<%
option explicit
%>
<html>
	<head>
	</head>
	<body>
<%
sub printnum(n)
  response.write n
	for i = 1 to n
		response.write i
	next
end sub

dim i

printnum 5



%>
	</body>
</html>
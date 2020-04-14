<%
option explicit
%>
<html>
	<head>
	</head>
	<body>
<%
	dim i, color
	
	i = 3
	if i = 1 then
		color = "red"
	elseif i = 2 then
		color = "green"
	elseif i = 3 then
		color = "blue"
	else
		color = "blank"
	end if
	response.write color
%>
	</body>
</html>
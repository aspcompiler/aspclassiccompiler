<html>
	<head>
	</head>
	<body>
<%
Dim MyIndex
Dim MyArray (3)
MyArray(0) = "Sunday"
MyArray(1) = "Monday"
MyArray(2) = "Tuesday"
MyIndex = Filter(MyArray, "day") ' MyIndex(0) contains "Monday".

for each s in MyIndex
	response.write s & "<BR/>"
next
%>
	</body>
</html>
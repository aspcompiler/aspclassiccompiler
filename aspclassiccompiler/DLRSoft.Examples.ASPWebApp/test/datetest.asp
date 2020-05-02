<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head>
    <title></title>
</head>
<body>
<%= WeekdayName(Weekday(DateSerial(1963, 11, 22))) %>
<%= WeekdayName(Weekday(DateSerial(1963, 10, 15))) %>
<%= WeekdayName(Weekday(DateSerial(1995, 6, 2))) %>
<%= WeekdayName(Weekday(DateSerial(1997, 5, 5))) %>
<%
setlocale(4100)
for i = 1 to 7
 Response.Write(weekdayname(i) & "<BR/>")
next
 %>
 <%
 for i = 1 to 12
 Response.Write(monthname(i) & "<BR/>")
next
 %>
 <%=FormatDateTime(DateSerial(1997, 5, 5), 1) %>
<%=FormatCurrency(10000000) %>
</body>
</html>

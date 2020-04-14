<%
    imports system
    dim s = new system.text.stringbuilder()
    dim i
    
    s = s + "<table>"
    for i = 1 to 12
        s = s + "<tr>"
        s = s + "<td>" + i + "</td>"
        s = s + "<td>" + MonthName(i) + "</td>"
        s = s + "</tr>"
    next
    s = s + "</table>"
    response.Write(s)

%>
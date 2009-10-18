<%
imports system

dim sb

sb = new System.Text.StringBuilder()
sb.append("this")
sb.append(" is ")
sb.append(" stringbuilder!")
response.write sb.toString()

%>
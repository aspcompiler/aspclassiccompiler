<%
    sub mysub(byval a)
        a = a + 1
        response.Write a
    end sub

    mysub 5
%>
<%
    sub main()
        dim a
        a = 2
        mysub a
        response.Write a
    end sub

    sub mysub(a)
        a = a + 1
        response.Write a
    end sub
   
    main
%>
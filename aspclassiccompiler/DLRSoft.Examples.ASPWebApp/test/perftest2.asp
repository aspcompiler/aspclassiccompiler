<%
option explicit

perftest	

sub perftest()
  dim loops
  dim counter
  dim tickes
  tickes = Timer()

  loops = 10000000
  counter = 0.0

  while loops > 0  
	  loops = loops - 1  
	  counter = counter + 1.0
  wend

  Response.Write("DoLoop result " & counter & "<BR/>")
  Response.Write("Executing the while loop took: " & ((Timer()-tickes) * 1000.0) & "ms" & "<BR/>")
end sub

%>
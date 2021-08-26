<%
  Response.Cookies("user")("firstname")="SorR"
  Response.Cookies("user")("lastname")="Ki"
  Response.Cookies("user")("country")="Nowhere"
  Response.Cookies("user")("age")="25"
%>
<!DOCTYPE html>
<html>
<body>
  <%
    dim x, y
    for each x in request.cookies
      response.write "<p>" 
      if request.cookies( x ).HasKeys then
        for each y in request.cookies( x )
          response.write( x & " : " & y & " = " & request.cookies( x )( y ) )
          response.write( "<br/>" )
        next
      else
        response.write( x & " = " & request.cookies( x ) & "<br/>" )
      end if
      response.write "</p>"
    next
  %>
  <br/>
  <h2>request.cookies</h2>
  <%=request.cookies%>
</body>
</html>
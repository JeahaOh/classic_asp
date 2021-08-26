<% Option Explicit %>
<!DOCTYPE html>
<html>
<head>
  <%
    sub vbproc( x, y )
      response.write( "x * y = " & x * y )
    end sub
  %>
</head>
<body>
  <p>Result : <%call vbproc( 3, 4 )%></p>
</body>
</html>
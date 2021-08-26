<% @CODEPAGE="65001" language="vbscript" %>
<!DOCTYPE html>
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
</head>
<html>
<body>
  <%
    dim h
    h = hour( now() )

    response.write( "<p>" & now() & "</p>" )

    if h < 12 then
      response.write( "Good Morning!" )
    else
      response.write( "Good Day!" )
    end if
  %>
</body>
</html>
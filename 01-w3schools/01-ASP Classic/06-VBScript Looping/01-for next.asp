<% @CODEPAGE="65001" language="vbscript" %>
<% Response.CharSet = "utf-8" %>
<!DOCTYPE html>
<html>
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
</head>
<body>
  <%
    for i = 0 to 5
      response.write( "i : <strong>" & i & "</strong><br/>" )
    next
  %>
</body>
</html>
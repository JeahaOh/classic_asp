<% @CODEPAGE="65001" language="vbscript" %>
<% Response.CharSet = "utf-8" %>
<!DOCTYPE html>
<html>
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
</head>
<body>
  <h3>exit do while loop</h3>
  <pre>
    <code>
      i = 20
      do until i = 10
        i = i - 1
        if i < 10 then exit do
        response.write( "i : <strong>" & i & "</strong><br/>" )
      loop
    </code>
  </pre>
  <%
    i = 20
    do until i = 10
      i = i - 1
      if i < 10 then exit do
      response.write( "i : <strong>" & i & "</strong><br/>" )
    loop
  %>
</body>
</html>
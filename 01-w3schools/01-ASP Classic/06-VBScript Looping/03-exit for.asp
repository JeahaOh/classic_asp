<% @CODEPAGE="65001" language="vbscript" %>
<% Response.CharSet = "utf-8" %>
<!DOCTYPE html>
<html>
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
</head>
<body>
  <h3>exit for ... next</h3>
  <pre>
    <code>
      for i = 0 to 10
        if i = 5 then exit for
        response.write( "i : " & i )
      next
    </code>
  </pre>
  <%
    for i = 0 to 10
      if i = 5 then exit for
      response.write( "i : <strong>" & i & "</strong><br/>" )
    next
  %>
</body>
</html>
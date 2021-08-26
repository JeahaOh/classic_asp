<% @CODEPAGE="65001" language="vbscript" %>
<% Response.CharSet = "utf-8" %>
<!DOCTYPE html>
<html>
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
</head>
<body>
  <h3>do while loop</h3>
  <pre>
    <code>
    do while i > 10
      response.write( "i : " & i )
      i = i + 1
    loop
    </code>
  </pre>
  <%
    do while i < 10
      response.write( "i : <strong>" & i & "</strong><br/>" )
      i = i + 1
    loop
  %>

  <hr/>

  <h3>do loop while</h3>
  <pre>
    <code>
    do while i > 10
      response.write( "i : " & i )
      i = i + 1
    loop
    </code>
  </pre>
  <%
    do
      response.write( "i : <strong>" & i & "</strong><br/>" )
      i = i + 1
    loop while i < 10
  %>
</body>
</html>
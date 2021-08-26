<% @CODEPAGE="65001" language="vbscript" %>
<% Response.CharSet = "utf-8" %>
<!DOCTYPE html>
<html>
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
</head>
<body>
  <h1>반복문</h1>
  <h3>for ... to</h3>
  <pre>
    <code>
      for i = 0 to 5
        response.write( "i : " & i )
      next
    </code>
  </pre>
  <%
    for i = 0 to 5
      response.write( "i : <strong>" & i & "</strong><br/>" )
    next
  %>

  <hr/>
  
  <h3>step</h3>
  <pre>
    <code>
      for i = 0 to 10 step 2
        response.write( "i : " & i )
      next
    </code>
  </pre>
  <%
    for i = 0 to 10 step 2
      response.write( "i : <strong>" & i & "</strong><br/>" )
    next
  %>

  <hr/>

  <h3>- step</h3>
  <pre>
    <code>
      for i = 10 to 2 step -2
        response.write( "i : " & i )
      next
    </code>
  </pre>
  <%
    for i = 10 to 2 step -2
      response.write( "i : <strong>" & i & "</strong><br/>" )
    next
  %>
</body>
</html>
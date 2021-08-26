<% @CODEPAGE="65001" language="vbscript" %>
<% Response.CharSet = "utf-8" %>
<!DOCTYPE html>
<html>
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
</head>
<body>
  <h3>for each ... next</h3>
  <pre>
    <code>
    dim cars(2)
    cars(0) = "Volvo"
    cars(1) = "SAAB"
    cars(2) = "BMW"

    for each e in cars
      response.write( e & \"<br/>\" )
    next
    </code>
  </pre>
  <%
    dim cars(2)
    cars(0) = "Volvo"
    cars(1) = "SAAB"
    cars(2) = "BMW"

    for each e in cars
      response.write( e & "<br/>" )
    next
  %>
</body>
</html>
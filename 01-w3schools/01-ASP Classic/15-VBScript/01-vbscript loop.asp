<% @CODEPAGE="65001" language="vbscript" %>
<% 'session.CodePage = "65001" %>
<% 'Response.CharSet = "utf-8" %>
<% Response.buffer = true %>
<% Response.Expires = 0 %>
<!DOCTYPE html>
<html lang="ko">
<head>
  <!-- <meta charset="UTF-8"> -->
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <style>
    * { background-color : black; color : white; margin : 1em; }
  </style>
  <title>VBScript</title>
</head>
<body>
  <head>
    There are <%=application( "visitors" )%> on online now!
  </head>
  <br/><hr/><br/>
  <h1>VBScript loop</h1>

  <br/><hr/><br/>

  <h2> for ... next </h2>
  <code>
    <pre>
      for i = 0 to 5
        response.write( "The Number Is : " & i & "&lt;br/&gt;" )
      next
    </pre>
  </code>

  <%
    for i = 0 to 5
      response.write( "The Number Is : " & i & "<br/>" )
    next
  %>

  <br/><hr/><br/>

  <h2> for ... next </h2>
  <code>
    <pre>
      dim cars(2)
      cars(0) = "Porsche"
      cars(1) = "BMW"
      cars(2) = "Jaguar"

      for each x in cars
        response.write( x & "&lt;br/&gt;" )
      next
    </pre>
  </code>

  <%
    dim cars(2)
    cars(0) = "Porsche"
    cars(1) = "BMW"
    cars(2) = "Jaguar"

    for each x in cars
      response.write( x & "<br/>" )
    next
  %>

  <br/><hr/><br/>

  <h2> do while </h2>
  <code>
    <pre>
      i = 0
      do while i < 10
        response.write( i & "&gt;br/&lt;" )
        i = i + 1
      loop
    </pre>
  </code>

  <%
    i = 0
    do while i < 10
      response.write( i & "<br/>" )
      i = i + 1
    loop
  %>

</body>
</html>
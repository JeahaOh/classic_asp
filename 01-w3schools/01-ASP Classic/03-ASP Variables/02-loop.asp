<% @CODEPAGE="65001" language="vbscript" %>
<!DOCTYPE html>
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
</head>
<html>
<body>
  <%
    dim i
    for i = 1 to 6
      response.write( "<h" & i & ">Heading " & i & "</h" & i & ">" )
    next
  %>
</body>
</html>
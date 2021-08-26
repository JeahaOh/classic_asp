<% @CODEPAGE="65001" language="vbscript" %>
<!DOCTYPE html>
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
</head>
<html>
<body>
  <%
    dim famname(5), i
    famname(0) = "Tov Lo"
    famname(1) = "Flume"
    famname(2) = "Heize"
    famname(3) = "쓰레기 같은"
    famname(4) = "ASP"
    famname(5) = "Borge"

    for i = 0 to 5
      response.write(famname(i) & "<br>")
    next
  %>
</body>
</html>
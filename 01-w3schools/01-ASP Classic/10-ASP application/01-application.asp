<% @CODEPAGE="65001" language="vbscript" %>
<% session.CodePage = "65001" %>
<% Response.CharSet = "utf-8" %>
<% Response.buffer=true %>
<% Response.Expires = 0 %>
<!DOCTYPE html>
<html>
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
  <style>
    * { background-color : black; color : white; margin : 1em; }
  </style>
</head>
<body>
  There are <%=application( "users" )%> active connections.

  <br/><hr/><br/>

  <h2>컨텐츠 콜렉션 반복문</h2>
  <%
    dim ac
    for each ac in application.contents
      response.write( ac & "<br/>" )
    next
  %>

  <br/><hr/><br/>

  <h2>컨텐츠 콜렉션 count</h2>
  <%
    dim acCnt
    acCnt = application.contents.count
    response.write( "application.contents.count : " & acCnt & "<br/>" )

    for ac = 1 to acCnt
      response.write( application.contents( ac ) & "<br/>" )
    next
  %>

  <br/><hr/><br/>

  <h2>StaticObjects collection을 통해 반복문</h2>
  <%
    dim so
    
    for each so in application.StaticObjects
      response.write( so & "<br/>" )
    next
  %>

</body>
</html>
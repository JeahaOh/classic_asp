<% @CODEPAGE="65001" language="vbscript" %>
<% session.CodePage = "65001" %>
<% Response.CharSet = "utf-8" %>
<% Response.buffer = true %>
<% Response.Expires = 0 %>
<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8">
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <style>
    * { background-color : black; color : white; margin : 1em; }
  </style>
  <title> session OBJECT </title>
</head>
<body>
  <head>
    There are <%=application( "visitors" )%> on online now!
  </head>

  <br/><hr/><br/>

  <h2> Return session id number for a user </h2>
  session id number : <%=Session.SessionID%>
  <br/>

  <br/><hr/><br/>

  <h2> Return session id number for a user </h2>
  <% 
    response.write( "<p>" )
    response.write( "  Default Timeout is: " & Session.Timeout & " minutes." )
    response.write( "</p>" )

    Session.Timeout = 30

    response.write( "<p>" )
    response.write( "Timeout is now: " & Session.Timeout & " minutes." )
    response.write( "</p>" )
  %>
  <br/>




</body>
</html>
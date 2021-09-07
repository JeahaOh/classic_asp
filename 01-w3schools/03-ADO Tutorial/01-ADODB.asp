<% @CODEPAGE="65001" language="vbscript" %>
<%
  session.CodePage = "65001"
  Response.CharSet = "utf-8"
  response.buffer = true
  response.expires = -1
  'response.expiresAbsolute = #Aug 31,2021 09:30:00#
  response.contentType = "text/html"
%>
<!DOCTYPE html>
<html lang="ko">
<head>
  <!-- <meta charset="UTF-8"> -->
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <style>
    * { background-color : black; color : white; margin : 1em; }
  </style>
  <title> ADODB CONNECTION </title>
</head>
<body>
  <head>
    <p>There are <%=application( "visitors" )%> on online now!</p>
  </head>

  <br/><hr/><br/>

  <h2> ADODB CONNECTION </h2>
  <p></p>

  <%
    dim db_conn, query, rs

    set db_conn = server.createObject( "ADODB.connection" )
    db_conn.open application( "db_conn" )

    query = "SELECT now() AS now FROM DUAL"

    set rs = db_conn.execute( query )

    response.write( rs("NOW") )

    set rs = nothing
    set db_conn = nothing
  %>


  <br/><hr/><br/>
  

</body>
</html>
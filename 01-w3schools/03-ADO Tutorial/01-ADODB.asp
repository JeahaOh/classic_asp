<% @CODEPAGE="65001" language="vbscript" %>
<%
  session.CodePage = "65001"
  ' Response.CharSet = "utf-8"
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
  <title>VBScript</title>
</head>
<body>
  <head>
    <p>There are <%=application( "visitors" )%> on online now!</p>
  </head>

  <br/><hr/><br/>

  <h2> ADODB CONNECTION </h2>
  <p></p>

  <%
    set conn = server.createObject( "ADODB.connection" )
    ' DB Driver 제공자와,
    conn.provider = "Microsoft.Jet.OLEDB.4.0"
    ' DB에 대한 물리적 경로를 지정해야 함.
    conn.open "c:/webdata/northwind.mdb"
    ' conn.open "northwind"
  %>

  <br/><hr/><br/>
  

</body>
</html>
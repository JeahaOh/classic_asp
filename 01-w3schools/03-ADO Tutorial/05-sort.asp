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
  <meta charset="UTF-8">
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <style>
    * { background-color : black; color : white; margin : 1em; }
  </style>
  <title> ADO record set </title>
</head>
<body>
  <head>
    <p>There are <%=application( "visitors" )%> on online now!</p>
  </head>

  <br/><hr/><br/>

  <h2> ADO record set </h2>
  <p></p>

  <%
    dim db_conn, query, rs

    set db_conn = server.createObject( "ADODB.connection" )
    db_conn.open application( "db_conn" )

    query = ""
    query = query & " SELECT "
    query = query & "  CUST_ID "
    query = query & "  , COMP_NM "
    query = query & "  , FRST_REG_DE "
    query = query & " FROM CUSTOMERS "
    query = query & " WHERE COMP_NM LIKE '%미%' "
    query = query & " ORDER BY COMP_NM "
    
    set rs = server.createObject( "ADODB.recordset" )
    rs.open query, db_conn 
  %>

  <table border="1" width="100%" bgcolor="#FFF5EE">
    <thead>
      <tr>
        <% for each x in rs.fields %>
          <th align='left' bgcolor='#B0C4DE'><%=x.name%></th>
        <% next %>
      </tr>
    </thead>
    <tbody>
      <% do until rs.eof %>
        <tr>
          <% for each x in rs.fields %>
            <td><%=x.value%></td>
          <% next %>
          <% rs.moveNext %>
        </tr>
      <% loop %>
    </tbody>
  </table>

  <%
    rs.close
    db_conn.close
  %>

</body>
</html>
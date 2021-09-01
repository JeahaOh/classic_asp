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
  <meta charset="UTF-8">
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <style>
    * { background-color : black; color : white; margin : 1em; }
  </style>
  <title>ASP File Object</title>
</head>
<body>
  <head>
    There are <%=application( "visitors" )%> on online now!
  </head>

  <br/><hr/><br/>

  <h2> ASP File Object </h2>

  <%
    dim fs, f
    set fs = server.createObject( "Scripting.fileSystemObject" )
    set f = fs.getFile( server.mapPath( "text/textfile.txt" ) )

    response.write( "<br/>the file was created on : " & f.dateCreated )
    response.write( "<br/>the file was last modified on : " & f.dateLastModified )
    response.write( "<br/>the file was last modified on : " & f.dateLastModified )
    response.write( "<br/>the file was last accessed on : " & f.dateLastAccessed )
    response.write( "<br/>the attributes of the file are : " & f.attributes )


    set f = nothing
    set fs = nothing

  %>

  <br/><hr/><br/>

</body>
</html>
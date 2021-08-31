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

  <h2> When was a file last modified? </h2>

  <%
    Set fs = Server.CreateObject("Scripting.FileSystemObject")
    Set rs = fs.GetFile(Server.MapPath("example.asp"))
    modified = rs.DateLastModified
  %>
    This file was last modified on: 
  <%
    response.write(modified)
    Set rs = Nothing
    Set fs = Nothing
  %>

  <br/><hr/><br/>

  <h2> Open a textfile for reading </h2>

  <%
    Set FS = Server.CreateObject( "Scripting.FileSystemObject" )
    Set RS = FS.OpenTextFile( Server.MapPath("text") & "\TextFile.txt", 1 )
    ' Set RS = FS.OpenTextFile( "\example.asp", 1 )
    While not rs.AtEndOfStream
      Response.Write RS.ReadLine
      Response.Write("<br>")
    Wend
  %>

  <p>
    <a href="text/textfile.txt">btn_view_text</a>
  </p>

</body>
</html>
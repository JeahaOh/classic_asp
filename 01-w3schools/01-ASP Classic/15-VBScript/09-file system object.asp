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
  <title> file system object </title>
</head>
<body>

  <head>
    There are <%=application( "visitors" )%> on online now!
  </head>


  <br/><hr/><br/>

  <h2> Does a specified file exist? </h2>

  <%
    dim filepath, filename
    filepath = "c:\asdf\asdf\"
    filename = "asdf.txt"
    set fs = server.createObject( "scripting.fileSystemObject" )

    if( fs.fileExists( filepath & filename ) ) = true then
      response.write( filepath & filename & " is exists." )
    else
      response.write( filepath & filename & " does not exists." )
    end if

    set fs = nothing
  %>

  <br/><hr/><br/>

  <h2> Does a specified folder exist? </h2>

  <%
    filepath = "c:\Users"

    set fs = server.createObject( "scripting.fileSystemObject" )

    if( fs.folderExists( filepath ) ) = true then
      response.write( "Dir " & filepath & " is exists." )
    else
      response.write( "Dir " & filepath & " does not exists." )
    end if

    set fs = nothing
  %>

  <br/><hr/><br/>

  <h2> Does a specified drive exist? </h2>

  <%
    set fs = server.createObject( "scripting.fileSystemObject" )

    if( fs.driveExists( "c:" ) ) = true then
      response.write( "drive c: is exists." )
    else
      response.write( "drive c: does not exists." )
    end if

    response.write( "<br/>" )

    if( fs.driveExists( "g:" ) ) = true then
      response.write( "drive g: is exists." )
    else
      response.write( "drive g: does not exists." )
    end if

    set fs = nothing
  %>

  <br/><hr/><br/>

  <h2> Get the name of a specified drive </h2>

  <%
    set fs = server.createObject( "scripting.fileSystemObject" )

    p = fs.getDriveName( "c:\" )

    response.write( "the drive name is : " & p )

    set fs = nothing
  %>

  <br/><hr/><br/>

  <h2> Get the name of the parent folder of a specified path </h2>

  <%
    set fs = server.createObject( "Scripting.FileSystemObject" )
    p = fs.getParentFolderName( "c:\winnt\cursors\3dgarro.cur" )

    response.write( "The parent folder name of c:\winnt\cursors\3dgarro.cur is: " & p )

    set fs = nothing
  %>

  <br/><hr/><br/>

  <h2> Get the file extension </h2>

  <%
    set fs = server.createObject( "Scripting.FileSystemObject" )

    response.write( "The file extension of the file 3dgarro is: " )
    response.write( fs.getExtensionName( "c:\winnt\cursors\3dgarro.cur" ) )

    set fs = nothing
  %>

  <br/><hr/><br/>

  <h2> Get the base name of a file or folder </h2>

  <%
    set fs = server.createObject( "Scripting.FileSystemObject" )

    response.write( fs.getBaseName( "c:\winnt\cursors\3dgarro.cur" ) )
    response.write( "<br>" )
    response.write( fs.getBaseName( "c:\winnt\cursors\" ) )
    response.write( "<br>" )
    response.write( fs.getBaseName( "c:\winnt\" ) )

    set fs = nothing
  %>



</body>
</html>
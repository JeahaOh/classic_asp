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
  <title> drive object </title>
</head>
<body>

  <head>
    There are <%=application( "visitors" )%> on online now!
  </head>

  <br/><hr/><br/>

  <h2> Get the available space of a specified drive </h2>

  <%
    dim fs, d
    set fs = server.createObject( "Scripting.fileSystemObject" )
    set d = fs.getDrive( "c:" )

    n = "Drive : " & d
    n = n & "<br/> Available Space in bytes : " & d.availableSpace
    
    response.write( n )

    set d = nothing
    set fs = nothing
  %>

  <br/><hr/><br/>

  <h2> Get the free space of a specified drive </h2>

  <%
    set fs = server.createObject( "Scripting.fileSystemObject" )
    set d = fs.getDrive( "c:" )

    n = "Drive : " & d
    n = n & "<br/> Free Space in bytes : " & d.freeSpace
    
    response.write( n )

    set d = nothing
    set fs = nothing
  %>


  <br/><hr/><br/>

  <h2> Get the total size of a specified drive </h2>

  <%
    set fs = server.createObject( "Scripting.fileSystemObject" )
    set d = fs.getDrive( "c:" )

    n = "Drive : " & d
    n = n & "<br/> Total size in bytes : " & d.totalsize
    
    response.write( n )

    set d = nothing
    set fs = nothing
  %>

  <br/><hr/><br/>

  <h2> Get the drive letter of a specified drive </h2>

  <%
    set fs = server.createObject( "Scripting.fileSystemObject" )
    set d = fs.getDrive( "c:" )

    response.write( "The drive letter is " & d.driveletter )

    set d = nothing
    set fs = nothing
  %>

  <br/><hr/><br/>

  <h2> Get the drive type of a specified drive </h2>

  <%
    set fs = server.createObject( "Scripting.fileSystemObject" )
    set d = fs.getDrive( "c:" )

    response.write( "The drive type is " & d.driveType )

    set d = nothing
    set fs = nothing
  %>

  <br/><hr/><br/>

  <h2> Get the file system of a specified drive </h2>

  <%
    set fs = server.createObject( "Scripting.fileSystemObject" )
    set d = fs.getDrive( "c:" )

    response.write( "The file system is " & d.fileSystem )

    set d = nothing
    set fs = nothing
  %>

  <br/><hr/><br/>

  <h2> Is the drive ready? </h2>

  <%
    set fs = server.createObject( "Scripting.fileSystemObject" )
    set d = fs.getDrive( "c:" )

    n = "the " & d.driveLetter
    
    if d.isReady = true then
      n = n & " drive is ready."
    else
      n = n & " drive is not ready."
    end if

    response.write n

    set d = nothing
    set fs = nothing
  %>

  <br/><hr/><br/>

  <h2> Get the path of a specified drive </h2>

  <%
    set fs = server.createObject( "Scripting.fileSystemObject" )
    set d = fs.getDrive( "c:" )

    response.write( "The path is " & d.path )

    set d = nothing
    set fs = nothing
  %>

  <br/><hr/><br/>

  <h2> Get the root folder of a specified drive </h2>


  <%
    set fs = server.createObject( "Scripting.fileSystemObject" )
    set d = fs.getDrive( "c:" )

    response.write( "The root folder is " & d.rootFolder )

    set d = nothing
    set fs = nothing
  %>

  <br/><hr/><br/>

  <h2> Get the serial number of a specified drive </h2>

  <%
    set fs = server.createObject( "Scripting.fileSystemObject" )
    set d = fs.getDrive( "c:" )

    response.write( "The serialnumber is " & d.serialnumber )

    set d = nothing
    set fs = nothing
  %>

  <br/><hr/><br/>

</body>
</html>
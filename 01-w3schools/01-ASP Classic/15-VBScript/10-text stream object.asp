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
  <title> text stream object </title>
</head>
<body>

  <head>
    There are <%=application( "visitors" )%> on online now!
  </head>


  <br/><hr/><br/>

  <h2> Read textfile </h2>

  <p>This is the text in the text file:</p>
  <%
    set fs = server.createObject( "Scripting.FileSystemObject" )
    set f = fs.openTextFile( Server.MapPath( "text/textfile.txt" ), 1 )

    response.write( f.readAll )
    f.Close

    set f = nothing
    set fs = nothing
  %>


  <br/><hr/><br/>

  <h2> Read only a part of a textfile </h2>

  <p>This is the first five characters from the text file:</p>
  <%
    set fs = server.createObject( "Scripting.FileSystemObject" )
    set f = fs.openTextFile( server.mapPath( "text/textfile.txt" ), 1 )

    response.write( f.read(5) )

    set f = nothing
    set fs = nothing
  %>


  <br/><hr/><br/>

  <h2> Read one line of a textfile </h2>

  <p>This is the first line of the text file:</p>
  <%
    set fs = server.createObject( "Scripting.FileSystemObject" )
    set f = fs.openTextFile( server.mapPath( "text/textfile.txt" ), 1 )

    response.write( f.readLine )

    f.close
    set f = nothing
    set fs = nothing
  %>


  <br/><hr/><br/>

  <h2> Read all lines from a textfile </h2>

  <p>This is all the lines in the text file:</p>
  <%
    set fs = server.createObject( "Scripting.FileSystemObject" )
    set f = fs.openTextFile( server.mapPath( "text/textfile.txt" ), 1 )

    do while f.atEndOfStream = false
      response.write( f.readline )
      response.write( "<br/>" )
    loop

    f.close
    set f = nothing
    set fs = nothing
  %>



  <br/><hr/><br/>

  <h2> Skip a part of a textfile </h2>

  <p>The first four characters in the text file are skipped:</p>
  <%
    set fs = server.createObject( "Scripting.FileSystemObject" )
    set f = fs.openTextFile( server.mapPath( "text/textfile.txt" ), 1 )

    f.skip( 14 )
    response.write( f.readAll )

    f.close
    set f = nothing
    set fs = nothing
  %>



  <br/><hr/><br/>

  <h2> Skip a line of a textfile </h2>

  <p>The first line in the text file is skipped:</p>
  <%
    set fs = server.createObject( "Scripting.FileSystemObject" )
    set f = fs.openTextFile( server.mapPath( "text/textfile.txt" ), 1 )

    f.skipLine

    do while f.atEndOfStream = false
      response.write( f.readline )
      response.write( "<br/>" )
    loop

    f.close
    set f = nothing
    set fs = nothing
  %>


  <br/><hr/><br/>

  <h2> Return current line-number in a text file </h2>

  <p>This is all the lines in the text file (with line numbers):</p>
  <ul>
    <%
      set fs = server.createObject( "Scripting.FileSystemObject" )
      set f = fs.openTextFile( server.mapPath( "text/textfile.txt" ), 1 )
      
      do while f.atEndOfStream = false
        response.write( "<li>" & f.line & " " )
        response.write( f.readline )
        response.write( "</li>" )
      loop

      f.close
      set f = nothing
      set fs = nothing
    %>
  </ul>


  <br/><hr/><br/>

  <h2> Get column number of the current character in a text file </h2>
  
  <%
    set fs = server.createObject( "Scripting.FileSystemObject" )
    set f = fs.openTextFile( server.mapPath( "text/textfile.txt" ), 1 )
    
    response.write( f.read( 20 ) )
    response.write( "<p>The cursor is now standing in position " & f.column & " in the text file.</p>" )

    f.close
    set f = nothing
    set fs = nothing
  %>


  <br/><hr/><br/>

</body>
</html>
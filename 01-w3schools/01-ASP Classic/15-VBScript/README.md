# VBScript

## VBScript 반복문

### for ... next

```script
  for i = 0 to 5
    response.write( "The Number Is : " & i & "<br/>" )
  next
```

### for ... each

```script
  dim cars(2)
  cars(0) = "Porsche"
  cars(1) = "BMW"
  cars(2) = "Jaguar"

  for each x in cars
    response.write( x & "<br/>" )
  next
```

### do while

```script
  i = 0
  do while i < 10
    response.write( i & "<br/>" )
    i = i + 1
  loop
```

  ---  
  ---  

## VBScript 날짜와 시간

### date and time

```script
  <%=date()%>
  <%=time()%>
```

### Get The Name Of A Day

```script
  <p>
    VBScripts' function <b>WeekdayName</b> is used to get a weekday:
  </p>
  <%
    response.Write(WeekDayName(1))
    response.Write("<br>")
    response.Write(WeekDayName(2))
  %>

  <p>Abbreviated name of a weekday:</p>
  <%
    response.Write(WeekDayName(1,true))
    response.Write("<br>")
    response.Write(WeekDayName(2,true))
  %>

  <p>Current weekday:</p>
  <%
    response.Write(WeekdayName(weekday(date)))
    response.Write("<br>")
    response.Write(WeekdayName(weekday(date), true))
  %>
```

### Get the name of a month

```script
  Tody it is <%=weekdayName( weekday( date ) )%>
  <br/>
  and the month is <%=monthName( month( date ) )%>
```

### Countdown to year 3000

```script
  <p>Countdown to year 3000 : </p>
  <%
    millennium = cdate( "1/1/3000 00:00:00" )
  %>

  It is <%=dateDiff( "yyyy", now(), millennium )%> years to year 3000!
  <br/>
  It is <%=dateDiff( "m", now(), millennium )%> months to year 3000!
  <br/>
  It is <%=dateDiff( "ww", now(), millennium )%> weeks to year 3000!
  <br/>
  It is <%=dateDiff( "d", now(), millennium )%> days to year 3000!
  <br/>
  It is <%=dateDiff( "h", now(), millennium )%> hours to year 3000!
  <br/>
  It is <%=dateDiff( "n", now(), millennium )%> minutes to year 3000!
  <br/>
  It is <%=dateDiff( "s", now(), millennium )%> seconds to year 3000!
  <br/>
```

### Calculate the day which is n days from today

```script
  <%=dateAdd( "d", 30, date() )%> is a date 30 days from today.
```

### Format date and time

```script
  <%=formatDateTime( date(), vbgeneraldate )%> <br/>
  <%=formatDateTime( date(), vblongdate )%> <br/>
  <%=formatDateTime( date(), vbshortdate )%> <br/>
  <%=formatDateTime( now(), vblongtime )%> <br/>
  <%=formatDateTime( now(), vbshorttime )%> <br/>
```

### is this a date?

```script
  somedate = "04/27/92"
  response.write( isDate( somedate ) & "<br/>" )
  somedate = "04-27-92"
  response.write( isDate( somedate ) & "<br/>" )
  somedate = "1992-04-27"
  response.write( isDate( somedate ) & "<br/>" )
```

  ---  
  ---  

## VBScript 내장함수

### Uppercase or Lowercase a string

```script
  dim name
  name = "Asdf Qwer"
  
  ucase( name )
  lcase( name )
```

### Trim a string

```asp
  name = "  Jeaha . dev  "

  response.write("visit" & name & "now<br/>")
  response.write("visit" & trim(name) & "now<br/>")
  response.write("visit" & ltrim(name) & "now<br/>")
  response.write("visit" & rtrim(name) & "now<br/>")
```

### How to reverse a string?

```asp
  txt = "Hell the Everyone!"
  response.write( strReverse( txt ) )
```

### How to round a number?

```asp
  i = 48.66776677
  j = 48.3333333
  response.write(Round(i))
  response.write("<br>")
  response.write(Round(j))
```

### A random number

```asp
  ' 난수 생성
  randomize()
  r = rnd()
  response.write("A random number: <b>" & r & "</b><br/>")

  ' 0 ~ 100 사이의 난수
  randomNumber = Int( 100 * r )
  response.write("A random number: <b>" & randomNumber & "</b><br/>")
```

### Return a specified number of characters from left/right of a string

```asp
  txt = "Helcome to this WEB"
  response.write( txt )
  response.write( "<br/>" )
  response.write( left( txt, 5 ) )
  response.write( "<br/>" )
  response.write( right( txt, 5 ) )
```

### Replace some characters in a string

```asp
  response.write( replace( txt, "WEB", "HELL!!" ) )
```

### Return a specified number of characters from a string

```asp
  response.write( mid( txt, 9, 2 ) )
```
  

  ---  
  ---  
  
## Procedure
  
### Sub procedure
  
```asp
  Sub mysub()
    response.write("I was written by a sub procedure")
  End Sub

  response.write("I was written by the script<br>")
  Call mysub()
```
  

### function procedure
  
```asp
  Function myfunction()
    myfunction = Date()
  End Function

  response.write("Today's date: ")
  response.write(myfunction())
```

  ---  
  ---  
  
## ASP Response Object
  
### Redirect the user to another URL
  
```asp
  <%
    if request.form( "select" ) <> "" then
      response.redirect( request.form( "select" ) )
    end if
  %>

  <form action="example.asp" method="post">
    <label for="select_server.asp">SERVER</label>
    <input
      type="radio" name="select" id="select_server.asp"
      value="select_server.asp"
    />
    <br/>
    <label for="select_text.asp">TEXT</label>
    <input
      type="radio" name="select" id="select_text.asp"
      value="select_text.asp"
    />
    <input type="submit" value="GO!" />
  </form>
```

### Random Links

```asp
  randomize()
  r = rnd()

  if r > 0.5 then
    response.write( "<a href='https://www.w3schools.com/asp/asp_examples.asp'>w3school asp</a>" )
  else
    response.write( "<a href='javascript:void(0)'>nowhere</a>" )
  end if
```

### Controlling Buffer

```html
  <% Response.Buffer = true %>
  <!DOCTYPE html>
  <html>
  <body>
    <p>
      This text will be sent to your browser when my response buffer is flushed.
    </p>
    <% Response.Flush %>
  </body>
  </html>
```

### Clear the buffer

```html
  <% Response.Buffer = true %>
  <!DOCTYPE html>
  <html>
  <body>
    <p>This is some text I want to send to you.</p>
    <p>No, I changed my mind. I wnat to clear the text.</p>
    <% response.clear %>
  </body>
  </html>
```

### End a script in the middle of processing

```asp
  I am writing some text. This text will never be
  <br>
  <% response.end %>
  finished! It's too late to write more!
```

### Set how many minutes a page will be cached in a browser before it expires

```asp
  <% response.expires = -1 %>
  <p>This page will be refreshed with each access!</p>
```

### Set a date/time when a page cached in a browser will expire

```asp
  response.expiresAbsolute = #Aug 31,2021 09:30:00#
```

### Check if the user is still connected

```asp
  if response.isClientConnected = true then
    response.write( "the user is still connected!" )
  else
    response.write( "the user is NOT connected!" )
  end if
```

### Set the type of content

```asp
  response.contentType = "text/html" 
```

### Set the name of character set

```asp
  response.charset = "utf-8"
```

  ---  
  ---  

## ASP Request Object

### Send extra information within a link

```asp
  <a href="example.asp?color=green">Example</a>
  <%=Request.QueryString%>
```

### Send extra information within a link

```asp
  <a href="example.asp?color=green">Example</a>
  <%=Request.QueryString%>
```

### A QueryString collection in its simplest use

```asp
  <form action="example.asp" method="get">
    <label for="fname">First name : </label>
    <input
      type="text" name="fname" id="fname"
    />
    <br/>
    <label for="lname">Last name : </label>
    <input
      type="text" name="lname" id="lname"
    />
    <br/>
    <input type="submit" value="Submit" />
  </form>

  <%=Request.QueryString%>
```

### How to use information from forms

```asp
  <form action="example.asp" method="get">
    <label for="name">Your name : </label>
    <input type="text" name="name" id="name" size="20" />
    <input type="submit" value="submit" />
  </form>

  <%
    dim name, text
    name = request.querystring( "name" )

    if name <> "" then
      response.write( "Hell the " & name & "!!<br/>" )
    end if
  %>
```

### More information from a form

```asp
  <form action="example.asp" method="post">
    <label for="first">first name : </label>
    <input
      type="text" name="name" id="first"
      size="20" value="FIRST"
    />
    <br/>
    <label for="last">last name : </label>
    <input
      type="text" name="name" id="last"
      size="20" value="LAST"
    />
    <input type="submit" value="submit" />
  </form>

  <p>The information received from the form above was : </p>
  <%
    if request.form( "name" ) <>"" then
      %>
      <p>name = <%=request.form( "name" )%></p>
      <p>the name property's count = <%=request.form( "name" ).count%></p>
      <p>first name = <%=request.form( "name" )(1)%></p>
      <p>last name = <%=request.form( "name" )(2)%></p>
      <%
    end if
  %>
```

### A form with checkboxes


```asp
  <%
    fruits = request.form( "fruits" )
  %>

  <form action="example.asp" method="post">
    <input
      type="checkbox" name="fruits" id="apple" value="apple"
      <%if instr( fruits, "apple" ) then response.write( "checked" )%>
    />
    <label for="apple">apple</label>
    <br/>
    <input
      type="checkbox" name="fruits" id="kiwi" value="kiwi"
      <%if instr( fruits, "kiwi" ) then response.write( "checked" )%>
    />
    <label for="kiwi">kiwi</label>
    <br/>
    <input
      type="checkbox" name="fruits" id="banana" value="banana"
      <%if instr( fruits, "banana" ) then response.write( "checked" )%>
    />
    <label for="banana">banana</label>
    <br/>
    <input type="submit" value="submit" />
  </form>

  <%
    if fruits <> "" then %>
    <p>SELECTED : <%=fruits%></p>
  <%end if%>
```

### How to find the visitors' browser type, IP address and more

```asp
  <p>
    <b>You are browsing this site with:</b>
    <%= request.serverVariables( "http_user_agent" )%>
  </p>
  <p>
    <b>Your IP address is:</b>
    <%= request.serverVariables( "remote_addr" )%>
  </p>
  <p>
    <b>The DNS lookup of the IP address is:</b>
    <%= request.serverVariables( "remote_host" )%>
  </p>
  <p>
    <b>The method used to call the page:</b>
    <%= request.serverVariables( "request_method" )%>
  </p>
  <p>
    <b>The server's domain name:</b>
    <%= request.serverVariables( "server_name" )%>
  </p>
  <p>
    <b>The server's port:</b>
    <%= request.serverVariables( "server_port" )%>
  </p>
  <p>
    <b>The server's software:</b>
    <%= request.serverVariables( "server_software" )%>
  </p>
```

### List all servervariables you can ask for

```asp
  <table>
    <thead>
      <tr>
        <th>item</th>
        <th>value</th>
      </tr>
    </thead>
    <tbody>
      <tr>
        <td></td>
        <td></td>
      </tr>
      <%
        For Each Item in Request.ServerVariables
          response.write( "<tr>" )
            Response.Write( "<td>" & Item & "</tdr>")
            Response.Write( "<td>" & Request.ServerVariables(Item) & "</td>")
          response.write( "</tr>" )
        Next
      %>
    </tbody>
  </table>
```

### Total number of bytes the user sent

```asp
  Total bytes : <%=request.totalbytes%>
```

## Session Object

### Return session id number for a user

```asp
  session id number : <%=Session.SessionID%>
```

### Get a session's timeout

```asp
  <% 
    response.write( "<p>" )
    response.write( "  Default Timeout is: " & Session.Timeout & " minutes." )
    response.write( "</p>" )

    Session.Timeout = 30

    response.write( "<p>" )
    response.write( "Timeout is now: " & Session.Timeout & " minutes." )
    response.write( "</p>" )
  %>
```

## Server Object

### When was a file last modified?

```asp
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
```

### Open a textfile for reading

```asp
  <%
    Set FS = Server.CreateObject( "Scripting.FileSystemObject" )
    Set RS = FS.OpenTextFile( Server.MapPath("text") & "\TextFile.txt", 1 )
    ' Set RS = FS.OpenTextFile( "\example.asp", 1 )
    While not rs.AtEndOfStream
      Response.Write RS.ReadLine
      Response.Write("<br>")
    Wend
  %>
```

## File System Object

### Does a specified file exist?

```asp

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
  
```

### Does a specified folder exist?

```asp

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
  
```

### Does a specified drive exist?

```asp

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
  
```

### Get the name of a specified drive

```asp

  <%
    set fs = server.createObject( "scripting.fileSystemObject" )

    p = fs.getDriveName( "c:\" )

    response.write( "the drive name is : " & p )

    set fs = nothing
  %>
  
```

### Get the name of the parent folder of a specified path


```asp

  <%
    set fs = server.createObject( "Scripting.FileSystemObject" )
    p = fs.getParentFolderName( "c:\winnt\cursors\3dgarro.cur" )

    response.write( "The parent folder name of c:\winnt\cursors\3dgarro.cur is: " & p )

    set fs = nothing
  %>
  
```

### Get the file extension

```asp

  <%
    set fs = server.createObject( "Scripting.FileSystemObject" )

    response.write( "The file extension of the file 3dgarro is: " )
    response.write( fs.getExtensionName( "c:\winnt\cursors\3dgarro.cur" ) )

    set fs = nothing
  %>
  
```

### Get the base name of a file or folder

```asp

  <%
    set fs = server.createObject( "Scripting.FileSystemObject" )

    response.write( fs.getBaseName( "c:\winnt\cursors\3dgarro.cur" ) )
    response.write( "<br>" )
    response.write( fs.getBaseName( "c:\winnt\cursors\" ) )
    response.write( "<br>" )
    response.write( fs.getBaseName( "c:\winnt\" ) )

    set fs = nothing
  %>
  
```

## text stream object

### Read textfile

```asp
  <p>This is the text in the text file:</p>
  <%
    set fs = server.createObject( "Scripting.FileSystemObject" )
    set f = fs.openTextFile( Server.MapPath( "text/textfile.txt" ), 1 )

    response.write( f.readAll )
    f.Close

    set f = nothing
    set fs = nothing
  %>
```

### Read only a part of a textfile

```asp
  <p>This is the first line of the text file:</p>
  <%
    set fs = server.createObject( "Scripting.FileSystemObject" )
    set f = fs. openTextFile( server.mapPath( "text/textfile.txt" ), 1 )

    response.write( f.readLine )

    f.close
    set f = nothing
    set fs = nothing
  %>
```

### Read one line of a textfile

```asp
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
```

### Read all lines from a textfile

```asp
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
```

### Skip a part of a textfile

```asp
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
```

### Skip a line of a textfile

```asp
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
```

### Return current line-number in a text file

```asp
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
```

### Get column number of the current character in a text file

```asp
  <%
    set fs = server.createObject( "Scripting.FileSystemObject" )
    set f = fs.openTextFile( server.mapPath( "text/textfile.txt" ), 1 )
    
    response.write( f.read( 20 ) )
    response.write( "<p>The cursor is now standing in position " & f.column & " in the text file.</p>" )

    f.close
    set f = nothing
    set fs = nothing
  %>
```

## Drive Object

### Get the available space of a specified drive

```asp
  dim fs, d
  set fs = server.createObject( "Scripting.fileSystemObject" )
  set d = fs.getDrive( "c:" )

  n = "Drive : " & d
  n = n & "<br/> Available Space in bytes : " & d.availableSpace
  
  response.write( n )

  set d = nothing
  set fs = nothing
```

### Get the free space of a specified drive

```asp
  set fs = server.createObject( "Scripting.fileSystemObject" )
  set d = fs.getDrive( "c:" )

  n = "Drive : " & d
  n = n & "<br/> Free Space in bytes : " & d.freeSpace
  
  response.write( n )

  set d = nothing
  set fs = nothing
```

### Get the total size of a specified drive

```asp
  set fs = server.createObject( "Scripting.fileSystemObject" )
  set d = fs.getDrive( "c:" )

  n = "Drive : " & d
  n = n & "<br/> Total size in bytes : " & d.totalsize
  
  response.write( n )

  set d = nothing
  set fs = nothing
```

### Get the drive letter of a specified drive

```asp
  set fs = server.createObject( "Scripting.fileSystemObject" )
  set d = fs.getDrive( "c:" )

  response.write( "The drive letter is " & d.driveletter )

  set d = nothing
  set fs = nothing
```

### Get the drive type of a specified drive

```asp
  set fs = server.createObject( "Scripting.fileSystemObject" )
  set d = fs.getDrive( "c:" )

  response.write( "The drive type is " & d.driveType )

  set d = nothing
  set fs = nothing
```

### Get the file system of a specified drive

```asp
  set fs = server.createObject( "Scripting.fileSystemObject" )
  set d = fs.getDrive( "c:" )

  response.write( "The file system is " & d.fileSystem )

  set d = nothing
  set fs = nothing
```

### Is the drive ready?

```asp
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
```

### Get the path of a specified drive

```asp
  set fs = server.createObject( "Scripting.fileSystemObject" )
  set d = fs.getDrive( "c:" )

  response.write( "The path is " & d.path )

  set d = nothing
  set fs = nothing
```

### Get the root folder of a specified drive

```asp
  set fs = server.createObject( "Scripting.fileSystemObject" )
  set d = fs.getDrive( "c:" )

  response.write( "The root folder is " & d.rootFolder )

  set d = nothing
  set fs = nothing
```

### Get the serial number of a specified drive

```asp
  set fs = server.createObject( "Scripting.fileSystemObject" )
  set d = fs.getDrive( "c:" )

  response.write( "The serialnumber is " & d.serialnumber )

  set d = nothing
  set fs = nothing
```

## File Object

### 

```asp
  
```



- 출처 : [https://www.w3schools.com/asp/asp_examples.asp](https://www.w3schools.com/asp/asp_examples.asp)  
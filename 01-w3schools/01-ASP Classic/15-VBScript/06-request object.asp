<% @CODEPAGE="65001" language="vbscript" %>
<% 'session.CodePage = "65001" %>
<% 'Response.CharSet = "utf-8" %>
<% Response.buffer = true %>
<% Response.Expires = 0 %>
<!DOCTYPE html>
<html lang="ko">
<head>
  <!-- <meta charset="UTF-8"> -->
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <style>
    * { background-color : black; color : white; margin : 1em; }
  </style>
  <title> REQUEST OBJECT </title>
</head>
<body>

  <head>
    There are <%=application( "visitors" )%> on online now!
  </head>

  <br/><hr/><br/>

  <h2> Send extra information within a link </h2>
  <code>
    <pre>
      &lt;a href="example.asp?color=green"&gt;Example&lt;/a&gt;
      &lt;%=Request.QueryString%&gt;
    </pre>
  </code>

  <a href="example.asp?color=green">Example</a>
  <%=Request.QueryString%>

  <br/>


  <br/><hr/><br/>

  <h2> A QueryString collection in its simplest use </h2>
  <code>
    <pre>
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
    </pre>
  </code>

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

  <br/>

  <br/><hr/><br/>

  <h2> How to use information from forms </h2>
  <code>
    <pre>
    </pre>
  </code>


  <br/>

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

  <br/><hr/><br/>

  <h2> More information from a form </h2>
  <code>
    <pre>
    </pre>
  </code>

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

  <br/>

  <br/><hr/><br/>

  <h2> A form with checkboxes </h2>
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


  <br/>

  <br/><hr/><br/>

  <h2> How to find the visitors' browser type, IP address and more </h2>

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

  <br/>

  <br/><hr/><br/>

  <h2> List all servervariables you can ask for </h2>

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

  <br/>

  <br/><hr/><br/>

  <h2> Total number of bytes the user sent </h2>

  Total bytes : <%=request.totalbytes%>


  <br/>

</body>
</html>
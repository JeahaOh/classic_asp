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
  <title>VBScript</title>
</head>
<body>

  
  <head>
    There are <%=application( "visitors" )%> on online now!
  </head>

  <br/><hr/><br/>

  <h2> Uppercase or Lowercase a string </h2>
  <code>
    <pre>
      dim name
      name = "Asdf Qwer"
      
      ucase( name )
      lcase( name )
    </pre>
  </code>

  <%
    dim name
    name = "Asdf Qwer"
  %>
  <%=ucase( name )%>
  <br/>
  <%=lcase( name )%>

  <br/><hr/><br/>

  <h2> Trim a string </h2>
  <code>
    <pre>
      name = "  Jeaha . dev  "

      response.write("visit" & name & "now&lt;br/&gt;")
      response.write("visit" & trim(name) & "now&lt;br/&gt;")
      response.write("visit" & ltrim(name) & "now&lt;br/&gt;")
      response.write("visit" & rtrim(name) & "now&lt;br/&gt;")
    </pre>
  </code>

  <%
    name = "  Jeaha . dev  "

    response.write("visit" & name & "now<br/>")
    response.write("visit" & trim(name) & "now<br/>")
    response.write("visit" & ltrim(name) & "now<br/>")
    response.write("visit" & rtrim(name) & "now<br/>")
  %>

  &lt;%=는 작동하지 않는다.<br/>
  visit <%=name%> now. <br/>
  visit <%=trim( name )%> now. <br/>
  visit <%=ltrim( name )%> now. <br/>
  visit <%=rtrim( name )%> now. <br/>

  <br/><hr/><br/>

  <h2> How to reverse a string? </h2>
  <code>
    <pre>
      txt = "Hell the Everyone!"
      response.write( strReverse( txt ) )
    </pre>
  </code>

  <%
    txt = "Hell the Everyone!"
    response.write( strReverse( txt ) )
  %>

  <br/><hr/><br/>

  <h2> How to round a number? </h2>
  <code>
    <pre>
      i = 48.66776677
      j = 48.3333333
      response.write(Round(i))
      response.write("<br>")
      response.write(Round(j))
    </pre>
  </code>

  <%
    i = 48.66776677
    j = 48.3333333
    response.write(Round(i))
    response.write("<br>")
    response.write(Round(j))
  %>

  <br/><hr/><br/>

  <h2> A random number </h2>
  <code>
    <pre>
      ' 난수 생성
      randomize()
      r = rnd()
      response.write("A random number: " & r )
      ' 0 ~ 100 사이의 난수
      randomNumber = Int( 100 * r )
      response.write("A random number: " & randomNumber )
    </pre>
  </code>

  <%
    ' 난수 생성
    randomize()
    r = rnd()
    response.write("A random number: <b>" & r & "</b><br/>")

    ' 0 ~ 100 사이의 난수
    randomNumber = Int( 100 * r )
    response.write("A random number: <b>" & randomNumber & "</b><br/>")
  %>

  <br/><hr/><br/>

  <h2> Return a specified number of characters from left/right of a string </h2>
  <code>
    <pre>
    txt = "Helcome to this WEB"
    response.write( txt )
    response.write( left( txt, 5 ) )
    response.write( right( txt, 5 ) )
    </pre>
  </code>

  <%
    txt = "Helcome to this WEB"
    response.write( txt )
    response.write( "<br/>" )
    response.write( left( txt, 5 ) )
    response.write( "<br/>" )
    response.write( right( txt, 5 ) )
  %>

  <br/><hr/><br/>

  <h2> Replace some characters in a string </h2>
  <code>
    <pre>
      response.write( replace( txt, "WEB", "HELL!!" ) )
    </pre>
  </code>

  <%
    response.write( replace( txt, "WEB", "HELL!!" ) )
  %>

  <br/><hr/><br/>

  <h2> Return a specified number of characters from a string </h2>
  <code>
    <pre>
      response.write( mid( txt, 9, 2 ) )
    </pre>
  </code>

  <%
    response.write( mid( txt, 9, 2 ) )
  %>

</body>
</html>
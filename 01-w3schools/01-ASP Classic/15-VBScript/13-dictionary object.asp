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
  <title>dictionary object</title>
</head>
<body>
  <head>
    There are <%=application( "visitors" )%> on online now!
  </head>

  <br/><hr/><br/>

  <h2> dictionary object </h2>

  <%

    dim d
    set d = server.createObject( "scripting.dictionary" )
    d.add "n", "Norway"
    d.add "i", "Italy"


    if d.exists( "n" ) = true then
      response.write "key exist."
    else
      response.write "key does not exist."
    end if


    response.write "<p>the value of the items are : </p>"
    a = d.items
    for i = 0 to d.count - 1
      s = s & a( i ) & "<br/>"
    next
    response.write s

    s = ""

    response.write "<p>the value of the keys are : </p>"
    a = d.keys
    for i = 0 to d.count - 1
      s = s & a( i ) & "<br/>"
    next
    response.write s

    s = ""

    response.write "<p>the value of the item n : </p>" & d.item( "n" )

    d.key( "i" ) = "it"
    response.write "<p>the key i has been set to it, and value is : </p>" & d.item( "it" )

    response.write "<p>the number of key / item pair is : </p>" & d.count


    set d = nothing

  %>

  <br/><hr/><br/>

</body>
</html>
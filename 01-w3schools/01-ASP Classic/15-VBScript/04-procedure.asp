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

  <h2> Sub procedure </h2>
  <code>
    <pre>
      Sub mysub()
        response.write("I was written by a sub procedure")
      End Sub

      response.write("I was written by the script")
      Call mysub()
    </pre>
  </code>

  <%
    Sub mysub()
      response.write("I was written by a sub procedure")
    End Sub

    response.write("I was written by the script<br>")
    Call mysub()
  %>

  <br/><hr/><br/>

  <h2> function procedure </h2>
  <code>
    <pre>
      Function myfunction()
        myfunction = Date()
      End Function

      response.write("Today's date: ")
      response.write(myfunction())
    </pre>
  </code>

  <%
    Function myfunction()
      myfunction = Date()
    End Function

    response.write("Today's date: ")
    response.write(myfunction())
  %>
  <br/>
  Today's date : <%=myfunction()%>

</body>
</html>
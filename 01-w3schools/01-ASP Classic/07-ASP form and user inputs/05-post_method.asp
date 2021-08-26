<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Document</title>
</head>
<body>
  <form method="post" action="05-post_method.asp">
    <label for="fname">First Name<label> : 
    <input
      type="text" name="fname" id="fname"
      value='<%=request.form("fname")%>'
    />
    <br/>
    <label for="lname">Last Name<label> : 
    <input
      type="text" name="lname" id="lname"
      value='<%=request.form("lname")%>'
    />
    <br/>
    <input type="submit" value="Submit" />
  </form>
  
  <br/>
  Welcome
  <%
    response.write(request.form("fname"))
    response.write(" " & request.form("lname"))
  %>
</body>
</html>
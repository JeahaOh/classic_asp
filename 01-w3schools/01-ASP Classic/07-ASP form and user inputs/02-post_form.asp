<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Document</title>
</head>
<body>
  <form action="02-post_form.asp" method="POST">
    <label for="string">string : </label>
    <input
      type="text" name="string" id="string" size="20"
      value='<%=request.form("string")%>'
    />
    <input type="submit" value="submit" />
  </form>
  <%
    dim string
    string = request.form("string")

    if string <> "" then
      response.write( "string is " & string & "!" )
    end if
  %>
</body>
</html>
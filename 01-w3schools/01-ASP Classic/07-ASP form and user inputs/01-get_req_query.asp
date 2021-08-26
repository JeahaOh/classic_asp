<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Document</title>
</head>
<body>
  <form action="01-get_req_query.asp" method="GET">
    <label for="string">string : </label>
    <input
      type="text" name="string" id="string" size="20"
      value='<%=request.querystring("string")%>'
    />
    <input type="submit" value="submit" />
  </form>
  <%
    dim string
    string = request.querystring("string")

    if string <> "" then
      response.write( "string is " & string & "!" )
    end if
  %>
</body>
</html>
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Document</title>
</head>
<body>
  <%
    dim cars
    cars = request.form("cars")
  %>
  <form action="03-post_radio_form.asp" method="POST">
    <p>
      SELECT ANY CAR : 
    </p>

    <label for="volvo">volvo</label>
    <input
      type="radio" name="cars" id="volvo"
      value="volvo"
      <%
        if cars = "volvo" then response.write( "checked" )
      %>
    />

    <label for="saab">saab</label>
    <input
      type="radio" name="cars" id="saab"
      value="saab"
      <%
        if cars = "saab" then response.write( "checked" )
      %>
    />

    <label for="BMW">BMW</label>
    <input
      type="radio" name="cars" id="BMW"
      value="BMW"
      <%
        if cars = "BMW" then response.write( "checked" )
      %>
    />

    <br/><br/>

    <input type="submit" value="submit" />
  </form>
  <%
    if cars <> "" then
      response.write( "<p>SELECTED VALUE IS " & cars & "</p>" )
    end if
  %>
</body>
</html>
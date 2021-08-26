<%@ language="javascript" %>
<!DOCTYPE html>
<html>
<body>
  <%
    var d = new Date();
    var h = d.getHours();

    Response.Write( "<p>" + d + "</p>" );

    if ( h < 12 ) {
      Response.Write("Good Morning!");
    } else {
      Response.Write("Good Day!");
    }
  %>
</body>
</html>
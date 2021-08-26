<%@ language="javascript" %>
<!DOCTYPE html>
<html>
<head>
  <%
    function jsproc( x, y ) {
      Response.Write( "x * y = " + x * y );
    }
  %>
</head>
<body>
  <p>Result : <% jsproc( 3, 4 ); %></p>
</body>
</html>
<!DOCTYPE html>
<html>
<head>
  <%
    sub vbproc( x, y )
      response.write( x * y )
    end sub
  %>
  <script language="javascript" runat="server">
    function jsproc( x, y ) {
      Response.Write( x * y );
    }
  </script>
</head>
<body>
  <p> call vbproc : <%call vbproc( 3, 4 )%></p>
  <p> call jsproc : <%call jsproc( 3, 4 )%></p>
</body>
</html>
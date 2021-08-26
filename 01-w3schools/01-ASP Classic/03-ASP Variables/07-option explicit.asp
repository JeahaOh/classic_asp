<% Option Explicit %>
<!DOCTYPE html>
<html>
<body>
  <%
    dim x
    dim carname

    carname = "Porsche"

    response.write( "carname : " & carname & "<br/>" )

    'carnime = "carnime"
    'response.write( "carnime : " & carnime & "<br/>" )
    
    
    public a
    a = "public"
    response.write( "a : " & a & "<br/>" )
    private c
    c = "private"
    response.write( "c : " & c & "<br/>" )
  %>
</body>
</html>
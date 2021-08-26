<!DOCTYPE html>
<html>
<body>
  <%
    sub mysub() 
      response.write( "written by a sub procedure.<br/>" )
    end sub

    response.write( "writtten by the script.<br/>" )
    call mysub()
    call mysub
    mysub()
    mysub
  %>
</body>
</html>
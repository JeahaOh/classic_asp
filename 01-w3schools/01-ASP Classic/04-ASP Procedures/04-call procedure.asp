<!DOCTYPE html>
<html>
<body>
  <%
    ' 인수 a, b를 받아서 그 합을 반환하는 func
    function func( a, b )
      func = a + b
    end function

    response.write( func( 5, 9 ) )
  %>
</body>
</html>
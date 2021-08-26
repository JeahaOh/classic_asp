<!DOCTYPE html>
<html>
<body>
  <%
    dim arr(2, 2)
    
    arr( 0, 0 ) = "Volvo"
    arr( 0, 1 ) = "BMW"
    arr( 0, 2 ) = "Ford"
    
    arr( 1, 0 ) = "Apple"
    arr( 1, 1 ) = "Orange"
    arr( 1, 2 ) = "Banana"
    
    arr( 2, 0 ) = "Coke"
    arr( 2, 1 ) = "Pepsi"
    arr( 2, 2 ) = "Sprite"

    for i = 0 to 2
      response.write( "<p>" )
        for j = 0 to 2
          response.write( arr(i, j) & "<br />" )
        next
      response.write( "</br>" )
    next
  %>
</body>
</html>
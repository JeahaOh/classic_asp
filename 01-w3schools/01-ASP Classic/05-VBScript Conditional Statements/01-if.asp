<!DOCTYPE html>
<html>
<body>
  <%
    i = hour( time )
    response.write( "i : " & i & "<br/>" )

    if i = 8 then
      response.write( "Just started..!" )
    elseif i = 11 then
      response.write( "Hungry" )
    elseif i = 12 then
      response.write( "Lunch-time!" )
    elseif i = 17 then
      response.write( "Time to go home!" )
    else
      response.write( "Holy SHIT!" )
    end if
  %>
</body>
</html>
<% @CODEPAGE="65001" language="vbscript" %>
<% Response.CharSet = "utf-8" %>
<!DOCTYPE html>
<html>
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
</head>
<body>
  <%
    d = weekday( data )

    response.write( "weekday( data ) : " & d & "<br/>" )

    select case d
      case 1
        response.write( "SUNDAY 아이스크림 먹고싶다" )
      case 2
        response.write( "MONDAY 출근 안 했으면 좋겠다" )
      case 3
        response.write( "TUESDAY 죽을 것 같다" )
      case 4
        response.write( "WEDNESDAY 점심이 길어서 그나마 참는다" )
      case 5
        response.write( "THURSDAY 피곤해 죽을 것 같다" )
      case 6
        response.write( "FRIDAY 주말 생각해서 참는다" )
      case else
        response.write( "SATURDAY 방전이다" )
    end select
  %>
</body>
</html>
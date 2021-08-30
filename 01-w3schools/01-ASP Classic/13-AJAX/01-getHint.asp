<% @CODEPAGE="65001" language="vbscript" %>
<% Response.buffer = true %>
<%
  response.expires = -1

  ' array 선언
  dim arr(30)
  arr(1) = "Anna"
  arr(2) = "Brittany"
  arr(3) = "Dianna"
  arr(4) = "Eva"
  arr(5) = "Fiona"
  arr(6) = "Gunda"
  arr(7) = "Hege"
  arr(8) = "Inga"
  arr(9) = "Johanna"
  arr(10) = "Kitty"
  arr(11) = "Linda"
  arr(12) = "Nina"
  arr(13) = "Ophelia"
  arr(14) = "Petunia"
  arr(15) = "Amanda"
  arr(16) = "Raquel"
  arr(17) = "Cindy"
  arr(18) = "Doris"
  arr(19) = "Eve"
  arr(20) = "Evita"
  arr(21) = "Sunniva"
  arr(22) = "Tove"
  arr(23) = "Unni"
  arr(24) = "Violet"
  arr(25) = "Liza"
  arr(26) = "Elizabeth"
  arr(27) = "Ellen"
  arr(28) = "Wenche"
  arr(29) = "Vicky"
  arr(30) = "cinderella"
  

  ' QueryString에서 q의 값 가져오기
  q = ucase( request.querystring( "q" ) )

  if len( q ) > 0 then
    hint = ""
    for i = 0 to 30
      if q = ucase( mid( arr( i ), 1, len( q ) ) ) then
        if hint = "" then
          hint = arr( i )
        else
          hint = hint & ", " & arr( i )
        end if
      end if
    next
  end if

  ' hint가 없다면 "No Suggestion"
  ' 있다면 hint
  if hint = "" then 
    response.write( "No Suggestion" )
  else
    response.write( hint )
  end if

%>
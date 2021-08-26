<%
  dim numvisits
  ' response.cookies로 쿠키 생성,
  ' **response.cookies 명령은 <html> 태그보다 앞에 있어야 함
  ' 방문 1일 뒤 쿠키 만료
  response.cookies( "NumVisits" ).expires = date + 1
  ' 쿠키가 만료될 날짜 설정도 가능함
  ' response.cookies( "NumVisits" ).expires = #Dec 10, 2021#

  ' 쿠키의 값 검색
  numvisits = request.cookies( "NumVisits" )

  if numvisits = "" then

    response.cookies( "NumVisits" ) = 1
    response.write( "First Time!" )
  else
    response.cookies( "NumVisits" ) = numvisits + 1
    response.write( numvisits )

    if numvisits = 1 then
      response.write " Time before!"
    else
      response.write " Times BEFORE!"
    end if
  end if
%>
<!DOCTYPE html>
<html>
<body>
</body>
</html>
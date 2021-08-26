# VBScript 조건문

## 조건문

VBScript에는 4개의 조건문이 있음.
- if 문 
- if then else 문
- if then elseif 문
- select case 문

... 생략

## if .. the .. elseif

```asp
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
```

## select case

switch 문과 같음.

```asp
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
```

---

- 출처 : [https://www.w3schools.com/asp/asp_conditionals.asp](https://www.w3schools.com/asp/asp_conditionals.asp)
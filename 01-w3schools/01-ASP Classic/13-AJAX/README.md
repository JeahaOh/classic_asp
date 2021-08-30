# ASP AJAX

AJAX는 전체 페이지를 다시 로드하지 않고 웹 페이지의 일부를 업데이트 하는 것임.  
  
AJAX = Asynchronouse JavaScript ans XML의 약자로 동적인 웹페이지를 만드는 기술임.  
AJAX는 인터넷 표준을 기반으로 다음을 조합하여 사용함.
- XMLHttpRequest 객체 : 서버와 비동기적으로 데이터 교환
- JavaScript/DOM : 정보 표시 / 상호작용
- CSS
- XML : 종종 데이터 전송 형식으로 사용됨.
  
나머지 기초적인 설명은 패스하도록 함.

## Example 1

```html
  <% @CODEPAGE="65001" language="vbscript" %>
  <% session.CodePage = "65001" %>
  <% Response.CharSet = "utf-8" %>
  <% Response.buffer=true %>
  <% Response.Expires = 0 %>
  <!DOCTYPE html>
  <html lang="ko">
  <head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
      * { background-color : black; color : white; margin : 1em; }
    </style>
    <title>The XMLHttpRequest Object</title>
  </head>
  <body>
    <head>
      There are <%=application( "visitors" )%> on online now!
    </head>
    <h1>Start typing a name in the input field below</h1>
    <br/><hr/><br/>

    <form action="">
      <label for="txt">INPUT</label>
      <input
        type="text" name="txt" id="txt"
        onkeyup="showHint(this.value)"
      />
    </form>

    <p>Suggestions : <span id="hint"></span></p>

    <script>
      const showHint = ( str ) => {
        let xhttp;
        
        if( str.length == 0 ) {
          document.getElementById( 'hint' ).innerHTML = '';
          return;
        }

        xhttp = new XMLHttpRequest();
        xhttp.onreadystatechange = function () {
          if( this.readyState == 4 && this.status == 200 ) {
            document.getElementById( 'hint' ).innerHTML = this.responseText;
          }
        };

        xhttp.open( 'GET', '01-getHint.asp?q=' + str, true);
        xhttp.send();
      }
    </script>

  </body>
  </html>
```
```asp
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
```
  

## Example 2 : AJAX 데이터 베이스 연결

DB 연결은 당장 여건이 안되므로 스킵 해둔다.


---

- 출처 : [https://www.w3schools.com/asp/asp_ajax.asp](https://www.w3schools.com/asp/asp_ajax.asp)  
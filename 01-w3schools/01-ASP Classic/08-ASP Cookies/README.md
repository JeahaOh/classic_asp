# ASP 쿠키

쿠키는 사용자 식별에 사용됨.  
동일한 브라우져로 페이지를 요청할 때마다 쿠키도 보냄.  
ASP로 쿠키 값을 만들 수도 검색할 수도 있음.  
  
환영 쿠키
  
```asp
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
```
  
---
  
## A Cookie with Keys

쿠키에 여러 값의 컬렉션이 포함되어 있다면 쿠키에 키가 있다고 함  
user라는 쿠키 컬렉션을 만들고, 사용자에 대한 정보가 포함된 키와 값을 할당함  
  
```asp
  <%
    Response.Cookies("user")("firstname")="SorR"
    Response.Cookies("user")("lastname")="Ki"
    Response.Cookies("user")("country")="Nowhere"
    Response.Cookies("user")("age")="25"
  %>
```
  
서버가 위의 모든 쿠키를 사용자단에 보냈다고 가정하면,  
사용자에게 전송된 쿠키를 읽을 수 있어야 함.  
  
```asp
  <%
    dim x, y
    for each x in request.cookies
      response.write "<p>" 
      if request.cookies( x ).HasKeys then
        for each y in request.cookies( x )
          response.write( x & " : " & y & " = " & request.cookies( x )( y ) )
          response.write( "<br/>" )
        next
      else
        response.write( x & " = " & request.cookies( x ) & "<br/>" )
      end if
      response.write "</p>"
    next
  %>
```
  
출력 결과  
user : age = 25
user : country = Nowhere
user : lastname = Ki
user : firstname = SorR
  
---
  
## 브라우져에서 쿠키를 지원하지 않는다면

어플리케이션이 쿠키를 지원하지 않는 브라우져를 다루는 경우 어플리케이션의 한 페이지에서 다른 페이지로 정보를 전달하기 위해 다음 두 가지 방법이 있음.

### 1. URL에 매개변수 추가

URL에 매개변수를 추가할 수 있음.  
`<a href="welcome.asp?fname=John&lname=Smith">Go to Welcome Page</a>`  
이후 다음 페이지에서 값을 읽을 수 있음.  
  
```asp
  <%
    fname=Request.querystring("fname")
    lname=Request.querystring("lname")
    response.write("<p>Hello " & fname & " " & lname & "!</p>")
    response.write("<p>Welcome to my Web site!</p>")
  %>
```
  

### 2. FORM 사용

form을 사용 사용자가 요청을 하면 form은 사용자의 입력 값을 다음 페이지로 전달함.  
  
```asp
  <form method="post" action="welcome.asp">
    First Name: <input type="text" name="fname" value="">
    Last Name: <input type="text" name="lname" value="">

    <input type="submit" value="Submit">
  </form>
```
  
이후 다음 페이지에서 값을 읽을 수 있음.
  
```asp
  <%
    fname=Request.form("fname")
    lname=Request.form("lname")
    response.write("<p>Hello " & fname & " " & lname & "!</p>")
    response.write("<p>Welcome to my Web site!</p>")
  %>
```
  
  
---
  
- 출처 : [https://www.w3schools.com/asp/asp_cookies.asp](https://www.w3schools.com/asp/asp_cookies.asp)  
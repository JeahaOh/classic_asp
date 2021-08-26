# ASP form과 사용자 입력
  
Request.QueryString 및 Request.Form 명령은 form 에서 사용자 입력을 검색하는데 사용함.
  

## GET 방식
  
Request.QueryString 명령어를 사용하여 사용자의 입력을 받음.
  
```asp
  <body>
    <form action="req_query.asp" method="GET">
      String : <input type="text" name="string" size="20" value='<%=request.querystring("string")%>' />
      <input type="submit" value="submit" />
    </form>
    <%
      dim string
      string = request.querystring("string")

      if string <> "" then
        response.write( "string is " & string & "!" )
      end if
    %>
  </body>
```
  

## POST 방식
  
Request.Form 명령어를 사용하여 사용자의 입력을 받음.
  
```asp
  <body>
    <form action="02-post_form.asp" method="POST">
      String : <input type="text" name="string" size="20" value='<%=request.form("string")%>' />
      <input type="submit" value="submit" />
    </form>
    <%
      dim string
      string = request.form("string")

      if string <> "" then
        response.write( "string is " & string & "!" )
      end if
    %>
  </body>
```
  

## input:type=[radio]
  
Request.Form 명령으로 라디오 버튼을 통해 입력을 받음.
  
```asp
  <body>
    <%
      dim cars
      cars = request.form("cars")
    %>
    <form action="03-post_radio_form.asp" method="POST">
      <p>
        SELECT ANY CAR : 
      </p>

      <label for="volvo">volvo</label>
      <input
        type="radio" name="cars" id="volvo"
        value="volvo"
        <%
          if cars = "volvo" then response.write( "checked" )
        %>
      />

      <label for="saab">saab</label>
      <input
        type="radio" name="cars" id="saab"
        value="saab"
        <%
          if cars = "saab" then response.write( "checked" )
        %>
      />

      <label for="BMW">BMW</label>
      <input
        type="radio" name="cars" id="BMW"
        value="BMW"
        <%
          if cars = "BMW" then response.write( "checked" )
        %>
      />

      <br/><br/>

      <input type="submit" value="submit" />
    </form>
    <%
      if cars <> "" then
        response.write( "<p>SELECTED VALUE IS " & cars & "</p>" )
      end if
    %>
  </body>
```
  

---
  
## 사용자 입력

request 객체틑 form에서 정보를 검색하는 데 사용할 수 있음.  
사용자 입력은 request.qeurtystring 또는 request.form 명령으로 검색 가능함.  
  
---
  
## request.querystring

**request.querystring**는 GET 메소드 방식의 request에서 값을 수집하는데 사용함.  
  
```asp
  <body>
    <form method="get" action="04-get.asp">
      <label for="fname">First Name<label> : 
      <input
        type="text" name="fname" id="fname"
        value='<%=request.querystring("fname")%>'
      />
      <br/>
      <label for="lname">Last Name<label> : 
      <input
        type="text" name="lname" id="lname"
        value='<%=request.querystring("lname")%>'
      />
      <br/>
      <input type="submit" value="Submit" />
    </form>
    
    <br/>
    Welcome
    <%
      response.write(request.querystring("fname"))
      response.write(" " & request.querystring("lname"))
    %>
  </body>
```

## request.form

**request.form**은 POST 메소드 방식의 request에서 값을 수집하는데 사용함.  
  
```asp
  <body>
    <form method="post" action="05-post_method.asp">
      <label for="fname">First Name<label> : 
      <input
        type="text" name="fname" id="fname"
        value='<%=request.form("fname")%>'
      />
      <br/>
      <label for="lname">Last Name<label> : 
      <input
        type="text" name="lname" id="lname"
        value='<%=request.form("lname")%>'
      />
      <br/>
      <input type="submit" value="Submit" />
    </form>
    
    <br/>
    Welcome
    <%
      response.write(request.form("fname"))
      response.write(" " & request.form("lname"))
    %>
  </body>
```
  
---
  
## form 유효성 검사
  
사용자 단에서 유효성 검사를하는 것이 빠르고 서버 부하를 줄임.  
사용자 입력이 DB에 입력 될 경우 서버단에서 유효성 검사를 해야함.  
서버에서 유효성 검사를 하는 제일 좋은 방법은 다른 페이지로 이동하는 대신 양식 자체에 POST 하는 방법임.  
이렇게 하면 사용자는 동일한 페이지에서 오류를 확인할 수 있고, 오류를 더 쉽게 찾을 수있음.
  

---
  
- 출처 : [https://www.w3schools.com/asp/asp_inputforms.asp](https://www.w3schools.com/asp/asp_inputforms.asp)  
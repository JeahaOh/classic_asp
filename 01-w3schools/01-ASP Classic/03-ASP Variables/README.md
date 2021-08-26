# ASP 변수

변수의 기본 개념은 다르지 않음.

---

## 예제

### 변수 선언

변수를 선언하고 변수에 값을 할당하고 텍스트에서 값을 출력하는 방법

```asp
  <!DOCTYPE html>
  <html>
  <body>
    <%
      dim name
      name = "ASDF Stupid"
      response.write("My name is: " & name)
    %>
  </body>
  </html>
```

### 반복문

```asp
  <% @CODEPAGE="65001" language="vbscript" %>
  <!DOCTYPE html>
  <head>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
  </head>
  <html>
  <body>
    <%
      dim i
      for i = 1 to 6
        response.write( "<h" & i & ">Heading " & i & "</h" & i & ">" )
      next
    %>
  </body>
  </html>
```

### 배열 생성

```asp
  <!DOCTYPE html>
  <html>
  <body>
    <%
      dim famname(5), i
      famname(0) = "Tov Lo"
      famname(1) = "Flume"
      famname(2) = "Heize"
      famname(3) = "쓰레기 같은"
      famname(4) = "ASP"
      famname(5) = "Borge"

      for i = 0 to 5
        response.write(famname(i) & "<br>")
      next
    %>
  </body>
  </html>
```

### VBScript를 이용한 시간 출력

서버의 시간 출력

```asp
  <% @CODEPAGE="65001" language="vbscript" %>
  <!DOCTYPE html>
  <head>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
  </head>
  <html>
  <body>
    <%
      dim h
      h = hour( now() )

      response.write( "<p>" & now() & "</p>" )

      if h < 12 then
        response.write( "Good Morning!" )
      else
        response.write( "Good Day!" )
      end if
    %>
  </body>
  </html>
```

### JavaScript를 이용한 시간 출력

위의 예제와 같지만 JS를 이용함.  
W3C예제를 따라한건데, javascript를 사용하는건지 JScript를 사용하는건지 솔직히 분간은 안됨.

```asp
  <%@ language="javascript" %>
  <!DOCTYPE html>
  <html>
  <body>
    <%
      var d = new Date();
      var h = d.getHours();

      Response.Write( "<p>" + d + "</p>" );

      if ( h < 12 ) {
        Response.Write("Good Morning!");
      } else {
        Response.Write("Good Day!");
      }
    %>
  </body>
  </html>
```

### 변수 생성 및 변경

```asp
  <!DOCTYPE html>
  <html>
  <body>
  <%
    dim asdf
    asdf = "Flume"

    response.write(asdf)

    response.write("<br>")

    asdf = "Tove Lo"
    response.write(asdf)
  %>
  </body>
  </html>
```
  
---
  

## 대수학

`x = 5, y = 6, z = x + y`  
x와 y에 값을 저장할 수 있고, 이 정보를 이용하여 z의 값 11을 구할 수 있음.  
이런 x, y, z를 변수라고 함.

## VBScript 변수

VBScript의 변수는 값이나 표현식을 저장하는 데 사용됨.  
변수는 x와 같은 짧은 이름이나 carname과 같이 더 명시적인 이름을 사용할 수 있음.  

- VBScript의 변수 작명 규칙
  - 문자로 시작해야 함
  - **.**를 포함할 수 없음
  - 255자를 넘길 수 없음
  
VBScript의 모든 변수는 타입이 지정되어 있지 않은 variant 유형임

## VBScript 변수 선언

VBScript에서 변수를 생성하는 것은 선언 변수라고 함.  
dim, public, private 키워드를 이용해서 VBScript 변수를 선언할 수 있음.

```asp
  dim x
  dim carname
```

두개의 변수 x와 carname을 만들었음.  
스크립트에서 이름을 사용해서 변수를 선언 할 수 있음.  
```asp
  carname = "Porsche"
```

변수에 값도 할당했음. 변수 이름은 **carname**, 값은 **Porsche** 임. 스크립트에서 변수 이름의 철자를 잘못 입력 할 수 있고, 코드가 실행될 때 의도하지 않은 결과를 초래할 수 있음.  
**carname** 변수를 **carnime**으로 잘못 입력하면 스크립트는 **carnime**라는 새 변수를 자동으로 생성함.  
스크립트가 이런 의도치 않은 변수 생성을 막으려면 Option Explicit 옵션을 넣을 수 있음.  
스크립트 맨 위에 Option Explicit을 넣으면 됨.

```asp
  <%@ Language=VBScript %>
  <% Option Explicit %>
  <!DOCTYPE html>
  <html>
  <body>
    <%
      dim x
      dim carname

      carname = "Porsche"

      response.write( "carname : " & carname & "<br/>" )

      ' Option Explicit 설정이 있어서 아래 구문은 에러가 생김
      carnime = "carnime"
      
      
      public a
      a = "public"
      response.write( "a : " & a & "<br/>" )
      private c
      c = "private"
      response.write( "c : " & c & "<br/>" )
    %>
  </body>
  </html>
```
```log
  carname : Porsche
  Microsoft VBScript 런타임 오류 오류 '800a01f4'

  변수가 정의되지 않았습니다.: 'carnime'

  /test/test/example.asp, 줄 15
```
  
- public, private 차이는 알 수 없네
  
  
---
  
## VBScript 배열 변수

배열 변수는 단일 변수에 여러 값을 저장하는데 사용함.  

```asp
  <!DOCTYPE html>
  <html>
  <body>
    <%
      ' 2차원 배열 선언
      dim arr(2, 2)
      
      ' 할당
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
            ' 배열에서 인덱스로 값 출력
            response.write( arr(i, j) & "<br />" )
          next
        response.write( "</br>" )
      next
    %>
  </body>
  </html>
```

## 변수의 수명

프로시져 외부에서 선언된 변수는 ASP 파일의 모든 스크립트에서 엑세스하고 변경 가능함.  
  
프로시져 내부에서 선언된 변수는 프로시져가 실행될 때마다 생성되고 소멸함. 프로시져 외부의 어떤 스크립트도 변수에 엑세스하거나 변수를 변경할 수 없음.  
  
둘 이상의 ASP 파일에 엑세스할 수 있는 변수를 선언하려면 **세션 변수** 또는 **어플리케이션 변수**로 선언해야 함.

### 세션변수

세션 변수는 단일 사용자에 대한 정보를 저장하는 데 사용되며 하나의 어플리케이션의 모든 페이지에서 사용할 수 있음.  
일반적으로 세션 변수에 저장되는 정보는 이름, ID 및 기본 설정임.

### 어플리케이션 변수

어플리케이션 변수는 한 어플리케이션의 모든 페이지에서 사용 가능.  
어플리케이션 변수는 하나의 특정 어플리케이션에 있는 모든 사용자에 대한 정보를 저장하는 데 사용함.

---

출처 : [https://www.w3schools.com/asp/asp_variables.asp](https://www.w3schools.com/asp/asp_variables.asp)
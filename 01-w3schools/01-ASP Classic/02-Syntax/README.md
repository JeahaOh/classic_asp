# ASP Syntax

## ASP는 VBScript를 사용함.

**ASP**의 기본 스크립팅 언어는 **VBScript**임.  
스크립팅 언어는 경량 프로그래밍 언어임.  
VBScript는 MS Visual Basic의 라이트 버젼임.  

## ASP 파일

ASP 파일은 일반적인 HTML 파일일 수 있음.  
또한 ASP 파일에는 서버 스크립트도 포함될 수 있음.  
`<% %>`로 둘러 싸인 스크립트는 서버에서 실행됨.  
**Response.Write()** 메소드는 ASP에 의해 HTML로 출력됨.  
다음 코드는 "Hello World"를 출력함.  
```asp
  <!DOCTYPE html>
  <html>
  <head></head>
  <body>
    <%
      response.write("Hello World!")
    %>
  </body>
  </html>
```

VBScript는 대소문자를 구분하지 않음.  
`Response.Write()`는 `response.write()`와 같음.

## ASP에서 JavaScript 사용

JS를 웹페이지의 스크립팅 언어로 설정하려면 페이지 상단에 언어 사양을 삽입해야 함.  
```asp
  <%@ language="javascript" %>
  <!DOCTYPE html>
  <html>
  <head></head>
  <body>
    <%
      Response.Write("Hello World!")
    %>
  </body>
  </html>
```

javascript는 `Response.Write()`의 대소문자를 구분함.  
앞으로의 ASP 튜토리얼은 VBScript를 사용함.  

## 그 외

### EX 03

`Response.Write()`는 `=`로 대치될 수 있음.  
```asp
  <%="reponse.write() can replaced by `=`"%>
```

### EX 04

HTML 태그는 출력 가능함.
```asp
  <%="<h2>Using HTML tags to format the TEXT is available, FUCK!</h2>"%>
```


### EX 04

HTML 태그 속성 역시 출력 가능.
```asp
  <%="<p style='color: #0000FF;'>This text is styled.</p>"%>
```
  

--- 
  
## VBScript 예제와 레퍼런스

- [예제](https://www.w3schools.com/asp/asp_examples.asp)
- [레퍼런스](https://www.w3schools.com/asp/asp_ref_vbscript_functions.asp)

--- 

출처 : [https://www.w3schools.com/asp/asp_syntax.asp](https://www.w3schools.com/asp/asp_syntax.asp)

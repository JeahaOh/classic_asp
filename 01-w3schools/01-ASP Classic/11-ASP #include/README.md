# ASP include file

## \#include 지시문

\#include 지시문을 사용하여 서버가 실행하기 전에 한 ASP 파일의 내용을 다른 ASP 파일에 삽입할 수 있음.  
\#include 지시문은 여러 페이지에서 재사용될 함수, 머리글, 바닥글, 또는 요소를 만드는 데 사용됨.

  ---  

## \#include 지시문 사용하는 방법

```html
  <!DOCTYPE html>
  <html lang="ko">
  <head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
      * { background-color : black; color : white; margin : 1em; }
    </style>
    <title>include my page</title>
  </head>
  <body>
    <h1>include my page</h1>
    <br/><hr/><br/>

    <h3>Word of Wisdom :</h3>
    <p><!--#include file="01-wisdom.inc" --></p>
    
    <br/><hr/><br/>

    <h3>The Time is :</h3>
    <p><!--#include file="01-time.inc" --></p>

  </body>
  </html>
```
  
01-wisdom.inc
  
```txt
"One should never increase, beyond what is necessary,
the number of entities required to explain anything."
```
  
01-time.inc
  
```txt
<% response.write( time ) %>
```

  ---  

## 파일 포함 구문

ASP 페이지에 파일을 포함하려면 주석 태그 안에 #include 지시문을 넣어야 함.

```txt
<!-- #include virtual = "someFileName" -->
of
<!-- #include file = "someFileName" -->
```

### virtual 키워드

virtual 키워드를 사용하여 가상 디렉토리로 시작하는 경로를 나타냄.  
"header.inc"라는 파일이 /html 이라는 가상 디렉토리에 있는 경우 다음 줄은 "header.inc"의 내용을 삽입함.  
`<!-- #include virtual = "/html/header.inc" -->`
  
### file 키워드

file 키워드를 사용하면 상대경로를 나타냄. 산대경로는 포함 파일이 포함된 디렉토리로 시작함.  
html 디렉토리에 파일이 있고, "header.inc"파일이 html/headers에 있는 경우 다음 줄은 파일에 "header.inc"를 삽입함.  
`<!-- #include file = "headers/header.inc" -->`  
  
포함된 파일의 경로 (header/header.inc)는 포함된 파일에 상대적임. 이 #include 문이 포함된 파일이 html 디렉토리에 없으면 명령문은 작동하지 않음.  
  
  ---  

## 기타 참고사항

위 섹션에서는 포함된 파일에 대해서 확장자 **.inc**를 사용했음. 사용자가 inc 파일을 직접 탐색하려고 하면 해당 내용이 표시됨. 포함된 파일에 기밀 정보나 다른 사용자에게 보여주고 싶지 않은 정보가 포함되어 있다면, ASP 확장자를 사용하는 것이 좋음. ASP 파일의 소스 코드는 해석 후에 보이지 않음. 포함된 파일에는 다른 파일명도 포함될 수 있으며 하나의 ASP 파일에는 동일한 파일이 두번 이상 포함될 수 있음.  
  
**중요 :** 포함된 파일은 스크립트가 실행되기 전에 처리되고 삽입됨. 다음 스크립틍는 ASP가 변수 값을 할당하기 전에 #include 지시문을 실행하기 때문에 작동하지 않음.

```asp
  <%
    filename = "header.inc"
  %>
  <!-- #include file = "<%filename%>" -->
```
  
inc 파일에서 스크립트 구분 기호를 열거나 닫을 수 없음. 다음 코드는 작동하지 않음.
  
```asp
  for i = 1 to n
    <!-- #include file = "count.inc" -->
  next
```
```asp
  <% for i = 1 to n %>
    <!-- #include file = "count.inc" -->
  <% next %>
```

---

- 출처 : [https://www.w3schools.com/asp/asp_incfiles.asp](https://www.w3schools.com/asp/asp_incfiles.asp)  
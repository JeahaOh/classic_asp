# ASP Intro

ASP는 동적 웹 페이지를 만들기 위한 오래된 기술임.  
ASP는 PHP 처럼 웹 서버에서 script를 실행하기 위한 기술임.  

## Show Example

- ex :  
  ```asp
    <!DOCTYPE html>
    <html>
    <body>
      <%
        response.write("My First ASP Script!")
      %>
    </body>
    </html>
  ```

## ASP 란

- Active Server Pages의 약자
- MS의 기술
- 웹 서버 내부에서 실행되는 프로그램

## ASP 파일이란

- **.asp** 확장자를 갖는 파일
- HTML 파일과 비슷
- HTML 외에 서버 스크립트를 포함하고 있음.
- ASP 파일의 서버 스크립트는 서버에서 실행 됨.

## ASP가 할 수 있는 것

- 웹 페이지 편집, 변경, 콘텐츠 추가 또는 사용자 지정
- HTML form에서 제출된 사용자 쿼리 또는 데이터에 응답
- DB 또는 기타 서버 데이터에 엑세스하고 결과를 응답
- 브라우져에서 ASP 코드를 볼 수 없으므로 웹 보안 제공
- 빠르고 단순함

## 작동 원리

일반 HTML 파일을 요청하면 서버는 파일을 변환,  
브라우져가 ASP 파일을 요청하면 ASP 파일을 읽고 파일에서 서버 스크립트를 실행하는 ASP 엔진에 요청을 전달.  
ASP파일은 HTML로 응답됨.

## reference and example

W3School에서 ASP의 기본 제공 개체 및 구성요소, 해당 속성 및 메소드에 대한 ASP reference를 제공함.  
ASP 스크립트는 서버에서 실행되기 때문에 브라우져에서 ASP 코드를 볼 수 없음.  
W3School에서 ASP 예제를 살펴볼 수 있음.
- [Reference](https://www.w3schools.com/asp/asp_ref_response.asp)
- [Example](https://www.w3schools.com/asp/asp_examples.asp)
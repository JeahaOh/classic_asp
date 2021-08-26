# ASP에서 UTF-8 처리

다른 언어나 프레임웤보다 쓸데 없는게 많다.

1. 모든 ASP 코드 첫줄에 다음과 같은 코드를 추가함.

```asp
  <% @CODEPAGE="65001" language="vbscript" %>
  <% Option Explicit %>
  <% session.CodePage = "65001" %>
  <% Response.CharSet = "utf-8" %>
  <% Response.buffer=true %>
  <% Response.Expires = 0 %>
```

2. meta tag를 추가함

```html
  <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
```

3. response.charset = "utf-8"
  - asp의 response.charset으로 문자 코드셋 지정
  - <html> 태그보다 앞에 있어야 함

4. 파일 저장시 encoding을 utf-8로 지정

5. DB insert / update시 숫자 타입을 제외하고는 모든 대상에 N을 추가
  ```sql
    Insert 테이블이름 (칼럼a, 칼럼b) value (N'입력a', N'입력b')
    update 테이블이를 set 칼럼a = N'입력a' where 고유칼럼 = '번호'
  ```

6. DB like 검색시 N 추가


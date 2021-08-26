# ASP Procedure

ASP에서는 VBScript와 JavaScript 간 프로시져를 상호 호출 가능함.

## Procedure

다음과 같은 프로시져 기능이 있을 수 있음.


VBScript Procedure
```asp
  <!DOCTYPE html>
  <html>
  <head>
    <%
      sub vbproc( x, y )
        response.write( "x * y = " & x * y )
      end sub
    %>
  </head>
  <body>
    <p>Result : <%call vbproc( 3, 4 )%></p>
  </body>
  </html>
```

다른 스크립팅 언어로 프로시져 / 함수를 작성하려면 `<html>` 태그 위에 `<%@ language="{사용할 언어}" %>`를 넣어주면 됨.  
  
js procedure
```asp
  <%@ language="javascript" %>
  <!DOCTYPE html>
  <html>
  <head>
    <%
      function jsproc( x, y ) {
        Response.Write( "x * y = " + x * y );
      }
    %>
  </head>
  <body>
    <p>Result : <% jsproc( 3, 4 ); %></p>
  </body>
  </html>
```
  
---
  
## VBScript와 JS의 차이점

VBScript로 작성된 ASP 파일에서 VBScript 또는 JS 프로시져를 호출할 때 프로시져 이름 뒤에 call 키워드를 사용할 수 있음.  
프로시져에 매개변수가 필요한 경우 **call** 키워드를 사용할 때 매개변수 목록을 괄호로 묶어야 함. **call** 키워드를 생략하는 경우 매개변수 목록을 괄호로 묶지 않아야 함.  
프로시져에 매개변수가 없으면 괄호는 선택사항임.  
  
JS로 작성된 ASP 파일에서 JS 또는 VBScript 프로시져를 호출할 때 항상 프로시져 이름 뒤에 괄호를 사용해야 함.  
  
---
  
## VBScript Procedure

VBScript에는 두가지 종류의 프로시져가 있음.
- Sub Procedure
- Function Procedure

### VBScript Sub Procedure

- **sub**, **end sub** 으로 묶인 일련의 문장임.
- 작업을 하지만 값을 반환하지 않음.
- 인수를 취할 수 있음.

```asp
  sub mysub() 
    some statements
  end sub
```
```asp
  sub mysub( arg1, arg2 ) 
    some statements
  end sub
```
```asp
  <!DOCTYPE html>
  <html>
  <body>
    <%
      sub mysub() 
        response.write( "written by a sub procedure.<br/>" )
      end sub

      response.write( "writtten by the script.<br/>" )
      call mysub()
      call mysub
      mysub()
      mysub
    %>
  </body>
  </html>
```

### VBScript Function Procedure

- **function**, **end function**으로 묶인 일련의 문장.
- 작업을 하고 값을 반환함.
- 이름에 값을 할당하여 값을 반환함.
- 호출 프로시져에 의새 인수를 취할 수 있음.
- 인수가 없으면 빈 괄호 **()**를 포함 해야함.

```javascript
  function func()
    some statements
    func = some value
  end function
```
```javascript
  function func( arg1, arg2 )
    some statements
    func = some value
  end function
```

### 프로시져 호출

```asp
  <!DOCTYPE html>
  <html>
  <body>
    <%
      ' 인수 a, b를 받아서 그 합을 반환하는 func
      function func( a, b )
        func = a + b
      end function

      response.write( func( 5, 9 ) )
    %>
  </body>
  </html>
```

### VBScript를 이용하여 프로시져 호출

ASP에서 javascript 와 VBScript 에서 만든 프로시져를 모두 호출하는 방법

```asp
  <!DOCTYPE html>
  <html>
  <head>
    <%
      sub vbproc( x, y )
        response.write( x * y )
      end sub
    %>
    <script language="javascript" runat="server">
      function jsproc( x, y ) {
        Response.Write( x * y );
      }
    </script>
  </head>
  <body>
    <p> call vbproc : <%call vbproc( 3, 4 )%></p>
    <p> call jsproc : <%call jsproc( 3, 4 )%></p>
  </body>
  </html>
```

---

- 출처 : [https://www.w3schools.com/asp/asp_procedures.asp](https://www.w3schools.com/asp/asp_procedures.asp)  
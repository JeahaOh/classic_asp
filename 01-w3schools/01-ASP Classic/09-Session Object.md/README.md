# ASP 세션 객체

컴퓨터에서 어플리케이션으로 작업할 때 어플리케이션을 열고 일부 변경한 다음 닫음.  
이건 세션와 유사한데, 컴푸터는 작업자가 누군지 알고있고, 어플리케이션을 열 때와 닫을 때를 알고 있음.  
그러나 인터넷에는 한 가지 문제점이 있는데, HTTP 주소는 상태 유지를 지원하지 않기 때문에 웹서버는 당신이 누구이고, 무엇을 하는지 모름.  
  
ASP는 각 사용자에 대해 고유 쿠키를 만들어 이 문제를 해결함. 쿠키는 컴퓨터로 전송되며 사용자를 식별하는 정보가 포함됨.  
이 인터페이스를 Session 객체라고 함.  
  
Session 객체는 사용자 세션에 대한 정보를 저장하거나 설정을 변경함.  
  
Sessio 객체에 저장된 변수는 단일 사용자에 대한 정보를 보유하며, 하나의 어플리케이션의 모든 페이지에서 사용할 수 있음. 세션 변수에 저장되는 공통 정보는 이름, ID 및 기본 설정등임. 서버는 각각의 새 사용자에 대해 새 Session 객체를 만들고, 세션이 만료되면 session 객체를 삭제함.  
  
---

## 세션의 시작

- 새 사용자가 ASP 파일을 요청하고 Global.asa 파일에는 Session_OnStart 프로시져가 포함됨.
- 값은 세션 변수에 저장됨.
- 사용자가 ASP 파일을 요청하고 Global.asa 파일은 `<object>` 태그를 사용하여 세션 범위의 개체를 인스턴스화 함.

## 세션의 종료

사용자가 지정된 기간 (기본 시간은 20분) 동안 어플리케이션에서 페이지를 요청하거나 새로 고치치 않으면 세션은 종료됨.  
  
기본보다 짧거나 긴 시간 초과 간격을 사용하려면 **Timeout** 속성을 사용해야 함.  
```javascript
  //  세션 유지 시간을 5분으로 설정함.
  Session.Timeout = 5
```
  
**Abandon** 메소드를 사용하여 세션을 즉시 종료함.
```javascript
  Session.Abondon
```
  
세션의 주요 문제는 세션을 종료해야하는 시간임. 사용자의 마지막 요청이 마지막 요청인지 아닌지 알 수 없음. 따라서 세션을 "활성" 상태로 유지해야하는 기간을 알 수 없음. 유휴 세션을 너무 오래 기다리면 서버의 리소스가 소모되지만 세션이 너무 빨리 삭제되면 서버가 모든 정보를 삭제했기 때문에 사용자가 처음부터 다시 시작해야 할 수 있음. 적절한 세션초과 시간을 찾는 것은 어려움.  
세션 변수에 데이터를 너무 많이 저장해 둬도 좋지 않음.
  
---
  
## 세션 변수 저장 및 검색
  
세션 객체에서 가장 중요한 것은 변수를 저장할수 있다는 것임.  

```asp
  <%
    session( "username" ) = "borry"
    seesion( "age" ) = 3
  %>
```
  
값이 세션 변수에 저장되면 ASP 어플리케이션의 모든 페이지에서 이 값에 도달할 수 있음.

`Welcome <%response.write( session( "username" ) ) %>`
  
세션 객체에 사용자 기본 설정을 저장한 다음 해당 기본 설정에서 엑세스하여 사용자에게 반환할 페이지를 선택할 수도 있음.

```asp
  <%IF session( "screenres" ) = "low" THEN%>
    This is the text version of the page
  <%ELSE%>
    This is the multimedia version of the page
  <%END IF%>
```

## 세션 변수 제거

Contents 컬렉션에는 모든 세션 변수가 포함됨.  
remove 메소드를 사용하면 세션 변수를 제거할 수 있음.  
"age"값이 18보다 작은 경우 세션 변수 "sale"을 제거함

```asp
  <%
    session.( "sale" ) = true

    if session.contents( "age" ) < 18 then
      session.contents.remove( "sale" )
    end if
  %>
```
  
세션의 모든 변수를 제거하려면 removeAll 메소드를 사용함.

```asp
  session.contents.removeAll()
```
  
---
  

## 컨텐츠 컬렉션을 통해 반복문
  
contents collection에 모든 세션 변수가 포함됨. contents collection을 반복하여 컬렉션에 무엇이 저장되어 있는지 확인 가능함.
  
```asp
  session( "username" ) = "borry"
  session( "age" ) = 3

  dim i
  for each i in session.contents
    response.write( i & "<br/>" )
  next
```

contents collection의 항목 수를 모르는 경우 count 속성을 사용할 수 있음.

```asp
  <%
    dim i, j
    j = session.contents.count
    response.write( "session variables : " & j )

    for i = 1, to j
      response.write( session.contents( i ) & "<br/>" )
    next
  %>
```

## StaticObjects Collection을 통한 반복문
  
StaticObjects Collection을 반복하여 session 객체에 저장된 모든 개체의 값을 볼 수 있음.
  
```asp
  <%
    dim k
    for each k in session.staticObjects
      response.write( k & "<br/>" )
    next
  %>
```


---
  
- 출처 : [https://www.w3schools.com/asp/asp_sessions.asp](https://www.w3schools.com/asp/asp_sessions.asp)  
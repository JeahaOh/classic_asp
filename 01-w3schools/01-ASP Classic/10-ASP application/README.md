# ASP 응용프로그램 객체
  
어떤 목적을 수행하기 위해 함께 작동하는 ASP 파일 그룹을 응용 프로그램이라고 함.
  
---
  
## 응용프로그램 객체
  
웹상의 응용 프로그램은 어떤 목적을 수행하기 위해 함께 작동하는 여러 ASP 파일로 구성될 수 있음. Application 객체는 이러한 파일을 함께 묶는데 사용 됨.  
  
어플리케이션 개체는 세션 개체와 마찬가지로 모든 페이지의 변수를 저장하고 엑세스하는 데 사용됨. 차이점은 모든 사용자가 하나의 응용 프로그램 개체를 공유한다는 것임. (세션에는 각 사용자에 대해 하나의 세션 개체가 있음.)
  
어플리케이션 개체는 데이터베이스 연결 정보와 같은 어플리케이션의 많은 페이지에서 사용할 정보를 보유함. 정보는 모든 페이지에 엑세스 가능. 정보 변경도 한 곳에서 가능하며, 변경 사항은 자동으로 모든 페이지에 반영됨.  
  
---
  
## 어플리케이션 변수 저장 및 검색
  

어플리케이션 변수는 어플리케이션의 모든 페이지에서 엑세스하고 변경 가능함.  
"Global.asa"파일에서 어플리케이션 변수를 만들 수 있음.

```html
  <script language="vbscript" runat="server">
    // vartime, users라는 두개의 응용프로그램 변수 생성
    sub Application_OnStart
      application( "vartime" ) = ""
      application( "users" ) = 1
    end sub
  </script>
```

```asp
  There are <%=application( "users" )%> active connections.
```

  ---  

## 컨텐츠 컬렉션을 통해 반복문

Contents Collection에는 모든 어플리케이션 변수가 포함됨. 콘텐츠 컬렉션을 반복하여 컬렉션에 무엇이 저장되어 있는지 확인 가능.
  
```asp
  dim ac
  for each ac in application.contents
    response.write( ac & "<br/>" )
  next
```
  
contents에 대해 count 속성 사용 가능
  
```asp
  dim acCnt
  acCnt = application.contents.acCnt

  for ac = 1 to acCnt
    response.write( application.contents( ac ) & "<br/>" )
  next
```

  ---  

## StaticObjects collection을 통해 반복문

StaticObjects 컬렉션을 반복하여 어플리케이션 객체에 저장된 모든 개체의 값을 볼 수있음.
세션에선 안 됬는데 이게 될 까?  
  
```asp
  dim so
  
  for each so in application.staticojbects
    response.write( so & "<br/>" )
  next
```

  ---  

## 잠금 및 잠금 해제

"lock"으로 어플리케이션을 잠글 수 있음. 어플리케이션이 잠긴 경우 사용자는 현재 엑세스 중인 변수를 제외한 어플리케이션변수를 변경할 수 없음.  
"unlock"으로 어플리케이션의 잠금을 해제할 수 있음.  
  
```asp
  application.lock
  ' do come application object operation
  application.unlock
```

---
  
- 출처 : [https://www.w3schools.com/asp/asp_applications.asp](https://www.w3schools.com/asp/asp_applications.asp)  
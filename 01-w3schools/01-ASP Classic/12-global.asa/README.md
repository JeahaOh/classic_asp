# ASP Global.asa 파일

Globals.asa 파일은 ASP 어플리케이션의 모든 페이지에서 엑세스 할 수 있는 객체, 변수 및 메소드의 선언을 포함할 수 있는 선택적 파일임.  
  
유효한 모든 프라우져 스크립트(JS, VBScript, JScript, PerlScript)등 Globals.asa 파일내에서 사용 가능.  
  
Globals.asa 파일에는 다음 항목만 포함할 수 있음.  
- 어플리케이션 이벤트
- 세션 이벤트
- `<object>` 선언
- TypeLibrary 선언
- #include 지시문

Globals.asa 파일은 ASP 어플리케이션의 루트 디렉토리에 저장해야 하며 각 어플리케이션에는 Globals.asa 파일이 하나만 있을 수 있음

  ---  

## Globals.asa의 이벤트

Globals.asa에서 어플리케이션/세션이 시작될 때 수행할 작업과 어플리케이션/세션이 종료될 때 수행할 작업을 어플리케이션 및 세션 객체에 알릴 수 있음. 이에 대한 코드는 이벤트 핸들러에 배치됨. Globals.asa 파일에는 네 가지 유형의 이벤트가 포함될 수 있음.  
  
- **Application_OnStart** : First 사용자가 ASP 어플리케이션의 첫 페이지를 호출할 때 발생.  
  이 이벤트는 웹 서버를 다시 시작한 후 또는 Global.asa 파일을 편집 후에 발생함.  
  Session_OnStart 이벤트는 이 이벤트 직후에 발생함.
- **Session_OnStart** : 이 이벤트는 새 사용자가 ASP 어플리케이션에서 첫 페이지를 요청할 때마다 발생.
- **Session_OnEnd** : 이 이벤트는 사용자가 세션을 종료할 때마다 발생함.  
  사용자 세션은 지정된 시간(기본 20분) 동안 사용자가 페이지 요청하지 않으면 종료됨.
- **Application_OnEnd** : 이 이벤트는 마지막 사용자가 세션을 종료한 후에 발생함.  
  일반적으로 이 이벤트는 웹 서버가 중지될 때 발생함. 이 절차는 레코드 삭제 또는 텍스트 파일에 정보 쓰기와 같이 어플리케이션이 중지된 후 설정을 정리하는데 사용됨.

Globals.asa
```script
  <script language="vbscript" runat="server" >

    sub Application_OnStart
    'some code
    end sub

    sub Application_OnEnd
    'some code
    end sub

    sub Session_OnStart
    'some code
    end sub

    sub Session_OnEnd
    'some code
    end sub

  </script>

```

  ---  
  ---  
  

## **<객체>** 선언

`<object>` 태그를 사용하여 Globals.asa에서 세션 또는 어플리케이션 범위의 객체를 만들 수 있음.  
`<object>` 태그는 `<script>` 태그 밖에 있어야 함.

```xml
  <object runat="server" scope="scope" id="id" {porgid="progID" | classid="classID"}>
    ...
  </object>
```

- scope :  
  Sets the scope of the object (either Session or Application)
- id :  
  Specifies a unique id for the object
- progId :  
  An id associated with a class id. The format for ProgID is [Vendor.]Component[.Version]  
  Either ProgID or ClassID must be specified.  
- ClassId :  
  Specifies a unique id for a COM class object.  
  Either ProgID or ClassID must be specified.  

## ex

첫 번째 예는 ProgID 매개변수를 사용하여 MyAd라는 세션 범위 객체 생성임.  
`<obejct runat="server" scope="session" id="MyAd" progId="MSWC.AdRotator" > </object>`
  
두 번째 예는 ClassID 매개변수를 사용하여 MyConnection이라는 어플리케이션 범위 객체 생성임.  
```xml
  <object
    runat="server" scope="application" id="MyConnection"
    classid="Clsid:8AD3067A-B3FC-11CF-A560-00A0C9081C21">
  </object>
```
  
Global.asa 파일에 선언된 개체는 응용프로그램의 모든 스크립트에서 사용 가능함.  
  
```xml
  <object runat="server" scope="session" id="MyAd" progid="MSWC.AdRotator">
  </object>
```
```asp
  <%=MyAd.GetAdvertisement("/banners/adrot.txt")%>
```

  ---  
  ---  
  

## TypeLibrary 선언
  

TypeLibrary는 COM 개체에 해당하는 DLL 파일의 내용에 대한 컨테이너임. Global.asa 파일에 TypeLibrary에 대한 호출을 포함하면 COM 개체의 상수에 엑세스할 수 있으며 ASP 코드에서 오류를 더 잘 보고할 수 있음. 웹 어플리케이션이 형식 라이브러리에서 데이터 형식을 선언한 COM 객체에 의존하는 경우 Global.asa에서 형식 라이브러리를 선언할 수 있음.
  

### 문법
  
```xml
  <!-- METADATA TYPE="TypeLib"
    file="filename" uuid="id" version="number" lcid="localeid"
   -->
```
  
| Parameter | Description |
| :--- : | :---: |
| file | 유형 라이브러리의 절대 경로를 지정함. 파일 매개변수 또는 UUID 매개변수가 필요함. |
| UUID | 유형 라이브러리의 고유 식별자 지정. 파일 매개변수 또는 UUID 매개변수가 필요함. |
| version | 선택사항으로 버젼을 선택하는데 사용. 요청된 버젼이 없다면 가장 최근 벼전이 사용됨. |
| lcid | 선택사항으로 유형 라이브러리에 사용할 지역 식별자. |
  

### 오류 값
  
서버는 다음 오류 메세지중 하나를 반환할 수 있음.
  
| Parameter | Description |
| :--- : | :---: |
| ASP 0222 | 잘못된 라이브러리 사양 |
| ASP 0223 | 라이브러리 찾을 수 없음 |
| ASP 0224 | 라이브러리 로드 불가 |
| ASP 0225 | 라이브러리 래핑 불가 |
  
METADATA 태그는 Global.sas 파일의 모든 위치에 나타날 수 있음. (`<script>` 태그 내 외부).  
BUT, METADATA 태그는 Global.sas 파일의 상단에 표시하는 것이 좋음.
  

### 제한

Globals.asa 파일에 포함할 수 있는 항목데 대한 제한 사항 :  
- Global.asa 파일에 작성된 텍스트는 표시할 수 없음. 이 파일은 정보를 표시할 수 없음.
- Application_OnStart 및 Application_OnEnd 서브루틴에서는 서버 및 어플리케이션 개체만 사용할 수 있음.  
  Session_OnStart 서브루틴에서 내장 객체를 사용할 수 있음.
  Session_OnEnd 서브루틴에서 서버, 어플리케이션 및 세션 개체를 사용할 수 있음.  
  

### 서브루틴 사용법

Global.asa는 종종 변수를 초기화 하는 데 사용함.  
방문자가 웹 사이트에 처음 도착한 시간을 감지하는 방법임.  
시간을 "started"라는 세션 변수에 저장되며 "started" 변수의 값은 어플리케이션의 모든 ASP 페이지에 엑세스할 수 있음.  
  
```asp
  <script language="vbscript" runat="server" >
    sub Session_OnStart
      Session( "started" ) = now()
    end sub
  </script>
```
  
아래 예는 모든 신규 방문자를 다른 페이지로 리다이렉션 하는 방법
  
```asp
  <script language="vbscript" runat="server">
    sub Session_OnStart
      response.redirect( "newpage.asp" )
    end sub
  </script>
```

그리고 Global.asa 파일에 기능을 포함할 수 있음.  
  
Application_OnStart 서브루팅은 웹 서버가 시작될 때 발생함.  
그런 다음 Application_OnStart 서브루틴은 "getcustomers" 라는 다른 서브루틴을 호출함.  
"getcustomers" 서브루틴은 DB의 "customers" 테이블에서 데이터를 배열로 가져옴.  

```asp
  <script language="vbscript" runat="server">
    sub Application_OnStart
      getCustomers
    end sub

    sub getCustomers
      set conn = Server.CreateObject( "ADODB.Connection" )
      conn.Provider = "Microsoft.Jet.OLEDB.4.0"
      con.Open "c:/webdata/northwind.mdb"

      set rs = conn.execute( "SELECT NAME FROM customers" )
      
      Application( "customers" ) = rs.GetRows

      rs.close
      conn.close
    end sub
  </script>
```
  

### Global.asa 예

현재 방문자수를 계산하는 Global.asa 파일을 만들어 봄
- Application_OnStart는 서버가 시작될 때 어플리케이션 변수 "visitors"를 0으로 할당.
- Session_OnStart 서브루틴은 새 접속자가 접근할 때 마다 변수 "visitors"를 1씩 추가.
- Session_OnEnd 서브루틴은 접속자가 나갈 때 마다 변수 "visitors"를 1씩 감소.

```asp
  sub Application_OnStart
    Application( "visitors" ) = 0
  end sub

  sub Session_OnStart
    Application.lock
    Application( "visitors" ) = Application( "visitors" ) + 1
    Application.unlock
  end sub

  sub Session_OnEnd
    Application.lock
    Application( "visitors" ) = Application( "visitors" ) - 1
    Application.unlock
  end sub
```

```asp
  <head>
    There are <%=application( "visitors" )%> on online now!
  </head>
```

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
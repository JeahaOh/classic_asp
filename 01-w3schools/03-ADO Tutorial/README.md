# ADO Tutorial

ADO는 웹 페이지에서 DB에 엑세스하는데 사용함.  
선수 지식으로는 HTML, ASP, SQL임.  
  

## ADO란
  
- MS의 **사장된** 기술임
- ActiveX Data Objects의 약자임
- IIS와 함께 자동 설치됨
- DB의 데이터에 엑세스하기 위한 프로그래밍 인터페이스임
  

## ASP 페이지에서 DB 엑세스
  
ASP 페이지 내부에서 데이터베이스에 엑세스하는 일반적인 방법은 다음과 같음.
1. DB에 대한 ADO 연결 만들기
2. DB 연결 열기
3. ADO 레코드 집합 만들기
4. 레코드 세트 열기
5. 레코드세트에서 필요항 데이터 추출
6. 레코드 집합 닫기
7. 연결 닫기
  
  ---  
  ---  
  
# 데이터베이스 연결
  
웹페이지에서 DB에 엑세스 하려면 먼저 DB에 연결이 설정되어야 함.
  

## DSN 없는 데이터베이스 연결 만들기
  
DB에 연결하는 가장 쉬운 방법은 DSN 없는 연결을 사용하는 것임. DSN 없는 연결은 웹 사이트의 모든 MS Access DB에 대해 사용할 수 있음.  
  
"c:/webdata/"와 같은 웹 디렉토리에 "northwind.mdb"라는 DB가 있는 경우 다음 ASP 코드를 사용하여 DB에 연결할 수 있음.
  
```javascript
  set conn = server.createObject( "ADDODB.connection" )
  ' DB Driver 제공자와,
  conn.provider = "Microsoft.Jet.OLEDB.4.0"
  ' DB에 대한 물리적 경로를 지정해야 함.
  conn.open "c:/webdata/northwind.db"
```
  

## ODBC DB 연결 생성
  
"northwind"라는 ODBC DB가 있는 경우 다음 ASP 소스를 사용하여 DB에 연결할 수 있음.
  
```javascript
  set conn = server.createObject( "ADDODB.connection" )
  conn.open "northwind"
```
  

## MS Access DB에 대한 ODBC 연결
  
MS Access DB에 대한 연결을 만드는 방법은 다음과 같음.
  
1. 제어판에서 **ODBC 데이터 원본 설정** 실행
2. **시스템 DSN 탭** 선택
3. **추가**
4. MS Access Driver (*.mdb, *.accdb) 선택
5. 저장 경로 선택
6. DSN (Data Source Name) 부여
7. Click OK.
  
해당 구성은 웹 사이트가 있는 컴퓨터에서 수행해야 함. 개인 웹 서버(PWS) 또는 인터넷 정보 서버(IIS)를 실행하는 경우 작동함. 웹 사이트가 원격 서버에 있는 경우 해당 서버에 물리적으로 엑세스 할 수 있어야 함.
  

## ADO 연결 객체
  
ADO 연결 객체는 데이터 원본에 대한 열린 연결을 만드는 데 사용됨.  
이 연결을 통해 DB에 엑세스하고 조작 가능.  
ADO 연결 객체에 대해서는 추후 설명.
  
  ---  
  ---  
  
# ADO 레코드 셋
  
DB 데이터를 읽을 수 있으려면 먼저 데이터를 레코드 집합에 로드해야 함.
  

## ADO 테이블 레코드 집합 만들기
  
ADO 데이터베이스 연결이 생성된 후 이전 장에서 설명한 것처럼 ADO 레코드 집합을 생성할 수 있음.  
  
"northwind"라는 DB가 있다고 가정하에 다음 행을 사용하여 DB 내의 "customers" 테이블에 엑세스 할 수 있음.  
  
```javascript
  set conn = server.createObject( "ADODB.connection" )
  conn.provider = "Microsoft.Jet.OLEDB.4.0" 
  conn.open "c:/webdata/northwind.mdb"

  set rs = server.createObject( "ADODB.recordset" )
  rs.open "customers", conn
```
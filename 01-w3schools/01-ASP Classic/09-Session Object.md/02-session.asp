<% @CODEPAGE="65001" language="vbscript" %>
<% session.CodePage = "65001" %>
<% Response.CharSet = "utf-8" %>
<% Response.buffer=true %>
<% Response.Expires = 0 %>
<!DOCTYPE html>
<html>
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
  <style>
    * { background-color : black; color : white; padding : 1em;}
  </style>
</head>
<body>

  <h1>session object</h1>

  <h2>세션 변수 저장 및 검색</h2>
  <%
    session( "username" ) = "borry"
    session( "age" ) = 3
  %>

  Welcome <%response.write( session( "username" ) ) %> (<%=session( "age" )%>)
  
  <br/>

  <%IF session( "screenres" ) = "low" THEN%>
    This is the text version of the page
  <%ELSE%>
    This is the multimedia version of the page
  <%END IF%>

  <br/><hr/><br/>


  <h2>세션 변수 제거</h2>
  <%
    session( "isAlive" ) = true
  %>
  isAlive : <%=session( "isAlive" )%>
  <br/>
  <%
    if session.contents( "isAlive" ) = true then
      session.contents.remove( "isAlive" )
    end if
  %>
  isAlive : <%=session( "isAlive" )%>
  <br/>
  username : <%=session( "username" )%>
  <br/>
  <% session.contents.removeAll() %>
  session all ?
  <%if session.contents( "age" ) = "" then%>
    empty
  <%else%>
    <%=session.contents( "age" )%>
  <%end if%>

  <br/><hr/><br/>


  <h2>콘텐츠 컬렉션을 통한 반복문</h2>

  <%
    session( "username" ) = "borry"
    session( "age" ) = 3

    response.write( "<h3>키 값만 반환</h3>" )
    ' 키 값만 반환
    dim i
    for each i in session.contents
      response.write( i & "<br/>" )
    next

    response.write( "<h3>밸류값만 반환</h3>" )
    ' 밸류값만 반환
    for each i in session.contents
      response.write( session.contents( i ) & "<br/>" )
    next

    response.write( "<h3>키 값으로 밸류 값까지 반환</h3>" )
    ' 키 값으로 밸류 값까지 반환
    for each i in session.contents
      response.write( i & " : " & session.contents( i ) & "<br/>" )
    next

  %>

  <br/><hr/><br/>

  <h3>contnets collection의 항목 수</h3>

  <%
    y = session.contents.count
    response.write( "session variables : " & y & "<br/>" )

    for x = 1 to y
      response.write( session.contents( x ) & "<br/>" )
    next
  %>

  <br/><hr/><br/>

  <h2>StaticObjects 컬렉션을 통해 반복문</h2>

  <%=session.staticObjects%>
  <%
    dim k
    for each k in session.staticObjects
      response.write( k & "<br/>" )
    next
  %>

  <br/><hr/><br/>

  <h2>세션 종료</h2>
  <%
    session.timeout = 5
    session.abandon
  %>
  <%=session.timeout%> 분


</body>
</html>
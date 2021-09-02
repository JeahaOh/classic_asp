<% @CODEPAGE="65001" language="vbscript" %>
<%
  session.CodePage = "65001"
  ' Response.CharSet = "utf-8"
  response.buffer = true
  response.expires = -1
  'response.expiresAbsolute = #Aug 31,2021 09:30:00#
  response.contentType = "text/html"
%>
<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8">
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <style>
    * { background-color : black; color : white; margin : 1em; }
  </style>
  <title>etc</title>
</head>
<body>
  <head>
    There are <%=application( "visitors" )%> on online now!
  </head>

  <br/><hr/><br/>

  <h2> ASP Browser Capabilities </h2>
  <p>Find the type, capabilities, and version of each browser that visits your site.</p>

  <%
    Set browser = server.createObject( "MSWC.BrowserType" )
  %>

  <table border="1" width="65%">
    <tr>
      <td width="52%">Client OS</td>
      <td width="48%"><%=browser.platform%></td>
    </tr>
    <tr>
      <td >Web Browser</td>
      <td ><%=browser.browser%></td>
    </tr>
    <tr>
      <td>Browser version</td>
      <td><%=browser.version%></td>
    </tr>
    <tr>
      <td>Frame support?</td>
      <td><%=browser.frames%></td>
    </tr>
    <tr>
      <td>Table support?</td>
      <td><%=browser.tables%></td>
    </tr>
    <tr>
      <td>Sound support?</td>
      <td><%=browser.backgroundsounds%></td>
    </tr>
    <tr>
      <td>Cookies support?</td>
      <td><%=browser.cookies%></td>
    </tr>
    <tr>
      <td>VBScript support?</td>
      <td><%=browser.vbscript%></td>
    </tr>
    <tr>
      <td>JavaScript support?</td>
      <td><%=browser.javascript%></td>
    </tr>
    <tr>
      <td colspan="2">정상 작동 할 리가 없지</td>
    </tr>
  </table>

  <br/><hr/><br/>

  <h2> Contents Rotator </h2>
  <p>Display a different content each time a user visits a page (ASP 3.0)</p>

  <%
    'set cr = server.CreateObject("MSWC.ContentRotator")
    'response.write( cr.chooseContent( "text/contentRotator.txt" ) )
  %>
  <p>
    NOTE: Because the content strings are changed randomly in the text file, and this page has only four content strings to choose from, sometimes the page will display the same content strings twice in a row.
  </p>
  <p>IS IT DOESN'T WORK.</p>
  

  <br/><hr/><br/>

  <br/><hr/><br/>

</body>
</html>
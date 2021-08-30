<% @CODEPAGE="65001" language="vbscript" %>
<% 'session.CodePage = "65001" %>
<% 'Response.CharSet = "utf-8" %>
<% Response.buffer = true %>
<% Response.Expires = 0 %>
<!DOCTYPE html>
<html lang="ko">
<head>
  <!-- <meta charset="UTF-8"> -->
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <style>
    * { background-color : black; color : white; margin : 1em; }
  </style>
  <title>ASP 이메일 보내기</title>
</head>
<body>
  <head>
    There are <%=application( "visitors" )%> on online now!
  </head>
  <br/><hr/><br/>
  <h1>text email</h1>
  <%
    set mail = CreateObject( "CDO.Message" )

    with mail
      .to = "jeaha1217@idlook.co.kr"
      .from = "jeaha1217@idlook.co.kr"
      .subject = "TEST Send Email with CDO."
      ' .bcc = "jeaha1217@idlook.co.kr"
      ' .cc = "jeaha1217@idlook.co.kr"

      ' 메세지만 보낼때
      .textBody = "This Is A Message."


      ' HTML 메세지 보낼 때
      '.HTMLBody = "<p>This Is A Message.</p>"

      ' 웹사이트에서 웹 페이지를 보내는 HTML 이메일 
      '.createMHTMLBody "https://www.w3schools.com/asp/"

      ' HTML 파일을 보낼 때
      '.CreateMHTMLBody "file::/c:/.../*.html"

      ' 첨부 파일이 있는 문자 이메일 보내기
      '.textBody = "This Is A Message."
      '.addAttachment "file::/c:/.../*.html"

      mail.Configuration( "http://schemas.microsoft.com/cdo/configuration/sendusing" ) = 1

'      set objConf = mail.Configuration
'      set flds = objConf.fields
'
'      with flds
'        ' 원격 서버를 이용해서 이메일을 보낼 경우 2, 내부망일 경우 1
'        flds.item( "http://schemas.microsoft.com/cdo/configuration/sendusing" ) = 2
'        ' Name Or IP remote SMTP server
'        flds.item( "http://schemas.microsoft.com/cdo/configuration/smtpserver" ) = "smtp.daum.net"
'        ' server port
'        flds.item( "http://schemas.microsoft.com/cdo.configuration/smtpserverport" ) = 465
'
'        ' 계정이름
'        Flds("http://schemas.microsoft.com/cdo/configuration/smtpaccountname") = ""
'
'        Flds("http://schemas.microsoft.com/cdo/configuration/sendemailaddress") = "no_reply_dev@kakao.com"
'        Flds("http://schemas.microsoft.com/cdo/configuration/smtpuserreplyemailaddress")= "no_reply_dev@kakao.com"
'        ' cdoBasic
'        Flds("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
'        ' 메일서버의 계정ID
'        Flds("http://schemas.microsoft.com/cdo/configuration/sendusername") = "no_reply_dev@kakao.com"
'        ' 메일서버의 계정 비밀번호
'        Flds("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "DPpHs^~~6429"
'
'        flds.update
'      end with

      .send
    end with

    set mail = nothing
  %>
  no_reply_dev@kakao.com



</body>
</html>
<% @CODEPAGE="65001" language="vbscript" %>
<% 'session.CodePage = "65001" %>
<% 'Response.CharSet = "utf-8" %>
<% Response.buffer = true %>
<% Response.Expires = 0 %>
<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8">
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
    ' 메일 내용을 담을 객체
    dim mailcontents
    dim iMsg, iConf

    ' 렌덤 키를 만들어서 메일 내용에 넣어주자
    ' 랜덤 key BGN
    randomize

    dim random, mailkey
    random = array( int( (122 - 97) * rnd + 97 ), int( (122 - 97) * rnd + 97 ), int( (122 - 97) * rnd + 97 ), int( (122 - 97) * rnd + 97 ), int( (9999 - 1000) * rnd + 1000) )

    mailkey = chr( random(0) ) & "" & chr( random(1) ) & "" & chr( random(2) ) & "" & chr( random(3) ) & "" & random(4)

    mailcontents = "<p>"
    mailcontents = mailcontents & "랜덤 키 : " & mailkey
    mailcontents = mailcontents & "</p>"

    response.write mailcontents
    ' 랜덤 key END

    set iConf = server.createObject( "CDO.Configuration" )

    with iConf.fields
      ' 1일 경우 로컬(SMTP), 2일 경우 외부(SMTP) 원격지로 메일전송
      .item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2

      ' mailroot 폴더 권한 설정
      ' 로컬 smtp를 이용해 보내는 경우 pickup 디렉토리를 이용하는데 폴더에 적절한 권한을 부여
      ' .item("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory") = "C:\Inetpub\mailroot\Pickup"

      ' Host 설정 메일서버 IP 또는 메일서버 URL 
      .item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.daum.net"

      '접속 시도할 제한 시간을 설정(30초)
      .item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30

      ' 메일 서버 포트
      .item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465

      .item("http://schemas.microsoft.com/cdo/configuration/smtpaccountname") = "no_reply_dev@kakao.com"
      .item("http://schemas.microsoft.com/cdo/configuration/sendmailaddress") = "no_reply_dev@kakao.com"

      ' 회신 메일 주소
      .item("http://schemas.microsoft.com/cdo/configuration/smtpuserreplyemailaddress") = "no_reply_dev@kakao.com"

      ' cdo basic
      .item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1

      ' 메일서버 계정 ID
      .item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "no_reply_dev@kakao.com"

      ' 메일서버 계정 비밀번호
      .item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = ""

      .Update
    end with


    ' 메일 객체 생성
    set iMsg = server.createObject( "CDO.Message" )
      ' 로컬이 아닌 외부 SMTP 
      iMsg.configuration = iConf

      set iConf = nothing

      with iMsg
        .to = "jeaha1217@idlook.co.kr"
        .from = "jeaha1217@idlook.co.kr"
        .subject = "TEST Send Email with CDO."

'      ' 메세지만 보낼때
'      .textBody = "This Is A Message."

        ' HTML 메세지 보낼 때
        .HTMLBody = mailcontents

        ' 메일 발송
        .send
      end with

      ' 메일 발송 개체 비활성화
      set iMsg = nothing


'       ' 참조
'       .cc = "jeaha1217@idlook.co.kr"
'       ' 숨은 참조
'       .bcc = "jeaha1217@idlook.co.kr"
'
'    with mail
'
'
'
'
'      ' 웹사이트에서 웹 페이지를 보내는 HTML 이메일 
'      '.createMHTMLBody "https://www.w3schools.com/asp/"
'
'      ' HTML 파일을 보낼 때
'      '.CreateMHTMLBody "file::/c:/.../*.html"
'
'      ' 첨부 파일이 있는 문자 이메일 보내기
'      '.textBody = "This Is A Message."
'      '.addAttachment "file::/c:/.../*.html"
'
'      mail.Configuration( "http://schemas.microsoft.com/cdo/configuration/sendusing" ) = 1
'
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
'        Flds("http://schemas.microsoft.com/cdo/configuration/smtpaccountname") = "jeaha1217"
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
'
'      .send
'    end with
'
'    set mail = nothing
  %>



</body>
</html>
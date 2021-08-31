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
  <title>VBScript</title>
</head>
<body>

  <head>
    There are <%=application( "visitors" )%> on online now!
  </head>

  <br/><hr/><br/>

  <h2> Redirect the user to another URL </h2>
  <code>
    <pre>
        &lt;%
          if request.form( "select" ) &lt;&gt; "" then
            response.redirect( request.form( "select" ) )
          end if
        %&gt;

        &lt;form action="example.asp" method="post"&gt;
          &lt;label for="select_server.asp"&gt;SERVER&lt;/label&gt;
          &lt;input
            type="radio" name="select" id="select_server.asp"
            value="select_server.asp"
          /&gt;
          &lt;br/&gt;
          &lt;label for="select_text.asp"&gt;TEXT&lt;/label&gt;
          &lt;input
            type="radio" name="select" id="select_text.asp"
            value="select_text.asp"
          /&gt;
          &lt;input type="submit" value="GO!" /&gt;
        &lt;/form&gt;
    </pre>
  </code>

  <%
    if request.form( "select" ) <> "" then
      response.redirect( request.form( "select" ) )
    end if
  %>

  <form action="example.asp" method="post">
    <label for="select_server.asp">SERVER</label>
    <input
      type="radio" name="select" id="select_server.asp"
      value="select_server.asp"
    />
    <br/>
    <label for="select_text.asp">TEXT</label>
    <input
      type="radio" name="select" id="select_text.asp"
      value="select_text.asp"
    />
    <br/>
    <input type="submit" value="GO!" />
  </form>

  <br/><hr/><br/>

  <h2> Random Links </h2>
  <code>
    <pre>
      randomize()
      r = rnd()
  
      if r > 0.5 then
        response.write( "<a href='https://www.w3schools.com/asp/asp_examples.asp'>w3school asp</a>" )
      else
        response.write( "<a href='javascript:void(0)'>nowhere</a>" )
      end if
    </pre>
  </code>

  <%
    randomize()
    r = rnd()

    if r > 0.5 then
      response.write( "<a href='https://www.w3schools.com/asp/asp_examples.asp'>w3school asp</a>" )
    else
      response.write( "<a href='javascript:void(0)'>nowhere</a>" )
    end if
  %>
  <br/>

  <br/><hr/><br/>

  <h2> Controlling Buffer </h2>
  <code>
    <pre>
      
    </pre>
  </code>

  <% response.buffer = true %>
  <br/>
  <p>
    This text wil be sent to your browser when my response buffer is flushed!
  </p>
  <% response.flush %>

  <br/><hr/><br/>

  <h2> Clear the buffer </h2>
  <code>
    <pre>
      This is some text I want to send to you.
      No, I changed my mind. I wnat to clear the text.
      &lt;% response.clear %&gt;
    </pre>
  </code>

  <% response.flush %>

  <p>This is some text I want to send to you.</p>
  <p>No, I changed my mind. I wnat to clear the text.</p>
  <% response.clear %>
  <br/>

  <br/><hr/><br/>

  <h2> End a script in the middle of processing </h2>
  <code>
    <pre>
      &lt;p&gt;
        I am writing some text. This text will never be
        &lt;br&gt;
        &lt;% response.end %&gt;
        finished! It's too late to write more!
      &lt;/p&gt;
    </pre>
  </code>

  <p>
    I am writing some text. This text will never be
    <br>
    <% ' response.end %>
    finished! It's too late to write more!
  </p>

  <br/>

  <br/><hr/><br/>

  <h2> Set a date/time when a page cached in a browser will expire </h2>
  <code>
    <pre>
      response.expiresAbsolute = #Aug 31,2021 09:30:00#
    </pre>
  </code>



  <br/>

  <br/><hr/><br/>

  <h2> Check if the user is still connected </h2>
  <code>
    <pre>
      if response.isClientConnected = true then
        response.write( "the user is still connected!" )
      else
        response.write( "the user is NOT connected!" )
      end if
    </pre>
  </code>

  <%
    if response.isClientConnected = true then
      response.write( "the user is still connected!" )
    else
      response.write( "the user is NOT connected!" )
    end if
  %>

  <br/>

  <br/><hr/><br/>

  <h2> Set the type of content </h2>
  <code>
    <pre>
      response.contentType = "text/html" 
    </pre>
  </code>
  <br/>

  <br/><hr/><br/>

  <h2> Set the name of character set </h2>
  <code>
    <pre>
      &lt;% response.charset = "utf-8" %&gt;
    </pre>
  </code>

  <br/>

</body>
</html>
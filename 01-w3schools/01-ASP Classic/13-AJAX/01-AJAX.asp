<% @CODEPAGE="65001" language="vbscript" %>
<% session.CodePage = "65001" %>
<% Response.CharSet = "utf-8" %>
<% Response.buffer=true %>
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
  <title>The XMLHttpRequest Object</title>
</head>
<body>
  <head>
    There are <%=application( "visitors" )%> on online now!
  </head>
  <h1>Start typing a name in the input field below</h1>
  <br/><hr/><br/>

  <form action="">
    <label for="txt">INPUT</label>
    <input
      type="text" name="txt" id="txt"
      onkeyup="showHint(this.value)"
    />
  </form>

  <p>Suggestions : <span id="hint"></span></p>

  <script>
    const showHint = ( str ) => {
      let xhttp;
      
      if( str.length == 0 ) {
        document.getElementById( 'hint' ).innerHTML = '';
        return;
      }

      xhttp = new XMLHttpRequest();
      xhttp.onreadystatechange = function () {
        if( this.readyState == 4 && this.status == 200 ) {
          document.getElementById( 'hint' ).innerHTML = this.responseText;
        }
      };

      xhttp.open( 'GET', '01-getHint.asp?q=' + str, true);
      xhttp.send();
    }
  </script>

</body>
</html>
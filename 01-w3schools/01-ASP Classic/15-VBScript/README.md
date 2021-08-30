# VBScript

## VBScript 반복문

### for ... next

```script
  for i = 0 to 5
    response.write( "The Number Is : " & i & "<br/>" )
  next
```

### for ... each

```script
  dim cars(2)
  cars(0) = "Porsche"
  cars(1) = "BMW"
  cars(2) = "Jaguar"

  for each x in cars
    response.write( x & "<br/>" )
  next
```

### do while

```script
  i = 0
  do while i < 10
    response.write( i & "<br/>" )
    i = i + 1
  loop
```

  ---  
  ---  

## VBScript 날짜와 시간

### date and time

```script
  <%=date()%>
  <%=time()%>
```

### Get The Name Of A Day

```script
  <p>
    VBScripts' function <b>WeekdayName</b> is used to get a weekday:
  </p>
  <%
    response.Write(WeekDayName(1))
    response.Write("<br>")
    response.Write(WeekDayName(2))
  %>

  <p>Abbreviated name of a weekday:</p>
  <%
    response.Write(WeekDayName(1,true))
    response.Write("<br>")
    response.Write(WeekDayName(2,true))
  %>

  <p>Current weekday:</p>
  <%
    response.Write(WeekdayName(weekday(date)))
    response.Write("<br>")
    response.Write(WeekdayName(weekday(date), true))
  %>
```

### Get the name of a month

```script
  Tody it is <%=weekdayName( weekday( date ) )%>
  <br/>
  and the month is <%=monthName( month( date ) )%>
```

### Countdown to year 3000

```script
  <p>Countdown to year 3000 : </p>
  <%
    millennium = cdate( "1/1/3000 00:00:00" )
  %>

  It is <%=dateDiff( "yyyy", now(), millennium )%> years to year 3000!
  <br/>
  It is <%=dateDiff( "m", now(), millennium )%> months to year 3000!
  <br/>
  It is <%=dateDiff( "ww", now(), millennium )%> weeks to year 3000!
  <br/>
  It is <%=dateDiff( "d", now(), millennium )%> days to year 3000!
  <br/>
  It is <%=dateDiff( "h", now(), millennium )%> hours to year 3000!
  <br/>
  It is <%=dateDiff( "n", now(), millennium )%> minutes to year 3000!
  <br/>
  It is <%=dateDiff( "s", now(), millennium )%> seconds to year 3000!
  <br/>
```

### Calculate the day which is n days from today

```script
  <%=dateAdd( "d", 30, date() )%> is a date 30 days from today.
```

### Format date and time

```script
  <%=formatDateTime( date(), vbgeneraldate )%> <br/>
  <%=formatDateTime( date(), vblongdate )%> <br/>
  <%=formatDateTime( date(), vbshortdate )%> <br/>
  <%=formatDateTime( now(), vblongtime )%> <br/>
  <%=formatDateTime( now(), vbshorttime )%> <br/>
```

### is this a date?

```script
  somedate = "04/27/92"
  response.write( isDate( somedate ) & "<br/>" )
  somedate = "04-27-92"
  response.write( isDate( somedate ) & "<br/>" )
  somedate = "1992-04-27"
  response.write( isDate( somedate ) & "<br/>" )
```

## VBScript 내장함수

### Uppercase or Lowercase a string

```script
  dim name
  name = "Asdf Qwer"
  
  ucase( name )
  lcase( name )
```

### Trim a string

```asp
  name = "  Jeaha . dev  "

  response.write("visit" & name & "now<br/>")
  response.write("visit" & trim(name) & "now<br/>")
  response.write("visit" & ltrim(name) & "now<br/>")
  response.write("visit" & rtrim(name) & "now<br/>")
```

### How to reverse a string?

```asp
  txt = "Hell the Everyone!"
  response.write( strReverse( txt ) )
```

### How to round a number?

```asp
  i = 48.66776677
  j = 48.3333333
  response.write(Round(i))
  response.write("<br>")
  response.write(Round(j))
```

### A random number

```asp
  ' 난수 생성
  randomize()
  r = rnd()
  response.write("A random number: <b>" & r & "</b><br/>")

  ' 0 ~ 100 사이의 난수
  randomNumber = Int( 100 * r )
  response.write("A random number: <b>" & randomNumber & "</b><br/>")
```

### Return a specified number of characters from left/right of a string

```asp
  txt = "Helcome to this WEB"
  response.write( txt )
  response.write( "<br/>" )
  response.write( left( txt, 5 ) )
  response.write( "<br/>" )
  response.write( right( txt, 5 ) )
```

### Replace some characters in a string

```asp
  response.write( replace( txt, "WEB", "HELL!!" ) )
```

### Return a specified number of characters from a string

```asp
```



- 출처 : [https://www.w3schools.com/asp/asp_examples.asp](https://www.w3schools.com/asp/asp_examples.asp)  
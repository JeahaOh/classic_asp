# VBScript 반복문

동일한 코드 블록을 지정된 횟수만큼 실행.

VBScript에는 4개의 반복문이 있음.
- **for ... next** : 지정된 횟수만큼 코드 실행
- **for each ... next** : 컬렉션 각 항목 또는 배열의 각 요소에 대한 코드 실행
- **do ... loop** : 조건이 true 일 때까지 반복
- **while ... w end** : 비추

## for ... next

**for ... next**문을 이용해서 코드 블럭을 지정된 횐수만큼 수행한다.  
**for**문의 **i**라는 카운터 값은 시작부터 종료를 세고, **next**문은 i를 1씩 증가시킨다.

```asp
  <% @CODEPAGE="65001" language="vbscript" %>
  <% Response.CharSet = "utf-8" %>
  <!DOCTYPE html>
  <html>
  <head>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
  </head>
  <body>
    <%
      for i = 0 to 5
        response.write( "i : <strong>" & i & "</strong><br/>" )
      next
    %>
  </body>
  </html>
```

### step 키워드

step 키워드를 사용하면 특정한 값으로 값을 증감 시킬 수 있음.

```c
  for i = 0 to 10 step 2
    response.write( "i : <strong>" & i & "</strong><br/>" )
  next
```
```c
  for i = 10 to 2 step -2
    response.write( "i : <strong>" & i & "</strong><br/>" )
  next
```

### exit a for ... next

**exit for** 키워드를 사용해서 for ... next 문을 종료할 수 있음.

```c
  for i = 0 to 10
    if i = 5 then exit for
    response.write( "i : <strong>" & i & "</strong><br/>" )
  next
```
  
  
--- 
  
## for each ... next loop

**for each ... next loop** 반복문은 콜렉션의 각 아이템이나, 배열의 각 요소에 대한 코드를 반복한다.

```c
  dim cars(2)
  cars(0) = "Volvo"
  cars(1) = "SAAB"
  cars(2) = "BMW"

  for each e in cars
    response.write( e & "<br/>" )
  next
```
  
  
---
  
## do ... loop

원하는 반복 횟수를 모르는 경우 **do ... loop**문을 사용함.  
**do ... loop**문은 조건이 참인 동안 또는 조건이 참이 될 때까지 코드를 반복함.

### 조건이 참인 동안 반복
  
while 키워드를 이용하여 do ... loop 문의 조건을 확인함
  

```asp
  do while i < 10
    response.write( "i : " & i )
    i = i + 1
  loop
```
i가 9 라면 반복문은 실행되지 않는다.  
  
```asp
  do
    response.write( "i : " & i )
    i = i + 1
  loop while i < 10
```
i가 10보다 작더라도 한번은 실행하고 본다.

### 조건이 참이 될때까지 반복

**until** 키워드를 사용하면 **do ... loop** 문의 조건을 확인함
  
```asp
  i = 0
  do until i = 10
    response.write( "i : <strong>" & i & "</strong><br/>" )
    i = i + 1
  loop
```
  
i가 10이면 코드가 실행되지 않음
  
```asp
  i = 0
  do
    response.write( "i : <strong>" & i & "</strong><br/>" )
    i = i + 1
  loop until i = 10
```
  
---
  
## do ... loop 문 종료

**exit do** 키워드를 사용하면 **do ... loop** 문 종료 가능

```asp
  i = 20
  do until i = 10
    i = i - 1
    if i < 10 then exit do
    response.write( "i : <strong>" & i & "</strong><br/>" )
  loop
```

i가 10과 다르고, 10보다 크면 실행

---

- 출처 : [https://www.w3schools.com/asp/asp_looping.asp](https://www.w3schools.com/asp/asp_looping.asp)  
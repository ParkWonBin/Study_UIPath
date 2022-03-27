# String 관련
#### Integer To String
```vb
'날짜 표시 형식
cint("1").ToString("0000") '= 0001  
"1".PadLeft(4,"0"c) '= 00001
```

#### DateTime To String
[한글 요일 표시](https://m.blog.naver.com/PostView.naver?isHttpsRedirect=true&blogId=elduque&logNo=120096308343)
- 1단계 : import 패널에서 System.Globalization 추가(CultureInfo 객체 사용을 위함)  
- 2단계 : writeLine 이나 LogMessage에서 출력값 확인하기목요일  

```vb
'날짜 표시 형식
now.ToString("yyyy_MM_dd")
'한글 날짜 표시
DateTime.Today.ToString("dddd", CultureInfo.CreateSpecificCulture("ko-KR"))  #목요일
DateTime.Today.ToString("ddd", CultureInfo.CreateSpecificCulture("ko-KR"))   #목
Date.ParseExact("20210212", "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)  
```

#### String To Charactor
```vb
' 문자열 & 아스키코드
Asc("A") '= 65 
Chr(65) '= "A"

' ShortCode : 문자열 - 아스키코드 번호
join(str_tmp.ToCharArray.Select(function(x) string.Format("{0} : {1}",x,asc(x).ToString) ).ToArray, vbNewLine)

' CSV 열 구분 : chr(44) = ','
' CSV 행 구분 : chr(13)+chr(10) = \r\n
' 엑셀 셀 내부 줄바꿈 : chr(10) 
```

' 파일명 제어
StrArr = System.IO.Directory.GetFiles("절대경로") '각 파일의 절대경로 얻음
StrArr = System.IO.Directory.GetFiles(Environment.CurrentDirectory) '프로젝트 경로파일 얻음

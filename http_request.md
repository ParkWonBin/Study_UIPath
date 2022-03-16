#### API 문서 보기
오케스트레이터 주소에서 "/swagger" 추가하여 입력하면 해당 오케에서 사용할 수 있는 API 문서를 볼 수 있다.   
[UIpath 공식문서](https://docs.uipath.com/orchestrator/reference/api-references) 참고. 사용하는 오케의 버전을 찾줘서 확인 바람.  
Tenant를 여러개 쓰는 경우 Autorization 통해 Token 발급받고 해당 토큰으로 API 접근하는 것이 좋다.  


#### vb로 http 요청 받기1
```vb
Dim Str_URL As String = "http://api.hostip.info/?ip=68.180.206.184"
'--------------------------- HTTP Request 받기 1 ---------------------------
Dim webClient As New System.Net.WebClient
webClient.Headers.Add("KEY","VALUE")
Dim result As String = webClient.DownloadString(Str_URL)
console.WriteLine(result)
' 출처 : https://stackoverflow.com/questions/92522/http-get-in-vb-net
```

#### vb로 http 요청 받기2
```vb
Dim Str_URL As String = "http://api.hostip.info/?ip=68.180.206.184"
'--------------------------- HTTP Request 받기 2 ---------------------------
Dim client As New System.Net.WebClient
client.Headers.Add("KEY","VALUE")
Dim data As System.IO.Stream = client.OpenRead(Str_URL)
Dim reader As System.IO.StreamReader =  New System.IO.StreamReader(data)
Dim s As String = reader.ReadToEnd()
data.Close()
reader.Close()
Console.WriteLine(s)
' 출처 : https://docs.microsoft.com/ko-kr/dotnet/api/system.net.webclient?view=net-6.0
```

# MailBoxToDataTable

invoke code 통해 사용할 수 있다. 


#### 목적 : 
Log함수 안에서 해당 함수를 쓰면, log와 값 저장을 동시에 할 수 있다.
- envGet("변수명") : 환경변수의 값을 반환합니다.
- envSet("변수명","값") : 환경변수에 값을 설정하고 "값"을 반환합니다.
- envApp("변수명","값") : ":"를 구분자로 값을 누적하고 "값"을 반환합니다.
- envPop("변수명") : ":"를 구분자로 마지막 값을 제거하고 반환합니다.

#####  Func_MailBox2DataTable
 
- List_MailBox :
  - Get Outlook에서 받아온 List of MailMessage 
- StrArr_Columns : 
  - Split("Subject|Body|Sender|To|Cc|Bcc|Uid|Date|DateCreated|DateRecieved|HtmlBody|PlainText|Size|Attach", "|")
  - 해당 열 순서대로 등록된 예약어에 해당하는 Data를 체웁니다.
  - 등록되지 않은 예약어(열)는 Null값으로 체워집니다.
  - 예약어는 ToUpper, Trim 처리가 되어있습니다.
- Retrun : DataTable
 
- 주의사항 : 
- - Body, HtmlBody, PlainText 는 안에 들어있는 
- - Data 가 크기 때문에 불필요시 넣지 말것


# MailBoxToDataTable

invoke code 통해 사용할 수 있다. 

```vb
' 인수설정
' in : List_MailBox // Get Outlook에서 받아온 Collection
' in : StrArr_Columns // 예시 :  Split("Subject|Body|Sender|To|Cc|Bcc|Uid|Date|DateCreated|DateRecieved|HtmlBody|PlainText|Size|Attach", "|")
' out : Dt_MailData  // 해당 열 순서대로 DataTable 만들어짐
' 주의사항 : Body, HtmlBody, PlainText 는 안에 들어있는 Data 가 크기 때문에 불필요시 넣지 말것

' DataTable 생성
Dt_MailData = New DataTable()
For Each colName As String In StrArr_Columns
   Dt_MailData.Columns.Add(colName.Trim)
Next

' 메일함 조회
Dim StrArr_NewRow As String()
Dim index As Int32
Dim StrArr_HeaderAttar As String()
Dim strArr_HeaderAttarUp  As String()
Dim IndexHeader As Int32

' 열이름에 대소문자, 공백 등 삭제
StrArr_Columns = StrArr_Columns.Select(Function(x) x.Trim.ToUpper).ToArray
StrArr_HeaderAttar = {"Uid","Date","DateCreated","DateRecieved","HtmlBody","PlainText","Size"}
strArr_HeaderAttarUp = StrArr_HeaderAttar.select(Function(x) x.tostring.ToUpper).ToArray

For Each mail As MailMessage In List_MailBox
   ' New Row 초기화
   StrArr_NewRow = New String(StrArr_Columns.Count-1){}
   index = 0
   For Each colName As String In StrArr_Columns
      If colName = "SENDER" Then
         StrArr_NewRow(index) = mail.Sender.Address
      Else If colName = "SUBJECT" Then
         StrArr_NewRow(index) = mail.Subject
      Else If colName = "BODY" Then 
         StrArr_NewRow(index) = mail.Body
      Else If colName = "TO" Then
         StrArr_NewRow(index) = String.Join( ";", mail.To.Select(Function(x) x.Address).ToArray)
      Else If colName = "CC" Then
         StrArr_NewRow(index) = String.Join( ";", mail.Cc.Select(Function(x) x.Address).ToArray)
      Else If colName = "BCC" Then
         StrArr_NewRow(index) = String.Join( ";", mail.Bcc.Select(Function(x) x.Address).ToArray)
      Else If  colName = "ATTACH" Then
         StrArr_NewRow(index) = String.Join(vbNewLine, mail.Attachments.Select(Function(x) x.Name).ToArray)         
      Else      
         IndexHeader = Array.IndexOf(strArr_HeaderAttarUp, colName)
         If IndexHeader <> -1 Then
            StrArr_NewRow(index) = mail.Headers(StrArr_HeaderAttar(IndexHeader))
         End If 
      End If 
      index = index +1
   Next
   Dt_MailData.Rows.Add(StrArr_NewRow)
Next
```

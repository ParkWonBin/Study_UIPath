Dim Fnc_Mail_GetAttr As System.Func(Of System.Net.Mail.MailMessage,String,String) = Function(msg_mail  As System.Net.Mail.MailMessage, Str_AttarName As String) As String
  ' 2022.04.05|wbpark|MailMessage 속성별 호출 방법 통일 시키기. 대소문자 구분 안함
  Str_AttarName=Str_AttarName.Trim.ToUpper
  If Str_AttarName = "SENDER" Then
    Return msg_mail.Sender.Address
  Else If Str_AttarName = "SUBJECT" Then
    Return msg_mail.Subject
  Else If Str_AttarName = "BODY" Then 
    Return msg_mail.Body
  Else If Str_AttarName = "TO" Then
    Return String.Join( ";", msg_mail.To.Select(Function(x) x.Address).ToArray)
  Else If Str_AttarName = "CC" Then
    Return String.Join( ";", msg_mail.CC.Select(Function(x) x.Address).ToArray)
  Else If Str_AttarName = "BCC" Then
    Return String.Join( ";", msg_mail.Bcc.Select(Function(x) x.Address).ToArray)
  Else If Str_AttarName = "ATTACH" Then
    Return String.Join(vbNewLine, msg_mail.Attachments.Select(Function(x) x.Name).ToArray)         
  Else
    Dim StrArr_HeaderKeys = "Uid|Date|DateCreated|DateRecieved|Size|Body|HtmlBody|PlainText".ToUpper().Split("|"c)
    Dim IndexHeader As Integer = Array.IndexOf(StrArr_HeaderKeys, Str_AttarName.ToUpper.Trim)
    If IndexHeader <> -1 Then
      Return msg_mail.Headers(StrArr_HeaderKeys(IndexHeader))
    Else
      Return "No AttarName : "+Str_AttarName
    End If 
  End If 
End Function
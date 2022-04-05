Dim Fnc_MailList_To_DataTable As System.Func(Of List(Of System.Net.Mail.MailMessage), String(), System.Data.DataTable)= Function(List_MailBox As List(Of System.Net.Mail.MailMessage), StrArr_Columns As String()) As System.Data.DataTable
  ' 2022.04.05|wbpark|메일함 데이터 DataTable로 만들어 저장
  Dim Dt_MailData As New DataTable
  For Each colName As String In StrArr_Columns
    Dt_MailData.Columns.Add(colName.Trim, System.Type.GetType("System.String"))
  Next
  For Each mail As System.Net.Mail.MailMessage In List_MailBox
    Dim StrArr_NewRow As String() = StrArr_Columns.Select(Function(x) Fnc_Mail_GetAttr(mail,x) ).ToArray
    Dt_MailData.Rows.Add(StrArr_NewRow)
  Next
  Return Dt_MailData
  ' 예시 : StrArr_ColNames = Sender|Subject|To|Cc|Bcc|Attach|Uid|Date|DateCreated|DateRecieved|Size|Body|HtmlBody|PlainText".Split("|"c).Distinct.toArray
  ' out_dt = Fnc_Convert_MailList_To_DataTable(in_list_mail,StrArr_ColNames)
End Function
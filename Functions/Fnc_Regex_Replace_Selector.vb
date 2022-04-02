Dim Fnc_Regex_Replace_Selector As System.Func(Of String, String, String) = Function(Str_Source As String, Str_ptn As String) As String
  ' 2022.04.02|wbpark|셀렉터 내 변수 부분 {{}}로 바꿔주기
  For Each x As System.Text.RegularExpressions.Match In System.Text.RegularExpressions.Regex.Matches(Str_Source,Str_ptn) 
    If Not (x.Tostring.ToUpper.Contains(".TOSTRING") OrElse x.ToString.Contains("(") OrElse  x.ToString.Contains(")") ) Then
      Dim Str_replaceTo As String = x.Tostring.replace("[&quot;","").replace("&quot;+","{{").replace("+&quot;","}}").replace("&quot;]","")
      Str_Source=Str_Source.replace(x.tostring,Str_replaceTo)
    End If
  Next
  Return Str_Source
End Function
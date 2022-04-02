Dim Fnc_Text_Replace As System.Func(Of String, String(), String(),String) = Function(Str_Source As String, Replace_Before As String(), Replace_After As String()) As String
  '2022.04.02|wbpark|입력된 문자열 일괄 replace
  For i As Integer = 0 To Replace_Before.length-1
    Str_Source=Str_Source.replace(Replace_Before(i),Replace_After(i))
  Next 
  Return Str_Source
End Function
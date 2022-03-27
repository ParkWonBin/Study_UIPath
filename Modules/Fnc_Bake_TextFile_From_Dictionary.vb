'Config.log
Dim Fnc_Back_Config As System.Func(Of System.Collections.Generic.Dictionary(Of String, String),String) = Function(dic_config As System.Collections.Generic.Dictionary(Of String, String)) As String
  Dim str_dic As String = "New Dictionary(Of String, String) From {"
  Dim str_path As String = System.IO.Path.Combine(System.Environment.CurrentDirectory,"Config")
  
  ' 내용 작성
  For Each key As String In dic_config.Keys
    Dim value As String = dic_config(key).ToString.Replace(System.Environment.NewLine," ").Replace(vbCr," ").Replace(vbLf," ").Replace(vbCrLf," ").Trim
    str_dic = str_dic + String.Format("{3}{1} {0}{4}{0} , {0}{5}{0} {2},","""","{","}",vbnewline,key,value)
  Next  
  str_dic = str_dic.Substring(0,str_dic.Length-1) + vblf + "}"

  ' 확장자 및 Index 확인
  Dim Str_tmp As String = ".log"
  Dim int_idx As Integer = 0
  While System.IO.File.Exists(str_path + Str_tmp)
    int_idx = int_idx+1
    Str_tmp = int_idx.ToString("(0)")+".log"
  End While

  Try ' 저장 및 로그 / 결과 반환
    System.IO.File.WriteAllText(str_path + Str_tmp, str_dic)
    System.Console.WriteLine("저장이 완료되었습니다." + vbnewline + str_path + Str_tmp)
  Catch ex As System.exception
	Throw New System.ApplicationException("Dictionary 저장 실패 " + vbnewline + ex.Message)
  End Try
  
  Return str_dic
End Function 
'Fnc_Back_Config(  New System.Collections.Generic.Dictionary(Of String, String) From {{"key","value"},{"k2","v2"}}  )}

' Write Text File
Dim Fnc_Back_Config_ShortCode As System.Func(Of System.Collections.Generic.Dictionary(Of String, String),String) = Function(dic_config As System.Collections.Generic.Dictionary(Of String, String)) As String
  Dim Str_Result = String.Format("{0}{0}{3}{0}{2}",vbNewLine,"New Dictionary(Of String,String) From {","}", Join( Dic_Config.Keys.Select(Function(key) String.Format("{0} {2}{3}{2} , {2}{4}{2} {1}", "{","}", chr(34), key, System.Convert.ToString(Dic_Config(key)).Replace(vbNewLine," ").Replace(chr(10)," ").Replace(chr(34),"'") ) ).ToArray, ","+vbNewLine) )
  System.IO.File.WriteAllText(System.IO.Path.Combine(System.Environment.CurrentDirectory,"Config.log",str_result))
  Return Str_Result
End Function
'Fnc_Back_Config(  New System.Collections.Generic.Dictionary(Of String, String) From {{"key","value"},{"k2","v2"}}  )}
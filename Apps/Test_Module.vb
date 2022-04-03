Dim Fnc_Kill_Process_By_Name As System.Func(Of String, String) = Function(ProcessName As String) As String
  '2022.04.04|wbpark|DRM 있는 엑셀 종료를 위해 제작
  Dim Arr_process As System.Diagnostics.Process() = System.Diagnostics.Process.GetProcesses.Where(Function(x) x.ProcessName.ToUpper = ProcessName.ToUpper).toArray
    While Arr_process.Count>0
        For Each p As System.Diagnostics.Process In Arr_process
            p.kill()
        Next
        System.Threading.Thread.Sleep(1000)	'1초 딜레이
        Arr_process=System.Diagnostics.Process.GetProcesses.Where(Function(x) x.ProcessName.ToUpper = ProcessName.ToUpper).toArray
    End While 
    Return ProcessName+" 종료"
End Function
'-----------------------
Dim Fnc_Back_Config As System.Func(Of System.Collections.Generic.Dictionary(Of String, String),String) = Function(dic_config As System.Collections.Generic.Dictionary(Of String, String)) As String
  '2022.04.04|wbpark|Config Value에 null값이 있으면 에러가 발생합니다.
  Dim Str_Result = String.Format("{0}{0}{3}{0}{2}",vbNewLine,"New Dictionary(Of String,String) From {","}", Join( Dic_Config.Keys.Select(Function(key) String.Format("{0} {2}{3}{2} , {2}{4}{2} {1}", "{","}", chr(34), key, System.Convert.ToString(Dic_Config(key)).Replace(vbNewLine," ").Replace(chr(10)," ").Replace(chr(34),"'") ) ).ToArray, ","+vbNewLine) )
  System.IO.File.WriteAllText(System.IO.Path.Combine(System.Environment.CurrentDirectory,"Config.log",str_result))
  Return Str_Result
  'Fnc_Back_Config( New System.Collections.Generic.Dictionary(Of String, String) From {{"key","value"},{"k2","v2"}}  )}
End Function
'-----------------------
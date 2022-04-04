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
Dim Fnc_Set_Env_With_Config As System.Func(Of System.Collections.Generic.Dictionary(Of String, Object), String) = Function(Dic_Config As System.Collections.Generic.Dictionary(Of String, Object) ) As String
  '2022.04.04|wbpark|Config값 환경변수로 넣기
  Dim Str_Error As String =""
  For Each Key As String In Dic_Config.Keys
    Try
      System.Environment.SetEnvironmentVariable(Key,in_Dic_Config(Key).tostring )
    Catch ex As System.Exception
		Str_Error=Str_Error+ex.message+vbnewline
    End Try
  Next
  System.Console.WriteLine(Str_Error)
  Return Str_Error
  'Fnc_Set_Env_With_Config(in_Dic_Config)
End Function
'-----------------------
Dim Fnc_Bake_UiElementAttar As System.Func(Of UiPath.Core.UiElement, String) = Function(ui_tmp As UiPath.Core.UiElement) As String
  '2022.04.04|wbpark|Uipath Activity와 함께 써야하는 함수. Invoke Code 내에서 Uipath Core 함수 호출 방법 불명.
  Dim Str_result As String  = ""
  Dim Str_fileName As String = Environment.CurrentDirectory+"\UiElementAttar.log"
  Dim dic_attar As dictionary(Of String, String) = ui_tmp.GetNodeAttributes(False)
  For Each key As String In dic_attar.keys()
    Str_result = String.format("{0}{2} : {3}{1}", Str_result, vbnewline,key,dic_attar(key))
  Next
  Str_result  = String.format("{0}{2} : {3}{1}", Str_result, vbnewline,"Selector",ui_tmp.Selector.Text)
  Str_result  = String.format("{0}{2} : {3}{1}", Str_result, vbnewline,"SelectorStrategy",ui_tmp.SelectorStrategy.ToString)
  Str_result  = String.format("{0}{2} : {3}{1}", Str_result, vbnewline,"ParentSelector",ui_tmp.Parent.Selector.Text)
  Str_result  = String.format("{0}{2} : {3}{1}", Str_result, vbnewline,"TopParent",ui_tmp.TopParent().Selector.Text)
  System.IO.File.WriteAllText(Str_fileName ,Str_result)
  Return Str_result
End Function
'-----------------------
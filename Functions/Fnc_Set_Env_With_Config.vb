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
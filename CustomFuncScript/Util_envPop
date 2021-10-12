out_envPop = Function(str_varName As String)
    Dim str_sep As String
	Dim str_val As String
	Dim str_Result As String
	Dim int_index As int32
		
	str_sep = Environment.GetEnvironmentVariable("envSep")
	If String.IsNullOrWhiteSpace(str_sep) Then 
		str_sep = ":"
		Environment.SetEnvironmentVariable("envSep",str_sep)
	End If 
	
	str_val = Environment.GetEnvironmentVariable(str_varName)
	int_index = str_val.IndexOf(str_sep)
	If int_index <> -1 Then
		str_Result = split(str_val,str_sep).Last.ToString
		Environment.SetEnvironmentVariable(str_varName,str_val.Substring(0,str_val.Count - str_Result.Count-str_sep.Count))
		Return str_Result
	Else 
		Environment.SetEnvironmentVariable(str_varName,"")
		Return str_val 
	End If 

End Function

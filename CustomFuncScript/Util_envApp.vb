out_envApp = Function(str_varName As String, str_Value As String)
    Dim str_sep As String
	Dim str_val As String
	
	str_sep = Environment.GetEnvironmentVariable("envSep")
	If String.IsNullOrWhiteSpace(str_sep) Then 
		str_sep = ":"
	End If 
	
	str_val = Environment.GetEnvironmentVariable(str_varName)
	If String.IsNullOrWhiteSpace(str_val) Then
		Environment.SetEnvironmentVariable(str_varName,str_Value)
	Else 
		Environment.SetEnvironmentVariable(str_varName, str_val + str_sep + str_Value)
	End If 
	
    Return str_Value
End Function

out_envShowAll = Function(str_filter As String)
Dim str_result As String = ""

For Each key As String In Environment.GetEnvironmentVariables.Keys
	If String.IsNullOrWhiteSpace(str_filter) OrElse key.Contains(str_filter)
		str_result = str_result + String.Format("Key :{1}{0}",vbNewLine, key)
		str_result = str_result + String.Format("Value :{1}{0}",vbNewLine, Environment.GetEnvironmentVariable(key))
	End If
Next
Return str_result 

End Function

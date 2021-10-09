out_envSet = Function(str_key As String, str_Value As String)
    Environment.SetEnvironmentVariable(str_key,str_Value)
    Return str_Value
End Function
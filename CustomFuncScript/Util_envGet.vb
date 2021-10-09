out_envGet = Function(str_key As String)
    If String.IsNullOrWhiteSpace(environment.GetEnvironmentVariable(str_key)) Then
            Throw New applicationException(String.Format("Environment에 {0}변수가 없습니다.",str_key))
        Else
            Return environment.GetEnvironmentVariable(str_key)
    End If
End Function
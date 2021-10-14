out_envEnter = Function(str_seqName As String)
	Dim str_sep As String  = ">"
	Dim str_RPAtree As String = Environment.GetEnvironmentVariable("[RPA]")
	Dim str_Now As String = now.ToString("yyyy/MM/dd HH:mm:ss") 
	
	' 값 설정
	If String.IsNullOrWhiteSpace(str_RPAtree) Then
		str_RPAtree = str_seqName 
	Else 
		str_RPAtree = str_RPAtree + str_sep + str_seqName
	End If 
	
	' Data 기록
	Environment.SetEnvironmentVariable("[RPA]", str_RPAtree)
	Environment.SetEnvironmentVariable("[RPA]>"+ str_RPAtree +"-StartTime" , str_Now )
	
	' Log에 넣을 값 반환
    Return String.Format("{1} 시작! {0}{1} 시작일시 : {2}", vbNewLine, str_seqName, str_Now)

End Function

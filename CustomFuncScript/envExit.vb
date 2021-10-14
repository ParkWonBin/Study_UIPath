out_envExit = Function()
    Dim str_sep As String = ">"
	Dim str_seqName As String = ""
	Dim str_RPAtree As String = Environment.GetEnvironmentVariable("[RPA]")
	Dim str_timeTaken As String = Environment.GetEnvironmentVariable("[RPA]>"+ str_RPAtree +"-StartTime")
	
	' [RPA]에서 seqName Pop
	If Not String.IsNullOrWhiteSpace(str_RPAtree)
		If str_RPAtree.IndexOf(str_sep) <> -1 Then
			str_seqName = split(str_RPAtree,str_sep).Last.ToString
			str_RPAtree = str_RPAtree.Substring(0, str_RPAtree.Count -str_sep.Count -str_seqName.Count)
		Else
			str_seqName = str_RPAtree
			str_RPAtree = ""
		End If 
	
		' Data 갱신
		Environment.SetEnvironmentVariable("[RPA]", str_RPAtree)
		
		'수행시간 계산
		If String.IsNullOrWhiteSpace(str_timeTaken) Then
			Return str_seqName + " 종료!"
		Else 
			str_timeTaken = now.Subtract( CDate(str_timeTaken) ).TotalSeconds.ToString("000")
			Return String.Format("{1} 종료!{0}{1} 수행시간 {2} 초 ",vbNewLine, str_seqName, str_timeTaken)
		End If
	Else
		Return "ERROR : envExit() >>> envEnter 의 값이 설정되지 않았습니다!"
	End If
End Function

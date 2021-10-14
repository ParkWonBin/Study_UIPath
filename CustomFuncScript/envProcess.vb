out_envProcess = Function(int_crtidex As int32, int_MaxCount As int32)
    Dim str_sep As String = ">"
	Dim str_seqName As String = ""
	Dim str_intFormat As String = ""
	Dim str_RPAtree As String = Environment.GetEnvironmentVariable("[RPA]")
	Dim str_timeTaken As String = Environment.GetEnvironmentVariable("[RPA]>"+ str_RPAtree +"-StartTime")
	
	If int_MaxCount > 100 Then
		str_intFormat = "000"
	Else If int_MaxCount > 10 Then
		str_intFormat = "00"
	Else 
		str_intFormat = "0"
	End If
	
	' [RPA]에서 seqName Pop
	If Not String.IsNullOrWhiteSpace(str_RPAtree)
		If str_RPAtree.IndexOf(str_sep) <> -1 Then
			str_seqName = split(str_RPAtree,str_sep).Last.ToString
			str_RPAtree = str_RPAtree.Substring(0, str_RPAtree.Count -str_sep.Count -str_seqName.Count)
		Else
			str_seqName = str_RPAtree
			str_RPAtree = ""
		End If 
	
		'수행시간 계산
		If String.IsNullOrWhiteSpace(str_timeTaken) Then
			Return String.Format("시작시간 Data가 없습니다! {0}([RPA]>{1}-StartTime)",vbnewline , str_RPAtree)
		Else 
			str_timeTaken = now.Subtract( CDate(str_timeTaken) ).TotalSeconds.ToString("000")
			Return String.Format("{0} 진행중... {1} / {2} | {3} 초 경과", str_seqName, int_crtidex.ToString(str_intFormat),int_MaxCount.ToString(str_intFormat),str_timeTaken )
		End If
	Else
		Return "ERROR : envProcess() >>> envEnter 의 값이 설정되지 않았습니다!"
	End If
End Function

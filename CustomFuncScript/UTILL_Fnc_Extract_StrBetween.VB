out_Fnc_Extract_StrBetween= Function(Str_source As String, Str_Start As String, Str_End As String) 
	Dim int_index As Integer
	
	' Str_Start 존재하는지 확인
	int_index = Str_Source.IndexOf(Str_Start)
	If int_index = -1 Then
		Return String.format("Str_Start = ''{0}''을(를) 찾지 못했습니다.",Str_Start)
	End If 
	
	' Str_Start 앞부분 버림
	Str_Source = Str_Source.substring(int_index+Str_Start.Length)
	
	' Str_End 존재하는지 확인
	int_index = Str_Source.IndexOf(Str_End)
	If int_index = -1 Then
		Return String.format("Str_End = ''{0}''을(를) 찾지 못했습니다.",Str_End) 
	End If 
	
	' Str_End 뒷부분 버림	
	Return Str_Source.substring(0,int_index).Trim
End Function


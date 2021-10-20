out_Bake_Config_AsTextMd= Function ( dic_config As dictionary(Of String, String) )
	Dim str_path As String = Environment.CurrentDirectory+"\Bake_Config_"+now.tostring("yyMMdd")+".md"
	Dim str_result As String
	Dim str_dic As String
	Dim value As String
	
	' 내용 생성
	str_dic = "New Dictionary(Of String, String) From {"
	
	Dim keys As String() = dic_config.keys().toarray
	system.array.sort(keys)
	
	For Each key As String In keys
		value = dic_config(key).ToString.Replace(Environment.NewLine," ").Replace(vbCr," ").Replace(vbLf," ").Replace(vbCrLf," ").Trim
		str_dic = str_dic + String.Format("{3}{1} {0}{4}{0} , {0}{5}{0} {2},","""","{","}",vbnewline,key,value)
	Next
	str_dic = str_dic.Substring(0,str_dic.Length-1) + vblf + "}"
		
	' 파일 저장
	Try 
		If file.Exists(str_path)
			file.Delete(str_path)
		End If
 		file.WriteAllText(str_path, str_dic)
		str_result = "저장이 완료되었습니다."+ vbnewline + str_path
	Catch exception As exception
		str_result = "저장 실패 " + vbnewline + exception.Message
	End Try 
	
	' 수행 결과 반환
	Return str_result
End Function 

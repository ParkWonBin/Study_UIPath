out_bake_Config = Function(dic_config As Dictionary(Of String, String), str_path As String)
	Dim str_dic As String
	Dim str_result As String
	Dim value As String
	
	' 내용 생성
	str_dic = "New Dictionary(Of String, String) From {"
	For Each key As String In dic_config.Keys
		value = dic_config(key).ToString.Replace(Environment.NewLine," ").Replace(vbCr," ").Replace(vbLf," ").Replace(vbCrLf," ").Trim
		str_dic = str_dic + String.Format("{3}{1} {0}{4}{0} , {0}{5}{0} {2},","""","{","}",vbnewline,key,value)
	Next
	str_dic = str_dic.Substring(0,str_dic.Length-1) + vblf + "}"
		
	' 절대경로 확인
	If Not str_path.Contains(":") ' 절대경로는 "C:" 포함되있으므로 :로 판단
		str_path = environment.CurrentDirectory+"\"+str_path
	End If
	
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

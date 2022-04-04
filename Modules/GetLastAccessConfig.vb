' out_ConfigPath as String
Dim Str_Directory As String = Environment.CurrentDirectory
Dim Str_name As String = "Config"
Dim Str_ext As String = ".xlsx"
Dim StrArr_Files As String()
Dim Str_LassAccessFile As String

' 프로젝트 경로에서 Config 선택
StrArr_Files  = directory.GetFiles(Str_Directory).Where(Function(x) split(x,"\").last.tostring.Contains(Str_name) AndAlso split(x,"\").last.tostring.Contains(Str_ext)).toarray

' 가장 최신 Config 사용
If StrArr_Files.count > 0 Then
	Str_LassAccessFile = StrArr_Files.first.tostring
	For Each filePath As String In StrArr_Files
		If file.GetLastAccessTime(Str_LassAccessFile).tostring("yyMMddHHmmss") < file.GetLastAccessTime(filePath).tostring("yyMMddHHmmss") Then
			Str_LassAccessFile = filePath
		End If 
	Next 
Else 
	Str_LassAccessFile = String.empty
End If

out_ConfigPath = Str_LassAccessFile

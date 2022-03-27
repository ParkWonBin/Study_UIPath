

### System.IO
```vb
' 경로 입력 관련
System.IO.Path.Combine(A,B,...) ' 입력 받은 경로를 모두 경로 구분자로 합친다. 
System.IO.Path.GetDirectoryName(Str_Path) : '입력 받은 경로의 상위 폴더의 절대경로를 반환한다.

' 경로 유효성 관련
System.IO.File.Exists(Str_Path) ' 해당 경로에 파일이 존재하면 True
System.IO.Directory.Exists(Str_Path) '해당 경로에 폴더가 존제하면 True
System.IO.File.GetCreationTimeUtc(Str_Path) '파일 생성일시 Date로 반환 
System.IO.File.GetLastAccessTimeUtc(Str_Path) ' 파일 마지막 접근일시 Date로 반환
System.IO.File.GetLastWriteTimeUtc(Str_Path) '파일 마지막 수정일시 Date로 반환  

' 경로 내 검색 관련
System.IO.Directory.GetDirectories(Str_Path) ' 해당 경로에 위치한 폴더의 절대경로를 String_Array로 반환한다.
System.IO.Directory.GetFiles(Str_Path) ' 해당 경로에 위치한 파일의 절대경로를 String_Array로 반환한다. 

' 파일 작성 및 삭제 관련
System.IO.Directory.CreateDirectory(Str_Path) ' 폴더 생성
System.IO.File.ReadAllText(Str_Path, System.Text.Encoding.UTF8) ' 파일 읽기. Encoding 설정 가능
System.IO.File.WriteAllText(Str_Path, Str_Content, System.Text.Encoding.UTF8) '파일 생성. Encoding 설정 가능
System.IO.File.AppendAllText(Str_Path, Str_Content, System.Text.Encoding.UTF8) '파일 이어서 쓰기. Encoding 설정 가능
System.IO.File.Copy(Str_Sorce,Str_Dest,Bln_overwite) ' 파일 복제. 덮어쓰기 여부 선택
System.IO.File.Delete(Str_Path) ' 파일 삭제
```


### System.Environment 관련
```vb
System.Environment.CurrentDirectory ' 현제 프로젝트 경로 반환
System.Environment.GetEnvironmentVariable(Str_KeyName) ' 환경변수에 저장된 값 반환, key가 없을 경우 "" 반환
System.Environment.SetEnvironmentVariable(Str_KeyName, Str_Value) ' 해당 Key에 value값을 적용한다.
```

### process
```vb
System.Diagnostics.Process.GetProcessesByName(Str_ProcessName) ' ProcessName으로 프로세스 검색, Enum
' 사용 예시 : Kill Process By ProcessName
For Each p As System.Diagnostics.Process In System.Diagnostics.Process.GetProcessesByName(Str_ProcessName)
 p.Kill()
Next
```

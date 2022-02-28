# Custom Function
## 자주 쓰는 함수


### System.IO 관련
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

'--- 사용자 정의 함수
Dim Fnc_Get_All_Files As System.Func(Of String, String()) = Function(str_path As String) As String()
	Dim list_str_dir As New System.Collections.Generic.List(Of String) 
	Dim list_str_file As New System.Collections.Generic.List(Of String)
	
	list_str_dir.Add(str_path)
	While list_str_dir.Count <> 0
		str_path = list_str_dir.Last
		list_str_dir.RemoveAt(list_str_dir.Count -1)
		list_str_dir.AddRange(System.IO.Directory.GetDirectories(str_path))
		list_str_file.AddRange(System.IO.Directory.GetFiles(str_path))
	End While
	Return list_str_file
End Function
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

' Kill All Process
For Each p As System.Diagnostics.Process In System.Diagnostics.Process.GetProcessesByName(Str_ProcessName)
 p.Kill()
Next

```

## 함수 정의 방법

### 1. invoke code를 통해 함수를 정의하고 out으로 반환하여 사용한다. 
- Edit Arguments 통해서 정의한 함수를 밖으로 출력한다.
- 자료형은 System.Func<T,...TResult> 로 되어있는 것으로 한다.
- T의 경우 인수로 입력받을 변수의 Type을 지정하는 설정이며, 마지막 TResult는 함수의 반환값의 Type이다.

```vb
Fnc_test = Function(str_test As String) As String()
 Return "input : " & str_test
End Function
```


### 2. invoke code 내에서 정의한 함수를 해당 invoke code 내에서 재사용 한다.
```vb
Dim Fnc_test As System.Func(Of String, String()) = Function(str_test As String) As String()
 Return "input : " & str_test
End Function

console.WriteLine( Fnc_test("Test") )
```

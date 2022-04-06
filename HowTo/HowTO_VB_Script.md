#### 초기화 관련
```vb
New String(){"1","2"} 'string array 생성및 할당   
New String(n){} '원소가 n개인 string array 생성

new List(of int32)  
new List(of string)   
new List(of string)(new string(){"가","나","다","라"})   
new List(of String) from {{"보험"},{"세금"}}   

New Dictionary(of string, int32)   
New Dictionary(of string,int32) from {{"red",50},{"yellow",10},{"green",80}}   
New Dictionary(Of String, string()) # 문자열 배열    
New Dictionary(of string, object) from {{"test1","50"},{"test2","10"},{"test3","80"}}   

'비밀번호 관련
(new System.Net.NetworkCredential("",Str)).SecurePassword 'secureString 반환
(new System.Net.NetworkCredential("",Str)).Password 'String 반환
new System.Net.NetworkCredential(string.Empty, secureStr).Password 'SecureStr -> str
``` 

### String 관련
```vb
'정수 표시 형식
cint("1").ToString("0000") '= 0001  

'날짜 표시 형식
now.ToString("yyyy_MM_dd")

'한글 날짜 표시
DateTime.Today.ToString("dddd", CultureInfo.CreateSpecificCulture("ko-KR"))  #목요일
DateTime.Today.ToString("ddd", CultureInfo.CreateSpecificCulture("ko-KR"))   #목
Date.ParseExact("20210212", "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)  
```
```yaml
한글 날짜 표시 : 
  - 초기 설정 : import 패널에서 System.Globalization 추가 [CultureInfo 객체 사용을 위함입니다.]
  - 호출 방법 : DateTime.Today.ToString("dddd", CultureInfo.CreateSpecificCulture("ko-KR"))  #목요일
  - 참고 자료 : https://m.blog.naver.com/PostView.naver?isHttpsRedirect=true&blogId=elduque&logNo=120096308343
```


#### String To Charactor
```vb
Asc("A") '= 65 
Chr(65) '= "A"
' ShortCode : 문자열 - 아스키코드 번호
join(str_tmp.ToCharArray.Select(function(x) string.Format("{0} : {1}",x,asc(x).ToString) ).ToArray, vbNewLine)

' CSV 열 구분 : chr(44) = ','
' CSV 행 구분 : chr(13)+chr(10) = \r\n
' 엑셀 셀 내부 줄바꿈 : chr(10) 
```

### System.IO
```vb
' 경로 입력 관련
System.IO.Path.Combine(A,B,...) ' 경로 합쳐서 반환. 
' A,B,C... 의 마지막 문자에 경로 구분자(\)가 있든 없든 정상적으로 경로 합쳐서 반환하기 떄문에 많이 씁니다.

' 경로 내 검색 관련
System.IO.Directory.GetFiles(Str_Path)       ' 해당 경로에 위치한 파일의 절대경로를 String_Array로 반환한다. 
System.IO.Directory.GetDirectories(Str_Path) ' 해당 경로에 위치한 폴더의 절대경로를 String_Array로 반환한다.

' 파일 작성 및 삭제 관련
System.IO.Directory.CreateDirectory(Str_Path) ' 폴더 생성
System.IO.File.ReadAllText(Str_Path, System.Text.Encoding.UTF8) ' 파일 읽기. Encoding 설정 가능
System.IO.File.WriteAllText(Str_Path, Str_Content, System.Text.Encoding.UTF8) '파일 생성. Encoding 설정 가능
System.IO.File.AppendAllText(Str_Path, Str_Content, System.Text.Encoding.UTF8) '파일 이어서 쓰기. Encoding 설정 가능
System.IO.File.Copy(Str_Sorce,Str_Dest,Bln_overwite) ' 파일 복제. 덮어쓰기 여부 선택
System.IO.File.Delete(Str_Path) ' 파일 삭제

' 파일 정보 관련 (Static 함수)
System.IO.Directory.Exists(Str_Path)        ' 폴더 존제하면 True
System.IO.File.Exists(Str_Path)             ' 파일 존재하면 True
System.IO.file.GetCreationTime(Str_Path)    ' 파일 최초 생성 시간
System.IO.file.GetLastWriteTime(Str_Path)   ' 파일 최종 수정 시간
System.IO.file.GetLastAccessTime(Str_Path)  ' 파일 최종 접근 시간
System.IO.file.GetAttributes(Str_AttarName) ' 파일 속성 확인
System.IO.Path.GetDirectoryName(Str_Path) : ' 해당 경로의 상위 폴더 경로를 반환한다.
System.IO.Path.GetFileName(Str_Path)        ' 확장자 포함 파일명 ex) "Main.xaml"
System.IO.Path.GetExtension(Str_Path)       ' 확장자 반환 ex) ".xaml"
System.IO.Path.GetFileNameWithoutExtension(Str_Path) ' 확장자 미포함 파일명 ex) "Main"

' 파일 정보 관련 (객체 매소드 사용)
(new System.IO.DirectoryInfo(Str_Path)).Exists    ' 폴더 존재하면 True
(new System.IO.FileInfo(Str_Path)).Exists         ' 파일 존재하면 True
(new System.IO.FileInfo(Str_Path)).CreationTime   ' 파일 최초 생성 시간
(new System.IO.FileInfo(Str_Path)).LastWriteTime  ' 파일 최종 수정 시간
(new System.IO.FileInfo(Str_Path)).LastAccessTime ' 파일 최종 접근 시간
(new System.IO.FileInfo(Str_Path)).Attributes(Str_AttarName) ' 파일 속성 확인
(new System.IO.FileInfo(Str_Path)).DirectoryName  ' 해당 경로의 상위 폴더 경로를 반환한다.
(new System.IO.FileInfo(Str_Path)).Name           ' 확장자 포함 파일명 ex) "Main.xaml"
(new System.IO.FileInfo(Str_Path)).Extension      ' 확장자 반환 ex) ".xaml"
```

### System.Environment 관련
```vb
System.Environment.CurrentDirectory ' 현제 프로젝트 경로 반환
System.Environment.GetEnvironmentVariable(Str_KeyName) ' 환경변수에 저장된 값 반환, key가 없을 경우 "" 반환
System.Environment.SetEnvironmentVariable(Str_KeyName, Str_Value) ' 해당 Key에 value값을 적용한다.
```

### TimeSpan
```vb
TimeSpan.FromMilliseconds(int_delayTime) ' 딜레이 시간 넣을 떄 사용    
System.Threading.Thread.Sleep(2000)
```

### process
```vb
System.Diagnostics.Process.GetProcessesByName(Str_ProcessName) ' ProcessName으로 프로세스 검색, Enum
' 사용 예시 : Kill Process By ProcessName
For Each p As System.Diagnostics.Process In System.Diagnostics.Process.GetProcessesByName(Str_ProcessName)
 p.Kill()
Next
```


#### Uipath Invoke Code 사용자 정의 함수
invoke code 안에서 함수 정의하고 재호출 하는 것도 가능하다.
```vb
Dim test As System.Func(Of String,String)  = Function (str_tmp As String) As String
	console.WriteLine(str_tmp)
	Return str_tmp
End Function
test(test("123"))
'123
'123
```


##### VB 문법 For Each
```vb

For int_i As Integer = 0 To 5
	console.writeline(int_i.tostring)
Next
' 0 1 2 3 4 5

For int_j As Integer = 5 To 0 Step -1
	console.writeline(int_j.tostring)
Next
' 5 4 3 2 1 0

For Each  int_k   As Integer In  {1, 2, 3, 4, 5}
	console.writeline(int_k.tostring)
Next
' 0 1 2 3 4 5
```

##### DataRow 다룰 때 주의사항
```vb
For Each row as Data.DataRow in DT_tmp
    For Each item as Object in row.ItemArray
    ' item 을 item as String 으로 쓰면 Null 들어간 Row 처리할 떄 에러 발생함.
	' item 은 꼭 Object로 선언하고, 호출할 떄 ToString 처리하는 것이 안전함.
        Console.WriteLine( item.ToString ) 
    Next
    ' 마찬가지로 Linq로 DataRow를 다룰 경우, Object에서 명시적으로 형변환을 해줘야한다.
    StrArr_datarow = row.ItemArray.Select(Function(x) x.Tostring).ToArray
Next
* Row를 ItemArray로 바꿀 때, 해당 변수를 받을 때는 꼭 Object로 받고 호출시 ToString을 하자.
* Row를 ItemArray로 바꾸는 과정에서 Null이 포함된 row에서 item을 String으로 받으면 에러가 발생한다. (Null을 String으로 형변환 못한다는 오류)
* 따라서 Row.ItemArray를 쓸 일이 있을 경우 Object()로 받거나, Select를 통해 ToString을 직접 시켜주는 게 좋다.
```

#### 엑셀 읽어서 해더명에 공백 제거
```vb
For Each col As System.Data.DataColumn in dt_tmp.Columns
    col.ColumnName = col.ColumnName.Replace(" ","")    
Next
```

## Datatable 
```vb
' Copy는 열 이름에 상관 없이 값을 복사 붙여넣기 한다.
DT_test = DT_tmp.Copy()

' Clone은 데이터는 복사하지 않고 Columns만 복사해서 넣는다.
DT_test = DT_tmp.Clone()

' DataTable Row 순서 반대로 바꾸기
DT_tmp = DT_tmp.AsEnumerable.Reverse().CopyToDataTable

' DT DataRow 필터링
dt_tmp.AsEnumerable.Skip(int_n).Take(int_m).CopyToDataTabl

' DataTable 짝수 행만 선택
DT_tmp = DT_tmp.AsEnumerable.Where(Function(x,i) i%2==1).CopyToDataTable

'DT DataColumn 필터링
dt_tmp = dt_tmp.DefaultView.ToTable(false, {"a","c","e"} ) '열 필터링
' DefaultView : 첫번쨰 인자는 false로 해야한다. True로 할 경우 오류 발생(distinct 속성)
if : drArr.count > 0  Then
    ' CopyToDataTable은 count가 0일 때 에러가 발생하기 때문에 DataRow[] 를 사용하여 예외처리
    dt_tmp  = drArr_tmp.CopyToDataTable.DefaultView.ToTable(False, {"a","c","e"}) '열 필터링
End if 

' 특정 열 Data String Array로 뽑기
StrArr_ColData = DT_tmp.AsEnumerable.Select(Function(row) row.Field(of string)("ColName").ToString).ToArray
StrArr_ColData = DT_tmp.AsEumnerable.Select(Function(row) row.item("ColName").ToString).ToArray

For Each row  As System.Data.DataRow In dt_tmp.AsEnumerable
    row.SetField("ColName","Value")
    row("ColName") = "Value"
Next
' argument dt_tmp는 in으로 주어도 정상적으로 수정됨, 

'### convert dt to Dictionary
DT_tmp.AsEnumerable.ToDictionary(Of String, Object)(Function (row) row("key").toString, Function (row) row("value").toString)

```

#### DataSet
```vb
Dim Ds_tmp As New System.Data.DataSet
Dim Dt_tmp As New System.Data.DataTable("TableName")
' table 추가
Ds_tmp.Tables.Add(Dt_tmp)
Ds_tmp.Tables.Add("TableName2")
"c1,c2,c3".split(","c).Select(Function(x) Ds_tmp.Tables("TableName2").Columns.Add(x System.Type.GetType("System.String") ))
```

#### Dictionary
```vb

dic_CSS = New Dictionary(Of String, String) From {
{ "table" , "color: black ; text-align: center; border-collapse: collapse; margin-top: 10px;" },
{ "tr" , "" },
{ "th" , "background-color:#d9d9d9; border:1px solid black; font-family:맑은 고딕; font-size:10pt; padding:4px; height:34px;" },
{ "td" , "background-color:#ffffff; border:1px solid black; font-family:맑은 고딕; font-size:10pt; padding:4px; height:34px;" },
{ "width_col0" , "100" },
{ "width_col1" , "150" },
{ "width_col2" , "50" }
}
```

### 함수 만들기 요렁

```yaml
Uipath 에서 사용자 정의 함수 만들기 요령 :
- Inovke Code를 통해 VB로 정의한 함수를 Out 변수로 뺴낸다. 
- 함수의 자료형은 System.Func<T,...TResult> 
- 인수 설명 : 
  - T : 인수로 입력받을 변수의 Type
  - TResult : 함수의 반환값의 Type
```

```vb
Fnc_test = Function(str_test As String) As String
 Return "input : " & str_test
End Function
```

### 람다 함수식에 변수 넣어주기
```vb
' 그냥 람다식에 () 치고 바로 뒤에 (인수) 넣어주면 됨.
Console.WriteLine( ( (Function(num As Integer) num + 0)(5) ).ToString ) '=> 6
```

## 내 함수 목록

```vb
Dim Fnc_Get_All_Files As System.Func(Of String, String()) = Function(str_path As String) As String
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


##### Print Dictionary / Bake Config.log, Config.excel
```vb
Dim in_Dic_Config As New Dictionary(Of String,String)
Dim Str_Config As String

Str_Config = 
String.Format("{1}{0}{3}{0}{2}",vbNewLine,"New Dictionary(Of String,String) From {","}", Join( in_Dic_Config.Keys.Select(Function(key) String.Format("{0} {2}{3}{2} , {2}{4}{2} {1}", "{","}", chr(34), key, in_Dic_Config(key).Replace(vbNewLine," ").Replace(chr(10)," ").Replace(chr(34),"'") ) ).ToArray, ","+vbNewLine) )

'chr(10) : 엑셀 줄바꿈 문자열
'chr(34) : 쌍따옴표 "
file.WriteAllText("Config.md", Str_Config)
```

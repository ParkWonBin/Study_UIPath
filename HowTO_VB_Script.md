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

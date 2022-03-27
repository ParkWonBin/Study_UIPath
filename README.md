#### Bake Config
```vb
' Write Text File
' "Config.log"
String.Format("{1}{0}{3}{0}{2}",vbNewLine,"New Dictionary(Of String,String) From {","}", Join( Dic_Config.Keys.Select(Function(key) String.Format("{0} {2}{3}{2} , {2}{4}{2} {1}", "{","}", chr(34), key, System.Convert.ToString(Dic_Config(key)).Replace(vbNewLine," ").Replace(chr(10)," ").Replace(chr(34),"'") ) ).ToArray, ","+vbNewLine) )
```

#### Dictionary 필터링
```vb
' 선언과 초기화 동시에 진행
Dim dic_config As New Dictionary(Of String, String) From{ {"a1","11"}, {"a2","12"}, {"b2","22"} }

' ToDictionary 사용법
dic_config = dic_config.Keys.Where(Function(key) key.Contains("a"))
             .ToDictionary(Function(key) key, Function(key) dic_config(key))
' 1. 호출 전 : String Array 형태로 가공한다. * pair 형태 아님!
' 2. 인수 값 : 인자는 ,를 구분자로 하여 key값과 value 같을 정의할 function을 2개 넣어주어야 한다.
	
' 출력 예시
For Each k As String In dic_config.Keys
	console.WriteLine(string.Format("{0} : {1}",k , dic_config(k)))
Next 
```


##### BuildDataTable by Sting
Uipath Debug에서 Immediate로 멈춰두고 아래 코드수행하면,  
Local에서 값이 바뀐다. (메모리에 저장된 dt 위치의 값을 직접 수정하는 명령 포함)  
에러났을 때 끄지 않고 DT 값을 수정 후 이어서 Retry 할 수 있게 한다.

```vb
' Variables 패널 설정
dt_tmp As System.Data.DataTable
ArrStr_colName As String()
ArrArrStr_data As String()()

'Assign
dt_tmp = new DataTable()
ArrStr_colName = "col1|col2|col3".Split("|"c).Select(function(x) x.trim).ToArray
ArrArrStr_data = "00|01\10|11|12|13|\20|21|22|23|24|25|26|27".Split("\"c).Select(function(tr) tr.Split("|"c).Select(function(td) td.trim).ToArray)

'Log Message - 입력 데이터 확인
string.Format("입력 데이터 확인{0}{1}{0}{2}",vbNewLine,join(ArrStr_colName," | "),  join(ArrArrStr_data.Select(function(tr) join(tr.Select(function(td) td.Trim).ToArray, " | ")).ToArray,vbNewLine) )

'Log Message - 데이터 적용
string.Format("BuildDataTable{0} {0}Dt_tmp <- Add Columns :{0}{1}{0} {0}Dt_tmp <- Add Data :{0}{2}{0}",vbNewLine,join(ArrStr_colName.Select(function(colName) dt_tmp.Columns.Add(colName.Trim).ToString).ToArray, " | "),  join(ArrArrStr_data.Select(function(tr) join( dt_tmp.Rows.Add(tr.Take(dt_tmp.Columns.Count).ToArray).itemArray.Select(function(td) td.ToString).ToArray , " | ")).ToArray , vbNewLine))

'원리 설명
'dataTable.columns.Add() 와 dataTable.Rows.Add()는 각각 하고 입력받은 인수(String, DataRow)를 그대로 Return하는 함수다.
'select로 dt를 수정하는 함수를 호출하고, 리턴값 잘 조작하여 최종적으로 String 형태를 만들면 LogMassage에서 해당 code를 사용할 수 있다.

'Log Message - Col만 추가
join("col1|col2|col3".Split("|"c).Select(function(x) if(dt_tmp.Columns.Contains(x.trim), x.Trim+" (Skip-중복)", dt_tmp.Columns.Add(x.Trim).ToString)).ToArray," | ")
' 중복된 이름의 Column을 추가하려고 할 경우 오류 발생

'Log Message - Data만 추가
join("00|01\10|11|12|13|\20|21|22|23|24|25|26|27".Split("\"c).Select(function(tr) tr.Split("|"c).Select(function(td) td.trim).ToArray ).Select(function(tr) join( dt_tmp.Rows.Add(tr.Take(dt_tmp.Columns.Count).ToArray).itemArray.Select(function(td) td.ToString).ToArray , " | ")).ToArray , vbNewLine)
' newRow의 item.count가 col.count보다 작을 때는, 부족하면 맨 null을 체워넣는다.
' newRow의 item.count가 col.count보다 클 때는, Take를 통해 col.count 개수만큼만 사용하고 초과된 item은 버린다.

' 요령
' 1. 반복문으로 수행할 함수는 Select를 통해 호출한다.
' 2. object의 경우 {object}.ToString 을 사용하고하여 String으로 객체로 만들어 작업한다. (Nothing, null도 ""객체로 만들어준다.)
' 3. IEnumerable의 경우 {enumarable}.Select(Function(x) x.ToString).Array 를 통해 String Arrray 형태로 만들어 작업한다.
' 4. String Array의 경우 Strings.Join() 함수를 통해 String으로 만든다.
' 5. 개별적으로 동작하는 함수를 String을 반환하도록 마들었다면.  String.Format() 함수를 통해 One-Line으로 병합한다.
' 6. String.Format에 인수가 입력되는 과정에서 위에서 정의한 함수가 1번씩 실행된다. 굳이 {0}등으로 내용을 표시를 하지 않아도, 입력받은 인자 순서대로 Code가 동작한다.
' 7. 위 방법의 한계는 "값을 산출하는 함수"만 호출할 수 있다는 것이다. 값을 산출하지 않는 함수는 function을 인자로 받는 함수를 통해 호출할 수 없다.

'Make DummyDataTable
dt_tmp = new DataTable()
'log Message - Add Column
join("|||||".Split("|"c).Select(function(x) if(dt_tmp.Columns.Contains(x.trim), x.Trim+" (Skip-중복)", dt_tmp.Columns.Add(x.Trim).ToString)).ToArray," | ")
'log Message - Add Data
join("00|01\10|\20|21\".Split("\"c).Select(function(tr) tr.Split("|"c).Select(function(td) td.trim).ToArray ).Select(function(tr) join( dt_tmp.Rows.Add(tr.Take(dt_tmp.Columns.Count).ToArray).itemArray.Select(function(td) td.ToString).ToArray , " | ")).ToArray , vbNewLine)
 
```
## 엑셀 다루기 팁


## StrArr to DataTable Column
```vb
DT_tmp = new DataTable
StrArr_data = "제목|data1|data2|data3".Split("|"c)

# 1. Column 추가
dT_tmp.Columns.Add( StrArr_data.First )

# 2. Data 추가
join( StrArr_data.Skip(1).Select(function(x) join( DT_tmp.Rows.Add({x}.take(1).ToArray ).ItemArray.select(function(y) y.ToString).ToArray , " | ") ).ToArray, vbNewLine )

# 3. 합본
string.Format("{1}{0}{2}",vbNewLine,dT_tmp.Columns.Add( StrArr_data.First ) , join(  StrArr_data.Skip(1).Select(function(x) join( DT_tmp.Rows.Add({x}.take(1).ToArray ).ItemArray.select(function(y) y.ToString).ToArray , " | ") ).ToArray, vbNewLine ) )

```

## DataSet에 Table 넣고 Data 초기화
```vb
DS_RPA = new dataset()
StrArr_cols = "DT1|col1|col2".Split("|"c)
StrArr_data = "data1|data2".Split("|"c)

# 1. Table 추가
DS_RPA.Tables.Add(StrArr_cols.First).TableName

# 2. Column 추가
join( StrArr_cols.Skip(1).Select(function(x) DS_RPA.Tables(StrArr_cols.First).Columns.Add( x.ToString.Trim ).ColumnName).ToArray , " | " )

# 3. Data 추가
join(DS_RPA.Tables(StrArr_cols.First).Rows.Add(StrArr_data).itemArray.select(function(x) x.ToString).ToArray , "|")

# 4. 합본
string.Format("{1}{0}{2}{0}{3}",vbNewLine, DS_RPA.Tables.Add(StrArr_cols.First).TableName , join( StrArr_cols.Skip(1).Select(function(x) DS_RPA.Tables(StrArr_cols.First).Columns.Add( x.ToString.Trim ).ColumnName).ToArray , " | " ) , join(DS_RPA.Tables(StrArr_cols.First).Rows.Add(StrArr_data).itemArray.select(function(x) x.ToString).ToArray , "|"))

```

## 특정일로부터 n개 날짜 선택하기
```vb
Str_tmp = "2021-11-21"
int_cnt = 50
join( Enumerable.Range(cint(DateTime.Parse(Str_tmp).ToOADate), int_cnt).Select(function(x) datetime.FromOADate(x).ToString("yyyy-MM-dd") ).ToArray ,vbNewLine)
```





### 셀렉터 잡을 때 팁
```vb
' Indicate To screen > ui_tmp
' Log Message
String.Format("{1} : {2}{0}{3}",vbNewLine,"Selector",ui_tmp.Selector.ToString, Join( ui_tmp.GetNodeAttributes(False).Keys.Select(Function(key) String.format("{1} : {2}", vbnewline, key, ui_tmp.GetNodeAttributes(False)(key))).ToArray, vbNewLine) )

' Assign : ui_tmp = ui_tmp.parent
' Log Message : 상동

' Assign : ui_tmp = ui_tmp.parent
' Log Message : 상동
```



##### Excel index2ColName
```vb
int_colIndex As String
' 엑셀 열 시작 = 0, 끝 = 16383

'Excel_Convert_index2ColName : 
if(int_colIndex=0,"A", join( Enumerable.Range(0, CInt(math.Ceiling(math.log(1+int_colIndex,26))) ).Select(Function(x) chr( if(x=0,65,64)+cint(((int_colIndex\cint(math.Pow(26,x)) ) mod 26) )).ToString ).reverse.ToArray, string.Empty))
 
```

##### UiElement 출력
```vb
Dim ui_tmp As Uipath.Core.UiElement ' [Indicate On Screen] Or [Find Element] 통해서 ui_tmp 초기화

' 셀렉터 및 Attribute 모두 출력
string.Format("{1} : {2}{0}{3}",vbNewLine,"Selector",ui_tmp.Selector.ToString, join( ui_tmp.GetNodeAttributes(False).Keys.Select(Function(key) String.format("{1} : {2}", vbnewline, key, ui_tmp.GetNodeAttributes(False)(key))).ToArray, vbNewLine) )
 
```

##### Xaml에서 사용된 모든 Key값 출력

```vb
Dim Str_ReadXamlFile As String
Dim StrArr_UsedKeys As String()
Dim StrArr_ShouldAdd As String()
Dim in_Dic_Config As Dictionary(Of String,String)

' Read Text File : Main.Xaml => Str_ReadXamlFile

' Xaml에서 사용된 모든 key 선택 (중복제거, 오름차순)
StrArr_UsedKeys = 
split( Str_ReadXamlFile.Replace(vbNewLine,"").Replace(" ",""), "onfig(").skip(1).Select(Function(x) if( x.IndexOf(")") = -1, "", x.Substring(0,x.IndexOf(")")) ) ).Distinct.OrderBy(function(x) x.ToString).select(function(x) x.replace("""","")).ToArray

' Xaml의 key 중 Config에 누락된 key 것만 선택
StrArr_ShouldAdd = 
StrArr_UsedKeys.Where(function(x) not in_Dic_Config.Keys.Contains(x) ).ToArray

' 누락된 key만 Dictionary에 바로 넣을 수 있는 형태로 출력
join( StrArr_ShouldAdd.Select(function(x) string.Format("{1} {0}{3}{0} , {0}dummy{0}  {2}", chr(34),"{","}",x.ToString) ).ToArray, ","+vbNewLine)

```
##### Bake DT as Log Text
```vb
' Assign
dt_temp = new DataTable()

' LogMassege -Add Column
join("A|B|C|D".Split("|"c).Select(function(x) if(dt_tmp.Columns.Contains(x.trim), x.Trim+" (Skip-중복)", dt_tmp.Columns.Add(x.Trim).ToString)).ToArray," | ")

' LogMassege - Add Data
join("a|2|0\a|2|1\a|2|0\b|2|0\b|2|0".Split("\"c).Select(function(tr) tr.Split("|"c).Select(function(td) td.trim).ToArray ).Select(function(tr) join( dt_tmp.Rows.Add(tr.Take(dt_tmp.Columns.Count).ToArray).itemArray.Select(function(td) td.ToString).ToArray , " | ")).ToArray , vbNewLine)

' LogMassege - print
string.Format("Index{1}{2}{0}{3}",vbNewLine,vbTab, join(Enumerable.Range(0,DT_tmp.Columns.Count).Select(function(x) DT_tmp.Columns.Item(x).ColumnName).ToArray," | "), join(Enumerable.Range(0,DT_tmp.Rows.Count).Select(function(tr) string.Format("{1}{0}{2}", vbTab, tr.ToString("000"),  join( DT_tmp.Rows(tr).ItemArray.Select(function(td) Convert.ToString(td)).ToArray, " | ") ) ).ToArray,vbNewLine))
```

##### Bake DT as HTML with CSS
```vb
Dim dic_CSS As Dictionary(Of String, String) 
Dim Str_HTML As String

dic_CSS = New Dictionary(Of String, String) From {
{ "table" , "color: black ; text-align: center; border-collapse: collapse; margin-top: 10px;" },
{ "tr" , "" },
{ "th" , "background-color:#d9d9d9; border:1px solid black; font-family:맑은 고딕; font-size:10pt; padding:4px; height:34px;" },
{ "td" , "background-color:#ffffff; border:1px solid black; font-family:맑은 고딕; font-size:10pt; padding:4px; height:34px;" },
{ "width_col0" , "100" },
{ "width_col1" , "150" },
{ "width_col2" , "50" }
}
' width가 너무 좁거나, width가 정의되지 않은 column은 "HTML 기본 width"로 설정됩니다.
' key로 ( "width_col" + index.Tostring ) 가 존재할 경우 해당 순서의 열에 width 설정을 함
' if(dic_CSS.Keys.Contains("width_col"+x.ToString), string.Format("width : {0}px;", dic_CSS("width_col"+x.ToString).Trim ) , string.Empty )

Str_HTML = ' 전체 Table 출력
String.Format("<table style=' {1} '> {0} {2} {0} {3} {0} </table>",vbNewLine,dic_CSS("table"),String.Format("<tr style=' {1} '> {0} {2} {0} </tr>",vbNewLine, dic_CSS("tr"), Join( Enumerable.Range(0,DT_tmp.Columns.Count).Select(Function(x) String.Format("<th style=' {0} {1} '> {2} </th>", dic_CSS("th"), If(dic_CSS.Keys.Contains("width_col"+x.ToString), String.Format("width : {0}px;", dic_CSS("width_col"+x.ToString).Trim ) , String.Empty ), DT_tmp.Columns.Item(x).ColumnName ) ).ToArray, vbNewLine) ),Join( DT_tmp.AsEnumerable.Select( Function(row) String.Format("<tr style=' {1} '> {0} {2} {0} </tr>",vbNewLine,dic_CSS("tr"), Join( Enumerable.Range(0,DT_tmp.Columns.Count).Select(Function(x) String.Format("<td style=' {0} {1} '> {2} </td>", dic_CSS("td"), If(dic_CSS.Keys.Contains("width_col"+x.ToString) , String.Format("width : {0}px;", dic_CSS("width_col"+x.ToString).Trim) ,string.Empty), row.Item(x).ToString ) ).ToArray, vbNewLine ) ) ).ToArray, vbNewLine))

' Column만 출력
string.Format("<tr style=' {1} '>{0} {2} </tr>{0}",vbNewLine, dic_CSS("tr"), join( Enumerable.Range(0,DT_tmp.Columns.Count).Select(function(x) string.Format("<th style=' {1} {3} '> {2} </th>{0}",vbNewLine, dic_CSS("th"), DT_tmp.Columns.Item(x).ColumnName, if(dic_CSS.Keys.Contains("width_col"+x.ToString), string.Format("width : {0}px;", dic_CSS("width_col"+x.ToString).Trim ) , string.Empty ) ) ).ToArray ) )

' Data만 출력
Join( DT_tmp.AsEnumerable.Select( Function(row) String.Format("<tr style=' {1} '>{0} {2} {0}</tr>{0}",vbNewLine,dic_CSS("tr"), Join( Enumerable.Range(0,DT_tmp.Columns.Count).Select(Function(x) String.Format("<td style=' {0} {2} '> {1} </td>",dic_CSS("td"), row.Item(x).ToString, if(dic_CSS.Keys.Contains("width_col"+x.ToString) , string.Format("width : {0}px;", dic_CSS("width_col"+x.ToString).Trim) ,"")) ).ToArray, vbNewLine ) ) ).ToArray, vbNewLine)
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


'out_Bake_Config_AsExcel = Function ( dic_config As dictionary(Of String, String) )
	Dim excel As New Microsoft.Office.Interop.Excel.Application
	Dim wb As Microsoft.Office.Interop.Excel.Workbook
	Dim ws As Microsoft.Office.Interop.Excel.Worksheet
	Dim strFileName As String = Environment.CurrentDirectory+"\"+now.tostring("yyMMdd")+"_Bake_Config.xlsx"
		
	' 초기변수 설정
	wb = excel.Workbooks.Add()
	ws = CType(wb.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
	ws.Name = "Config"
	
	' Header 작성 Cells(row, col)
	 excel.Cells(1, 1) = "Name"
	 excel.Cells(1, 2) = "Value"
	 excel.Cells(1, 3) = "Description"
	 ws.Range("A1:C1").Font.Bold = True 
	 ws.Range("A1:C1").Interior.Color = Color.LightGray
	
	Dim rowIndex As Integer = 1
	Dim keys As String()
	keys = dic_config.keys().toarray
	system.array.sort(keys)
	For Each key As String In keys 
	    rowIndex = rowIndex + 1
		excel.Cells(rowIndex, 1) = key
	 	excel.Cells(rowIndex, 2) = dic_config(key)
	Next
	
	' 열 너비 설정
	ws.Columns.AutoFit()
	
	' 파일 존재여부 확인
	If System.IO.File.Exists(strFileName) Then
	    System.IO.File.Delete(strFileName)
	End If
	
	' 저장 및 종료
	wb.SaveAs(strFileName)
	wb.Close()
	excel.Quit()
	'Return "저장 성공 : "+vbnewline+strFileName
'End Function
```

##### invokeCode Excel 제어 관련
[dataTable 생성 관련](https://stackoverflow.com/questions/41454836/vb-net-datatable-to-excel)   
위에 링크와 다른 것은 워크북과 시트에 이름 설정 부분이 다르다.   
Grammar StrictOn의 경우 InvokeCode에서 암시적으로 워크시트를 인식하지 못하여 CTpe을 써주어야 사용이 가능하다.   

```vb
Dim excel As Microsoft.Office.Interop.Excel.Application
Dim wb As Microsoft.Office.Interop.Excel.Workbook
Dim ws As Microsoft.Office.Interop.Excel.Worksheet

excel = New Microsoft.Office.Interop.Excel.Application
wb = excel.Add() 
ws = CTpe(wb.Worksheets.Add(), Workbook) ' 새로 추가된 시트가 ws에 담김
ws = CType(wb.ActiveSheet, Worksheet) ' 현재 작업 중인 시트가 ws에 담김
ws.Name = "변경할 시트명"
wb.SaveAs("저장할 이름/경로")

wb.Close()
excel.Quit()
```

## 인수 사용하는법
Extract WorkFlow하기 전에 변수 scope 설정만 잘 만져도 설정 편함.
지역변수는 variable로, 상위 scope와 연결된 변수는 인수로 자동설정됨.
invoke할 때 인수의 이름이 같으면 자동으로 설정해줌
인수 추가 단축키 : Ctrl + M

Invoke Workflow 에서 설정 방법
import Argument : 
 - Name/Direction/Type은 자동설정 됨, 이하는 Value값에 들어가는 내용 설명
 - in  : [value] 해당 워크플로우 시작 전에 넘겨줄 값을 입력한다.
 - out : [varible] 현제 워크플로우에서 받아올 값을 저장할 변수를 입력한다.
 - i/o : [varible] 이 변수의 값으로 해당 인수를 초기화하고 WF 종료 후 해당 인수값을 다시 이 변수에 넣는다.



### split(str, Environment.NewLine) 안먹힐 때
app 스크래핑 중 \r\n과 \n이 혼합되는 경우 full test나 get text로는 split 처리가 안됨. 
visual 스크래핑을 해야 정상적으로 문자열을 자를 수 있다.
### full text scraping => DataTable
tab = Chr(9) // enter = Environment.NewLine
ForEach : row in Split(str_DataTable,Environment.NewLine)
WriteLine : join(Split(row,Chr(9)), " | ") 

{"January","February","March","April","May","June","July","August","September","October","November","December"}




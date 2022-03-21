## 레퍼런스 모음
- #### [Markdown 사용법](https://gist.github.com/ihoneymon/652be052a0727ad59601)  
- #### [YAML 사용법](https://luran.me/397) , [YAML 공식문서](https://yaml.org/)
- #### [.NET 공식문서](https://docs.microsoft.com/ko-kr/dotnet/api/?view=net-6.0) (CODE짤 때 틈틈히 읽을 것)
   - [System.Linq](https://docs.microsoft.com/ko-kr/dotnet/api/system.linq?view=net-6.0)
   - [System.Data](https://docs.microsoft.com/ko-kr/dotnet/api/system.data?view=net-6.0)
   - [System.IO](https://docs.microsoft.com/ko-kr/dotnet/api/system.io?view=net-6.0)
   - System. [Reflection, Diagnostics, Net, Runtime, Xml, Text, Security ]


| Uipath 근본                 | 개발 관련                       | 리서치                              |
| --------------------------- | ------------------------------- | ----------------------------------- |
| [StateMachine][WdSm]        | **[※ Linq 코드 예시][Dev_LinqCode]**  | [참고하기 좋은 블로그][RS_참고하기] |
| [FlowChart][WdFc]           | [Linq 공식 문서][Dev_LinqDoc]   | [Custom 액티비티 만들기][RS_Custom] |
| [Sequnce][WdSq]             | [Split 공식 문서][Dev_SplitDoc] | [SetValue 관련][RS_SetValue]        |
| [SW접근성][SWAuto]          | [Join 공식 문서][Dev_JoinDoc]   | [Excel VB 참고 블로그][RS_ExcelVB]  |
| [Server-Side 권장안함][SSA] | [Strings 클래스][Dev_StrClass]  | [UIPATH 단축키][RS_UIPATH]          |

[WdFc]:https://docs.microsoft.com/en-us/dotnet/framework/windows-workflow-foundation/how-to-create-a-flowchart-workflow
[WdSq]:https://docs.microsoft.com/en-us/dotnet/framework/windows-workflow-foundation/how-to-create-a-sequential-workflow
[WdSm]:https://docs.microsoft.com/en-us/dotnet/framework/windows-workflow-foundation/how-to-create-a-state-machine-workflow
[SWAuto]:https://ehpub.co.kr/category/%ED%94%84%EB%A1%9C%EA%B7%B8%EB%9E%98%EB%B0%8D-%EA%B8%B0%EC%88%A0/sw%EC%A0%91%EA%B7%BC%EC%84%B1-%EA%B8%B0%EC%88%A0-ui-%EC%9E%90%EB%8F%99%ED%99%94/
[SSA]:https://support.microsoft.com/en-us/topic/considerations-for-server-side-automation-of-office-48bcfe93-8a89-47f1-0bce-017433ad79e2
[Dev_StrClass]:https://docs.microsoft.com/ko-kr/dotnet/api/microsoft.visualbasic.strings?view=net-5.0
[Dev_SplitDoc]:https://docs.microsoft.com/ko-kr/dotnet/api/microsoft.visualbasic.strings.split?view=net-5.0#Microsoft_VisualBasic_Strings_Split_System_String_System_String_System_Int32_Microsoft_VisualBasic_CompareMethod_
[Dev_JoinDoc]:https://docs.microsoft.com/ko-kr/dotnet/api/microsoft.visualbasic.strings.join?view=net-5.0#Microsoft_VisualBasic_Strings_Join_System_String___System_String_
[Dev_LinqDoc]:https://docs.microsoft.com/ko-kr/dotnet/visual-basic/programming-guide/language-features/linq/introduction-to-linq
[Dev_LinqCode]:https://linqsamples.com/linq-to-objects/
[RS_UIPATH]:https://docs.uipath.com/studio/docs/keyboard-shortcuts
[RS_SetValue]:https://stackoverflow.com/questions/10371712/how-to-assign-value-to-string-using-vb-net
[RS_ExcelVB]:https://kdsoft-zeros.tistory.com/36?category=846222
[RS_Custom]:https://mpaper-blog.tistory.com/15?category=832250
[RS_참고하기]:https://mpaper-blog.tistory.com/

#### Uipath 사용시 주의사항
패키지 및 studio 버전에 유의해야합니다. 
- "UiPath.UIAutomation.Activities" : 버전과 브라우저에 따라 Selector의 '구조'가 달라질 수 있습니다.
- "UiPath.System.Activities": 버전에 따라 kill process, while scope 등을 인식 못할 수 있습니다. 
- "UiPath.System.Activities": invokeCode에서 사용가능한 library 버전이 다릅니다. 
- "UiPath.Excel.Activities": 버전에 따라 excel scope 를 인식 못할 수 있습니다.

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
##### Tostring 관련
```vb
'숫자 표시형식
cint("1").ToString("0000") '= 0001

'날짜 표시 형식
now.ToString("yyyy_MM_dd")

' 문자열 & 아스키코드
Asc("A") '= 65 
Chr(65) '= "A"
' Convert 문자열 -> 아스키코드 번호
join(str_tmp.ToCharArray.Select(function(x) asc(x).ToString).ToArray, " ")

' ShortCode :
"문자열 : 아스키코드 "+vbNewLine+join(str_tmp.ToCharArray.Select(function(x) string.Format("{0} : {1}",x,asc(x).ToString) ).ToArray, vbNewLine)

' CSV 열 구분 : chr(44) | ,
' CSV 행 구분 : chr(13)+chr(10) | 줄바꿈
' 엑셀 셀 내부 줄바꿈 : chr(10) 
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
#### Array 다루기
##### Split, join

```vb
' uipath에서 사용할 수 있는 split 함수는 2가지 종류다.

' strings.split
StrArr = split("1,2,3",",") 
' StrArr : {"1","2","3"} 'Split(Str_source, Str_Seperator)
' StrArr = split("1 2 3") '기본 Seperator는 " " 이다.
' Strings.Split(Expression As String, Delimiter As String, ... )

' string.split
StrArr = "1,2,3".split(","c) 
' StrArr : {"1","2","3"} 
' String.Split(Seperator As Char(), ... )

StrArr = "1, .2:;3".Split(New Char() {" "c, ","c, "."c, ";"c, ":"c}, StringSplitOptions.RemoveEmptyEntries )
' StrArr : {"1","2","3"}

' strings.join
Str_Result = join(Split("1 2 3"), "|")
' Str_Result : "1|2|3" ' join(StrArr_source, Str_Seperator)
' Str_Result = join({"1","2"}) '기본 Seperator는 " " 이다.

Str_Result = join("1 2,3..::.4;;5 :6".Split(New Char() {" "c, ","c, "."c, ";"c, ":"c}, StringSplitOptions.RemoveEmptyEntries), " ")
'Str_Result : "1 2 3 4 5 6

* 참고 : UiPath에서 기본 split, join은 Strings 라이브러리의 것이다.
* "문자열".split , {"String Array",""}.join 은 strings.split, strings.join과 다른 함수이다.
```

#### Linq 다루기

##### 생성 관련
```vb
Dim StrArr as string() ' 생성
Dim StrArr as New String(){"1","2"} '#생성 및 할당

StrArr = Enumerable.Range(1,3).Select(function(x) x.ToString).ToArray
' StrArr : {"1","2","3"} 'Range(int_start, int_count) ' 기본 반환형은 Integer이다.
' IntArr = Enumerable.Range(0,3) '=> {0,1,2}

'repeat
StrArr = Enumerable.Repeat(of string)("1", 3).toarray
' StrArr : {"1","1","1"} 'Repeat(Type)(Str_source, int_count)

'null
StrArr = new string(2){} 
' StrArr : {null,null,null} '안에 있는 숫자는 최대 index
```

##### 편집 관련 Linq
```vb
'concat
StrArr = split("1 2").Concat( split("3 4 5") ).ToArray
' StrArr : : {"1","2","3","4","5"} ' split 과 join의 기본 구분자는 " "이다. 

' Distinct
StrArr = split("1 2 3 2 1 3 2 1").Distinct.ToArray
'StrArr : {"1","2","3"} '중복된 값 제거(뒤쪽 인덱스에 중복값 등장 시 누락시키는 로직)

' Select 
StrArr = split("1 2 3").Select(function(x) "["+x+"]").ToArray
' StrArr : {"[1]","[2]","[3]"} '원소 하나씩 select에 들어온 함수를 적용하여 갱신

' OrderBy
StrArr = split("2 3 1").OrderBy(function(x) cint(x) ).ToArray 
' StrArr : {"1","2","3"} ' 정렬-오름차순

' OrderByDescending
StrArr = split("2 3 1").OrderByDescending(function(x) cint(x) ).ToArray 
' StrArr : {"3","2","1"} ' 정렬-내림차순

'Reverse
StrArr = split("1 2 3").Reverse.ToArray 
' StrArr : {"3","2","1"} ' 순서- 거꾸로

'Skip, Take
StrArr = split("0 1 2 3 4 5").Skip(3).Take(2).ToArray
' StrArr : {"3","4"} 'Skip 개수만큼 앞에서 누락시키고, Take 개수만큼 취합

'Intersect
StrArr = split("0 1 2 3 4 5").Intersect(Split("1 3 5 7 9")).ToArray
' StrArr : {"1","3","5"} ' 교집합

'Linq 맛보기

' 쿼리형 : From Where Select
StrArr_tmp = (From x In Split("1 2 3 4 5 6") Where (2<Cint(x) AndAlso Cint(x)<5)  Select "["+x+"]").ToArray
'StrArr_tmp : {"[3]","[4]"}

' 람다형 : where(function() ).select(function())
StrArr_tmp = Split("1 2 3 4 5 6").Where(function(x) (2<Cint(x) AndAlso Cint(x)<5) ).Select(function(x) "["+x+"]").ToArray
'StrArr_tmp : {"[3]","[4]"}

```


#### 람다식에 인수 넣어주기
```vb
Console.WriteLine(((Function(num As Integer) num + 1)(5)).ToString)
' 그냥 람다식에 () 치고 바로 뒤에 (인수) 넣어주면 됨.

'람다식에 변수 2개 넣어줄 시 첫번째 변수는 값, 2번째 변수는 index를 의미함
StrArr_tmp = Split("가 나 다").Select(function(x,i) string.format("x='{0}'|i={1}",x, i.tostring) ).ToArray
'StrArr_tmp : {"x='가'|i=0" , "x='나'|i=1" , "x='다'|i=2" }

```
##### 자주 쓰게 되는 String Array 모음
```vb
' DT 열이름 Array 추출
StrArr = Enumerable.Range(0,dt_tmp.Columns.Count-1).Select(function(x) dt_tmp.Columns.Item(x).ColumnName).ToArray 

'DT 열 하나만 뽑아서 Array로 추출
StrArr = dt_tmp.AsEnumerable.Select(function(x) x("ColName").ToString).ToArray

StrArr = in_DIc_Config.Keys

' 파일명 제어
StrArr = Directory.GetFiles("절대경로") '각 파일의 절대경로 얻음
StrArr = Directory.GetFiles("절대경로").Select(function(x) new FileInfo(x).Name).ToArray '파일명 및 확장자만 얻음
StrArr = Directory.GetFiles("절대경로").Select(function(x) Split(x,"\").Last.ToString).ToArray '파일명, 확장자 얻음
StrArr = Directory.GetFiles(Environment.CurrentDirectory) '프로젝트 경로파일 얻음

For Each row as Data.DataRow in DT_tmp
    For Each item as Object in row.ItemArray
        ' item 을 item as String 으로 쓰면 Null 들어간 Row 처리할 떄 에러 발생함.
	' item 은 꼭 Object로 선언하고, 호출할 떄 ToString 처리하는 것이 안전함.
        Console.WriteLine( item.ToString ) 
    Next
Next
* Row를 ItemArray로 바꿀 때, 해당 변수를 받을 때는 꼭 Object로 받고 호출시 ToString을 하자.
* Row를 ItemArray로 바꾸는 과정에서 Null이 포함된 row에서 item을 String으로 받으면 에러가 발생한다. (Null을 String으로 형변환 못한다는 오류)
* 따라서 Row.ItemArray를 쓸 일이 있을 경우 Object()로 받거나, Select를 통해 ToString을 직접 시켜주는 게 좋다.

'[FileInfo]에 있는 유용한 속성값 Attributes, Name, Extension, FullName, DirectoryName, CreationTime, LastWriteTime, LastAccessTime...
'System.IO.Directory.GetFiles
'System.IO.FileInfo
```

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

### [Linq 설명](https://www.tutlane.com/tutorial/linq/linq-aggregate-function-with-example) (Lambda/ Query)
Lambda 식은 무명함수로, Function(x) x는 해당 Enun(Array 등)의 item을 부르는 변수다.   
반드시 x로 쓸 필요는 없고, ForEach안에 있는 로컬변수 선언하듯이 적당한 이름을 넣곤 한다.    
Dt.AsEnumerable을 사용한 경우 funtion(row) 이런식으로 지역변수명을 row로 선언하면 보다 알아보기 좋은 수식이 된다.   
(예시 : Arr_StrArr_dt = Dt.AsEnumerable.Select( funtion(row) row.itemArray.Select(function(x) x.Tostring).ToArray ).ToArray )

Query 식은 SQL 식과 유사한 쿼리식이다. From 이나 Aggregate 로 수식을 시작한다.   
쿼리식은 직관성이 떨어지기 때문에 개인적으로 lambda식만 사용하고 있다.   
- [select 문](https://linqsamples.com/linq-to-objects/projection/Select-anonymousType-lambda-vb) : 데이터를 수정/생성 할 떄 사용
- [GroupBY문](https://linqsamples.com/linq-to-objects/grouping/GroupBy-lambda-vb) : 인자로 받은 함수의 Return값을 key로 하여 구룹을 나눔.
- [ThenBy 문](https://linqsamples.com/linq-to-objects/ordering/ThenBy-lambda-vb) : Orderby로 정렬한 순서에서, 같은 레벨에 있는 항목을 제2 기준으로 정렬
- [Aggregate](https://linqsamples.com/linq-to-objects/aggregation/Aggregate-lambda-vb) : 특정 값을 누적하여 계산할 떄 사용. function(a,b)에서 a는 누적된 값, b는 작업중인 항목 의미.
- [Zip 문](https://linqsamples.com/linq-to-objects/other/Zip-lambda-vb) : 2개의 array를 동일한 index에 대해 대해 매핑 작업을 할 떄 쓰임. (ex : 백터 내적 연산 등)

##### 함수 설명
```vb
TypeName(<T>) : 해당 인자의 Type 이름을 String으로 반환한다.
file.WriteAllText("절대경로", Str_Source) :  해당 경로에 파일을 저장한다.
```
[VB 배열 관련](https://docs.microsoft.com/ko-kr/dotnet/visual-basic/programming-guide/language-features/arrays/)
[Linq 사용한 계산](https://docs.microsoft.com/ko-kr/dotnet/visual-basic/programming-guide/language-features/linq/how-to-count-sum-or-average-data-by-using-linq)
#### [Linq 사용 예시1](https://linqsamples.com/linq-to-objects/element)
#### [Linq 사용 예시2](https://www.tutlane.com/tutorial/linq/linq-aggregate-function-with-example)

```vb
TypeName({1,2,3}) 'Integer()
Dim numbers = New Integer() {1,2,3,4,5}
Dim numbers() As Integer = {1,2,3,4,5}

Aggregate x in {1,2,3,4,5} into sum ' 15
Aggregate x in {1,2,3,4,5} into count ' 5
Aggregate x in {1,2,3,4,5} into average '3
Aggregate x in split("1 2 3 4 5").Select(function(x) cint(x)) into sum

' a는 누적되어 저장된 값, b는 new Item. 
{1,2,3,4,5}.Aggregate(function(a,b) a+b) ' 15
{1,2,3,4,5}.Aggregate(function(a,b) a*b) ' 120 
{1,2,3,4,5}.Aggregate(10, Function(a,b) a+b) '25 : Aggregated numbers by addition with a seed of 10
{1,2,3,4,5}.sum() ' 15
{1,2,3,4,5}.Average() '3
{1,2,3,4,5}.Count()
{1,2,3,4,5}.Min()
{1,2,3,4,5}.Max()
{1,2,3,4,5}.
```

#### Groupby 사용하기
groupby는 인자로 넣어준 functnion의 계산값을 기준으로 data를 grouping합니다.    
groupby의 반환형은 iEnumerable(of iGrouping(of key, Tsource ))입니다.    
해석하자면, 함수를 호출하기 전에 원래 갖고 있었던 자료형(Tscource)을 유지하되, grouping할 떄 기준이 되었던 값을 key로 저장을 하고.    
여러개로 나눠진 Grouping 객체를 Enum의 형태로 반환한다는 뜻입니다.   
변수형 앞에 있는 i는 interface의 약자이며, 뭉뚱그려 생각하자면, 해당 객체의 interface(직접적으로 말하면 매소드=내장된 함수 등)를 사용할 수 있는 객체 형태라는 뜻입니다.   
(예시) ienumerable : 해당 객체는 Enumerable 자료형 안에 내장된 함수를 사용할 수 있습니다. (대충 AsEnumerable 처리 된 것과 유사하다고 생각하면 됩니다.)    
iEnumerable의 형태는 초급 개발자들이 공부하고 사용하기에 혼란스러울 수 있으므로, 개발 작업 시 ToArray() 처리를 하여, Array형식으로 저장 및 사용 하는 것을 권장합니다.    

##### 코드 예시
```vb
'이런 자료형을 사용한다는 것 정도만 보고 넘어갑니다.
Dim arr_groupby_BusinessNumber As System.Linq.IGrouping<System.String, System.Data.DataRow>[]

'DT 를 "사업자 등록번호"열의 값을 기준으로 Grouping 합니다.
arr_groupby_BusinessNumber = DT_Source.AsEnumerable.GroupBy(Function(row) row("사업자 등록번호").ToString).ToArray

'해당 Group의 사업자 등록번호를 모두 가져옵니다.
arr_Key_gpName_by_BSNum = arr_groupby_BusinessNumber.Select(function(gp) gp.key).ToArray

`사업자 번호별 공급가액, 세액의 부분합계를 가져옵니다.
arr_sum_taxBase_by_BSNum = arr_groupby_BusinessNumber.Select(function(gp) gp.sum(function(row) Cdbl(row("공급가액").tostring) )).ToArray
arr_sum_taxAmnt_by_BSNum = arr_groupby_BusinessNumber.Select(function(gp) gp.sum(function(row) Cdbl(row("세액").tostring) )).ToArray

' DT 생성
Dt_result = New DataTable
For Each colName As String In "사업자등록번호|공급가액_합계|세액_합계".Split("|"c)
    Dt_result.Columns.Add(colName, System.Type.GetType("System.String") )
Next
' DT 데이터 넣기
For Each i As Integer in enumerable.Range(0,arr_group_BSNum.Count)
    Dt_result.Rows.Add({arr_Key_gpName_by_BSNum(i).ToString, arr_sum_taxBase_by_BSNum(i).ToString, arr_sum_taxAmnt_by_BSNum(i).ToString})
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
#### [VBA : Excel Range -> HTML](https://stackoverflow.com/questions/54033321/excel-vba-convert-range-with-pictures-and-buttons-to-html)

#### 엑셀 시트명 갖고오기
excel scope에서 output workbook에 변수 만들기(wb)  
엑셀 시트명 확인 : if : wb.GetSheets.Contains(str_sheetName)

#### DataTable 관련
```vb
'### convert dt to Dictionary
DT_tmp.AsEnumerable.ToDictionary(Of String, Object)(Function (row) row("key").toString, Function (row) row("value").toString)

'데이터 필터링(abc열에서 값이 bcd인 행 찾기)
DT_tmp = DT_tmp.AsEnumerable.where(Function(x) x("abc").TosTing = "bdc").ToArray()

'Convert Column in Data Table to Array
DT_tmp.AsEnumerable().Select(Function (a) a.Field(of string)("columnname").ToString).ToArray()

'### Row Reverse 
DT_tmp = DT_tmp.AsEnumerable.Reverse().CopyToDataTable

'### Filtering abc열에서 값이 bcd인 행 모두 찾기
DT_tmp = DT_tmp.AsEnumerable.where(Function(x) x("abc").TosTing = "bdc").ToArray

' Copy는 열 이름에 상관 없이 값을 복사 붙여넣기 한다.
DT_test = DT_tmp.Copy()

' Clone은 데이터는 복사하지 않고 Columns만 복사해서 넣는다.
DT_test = DT_tmp.Clone()
```

1. 웬만한 서식은 모두 "조건부 서식"을 사용한다.
```EXCEL
조건문 관련
=Not(IsBlank($A1))        # 해당 행의 A열의 값이 비어있지 않으면 서식 적용
=ISNUMBER(SEARCH("1",A1)) # 해당 위치에 있는 문자열이 '1'을 포함하고 있으면 True

데이터 편집 관련
=TEXTJOIN(",",TRUE, B3,B4,...) # 구분자를 ","로 하고, 빈셀무시=True로 하여, 해당 좌표들의 값을 join함
```

2. Uipath > Excel에 CopyPasteRange 사용
- 해당 range의 병합된 셀의 서식까지 모두 붙여넣어진다.
- WriteRange할떄, 병합되어 생략된 위치에 data는 제대로 갱신되지 않으니 주의 
- 병합된 셀에 값을 갱신할 떄는(dummy rowdata로 생략된 셀을 체우든, writecell로 필요한 좌표만 찍든 해야함

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
## OutLook 사진첨부
Attach 로 이미지 파일 첨부하고, 
메일 본문을 html형식으로 설정한 이후 <img> 테그를 사용하여 보낼 수 있습니다.참
```html
<!--  첨부파일 이미지가 "123.png"라면 -->
<img src='cid:123.png' width='300' height='300' >
```
[참고](https://stackoverflow.com/questions/29369862/outlook-email-picture-attachment-not-showing-when-i-displaying-outlook-html-ema?rq=1)
위와 같이 이미지를 첨부하고 크기를 설정할 수 있습니다.



## [EDGE 관련 단축키](https://mainia.tistory.com/4086)
[Reference1](https://mainia.tistory.com/4086)
[Reference2](https://thelumine.wordpress.com/2015/08/27/microsoft-edge-keyboard-shortcuts/)

|단축키 | 기능 |
|--|--|
| Ctrl+W | 현재 탭 닫기 |
| Ctrl+1~Ctrl+8 | 창의 특정 위치에 있는 탭으로 이동 |
| Ctrl+9 | 창의 마지막 탭으로 이동 |
| Ctrl+0 | 창의 화면 비율 100%로 조정 |
| Ctrl+Shift+T | 마지막으로 닫았던 탭 열기 |
| Ctrl+Tab | 창의 다음 탭으로 이동	|
| Ctrl+Shift+Tab | 창의 이전 탭으로 이동	|
| Ctrl+U | 페이지 소스 보기 |
| Ctrl+Shift+I | 개발자 도구 패널 표시/숨김 | 


 
#### Naming Tip Boolean
| 자료형  | 요령                     | 예시                    | 비고                                    |
| ------- | ------------------------ | ----------------------- | --------------------------------------- |
| Boolean | Is_값                    | Is_WhiteSpace           | null, Empty, Nothing, nan               |
| Boolean | Is_상태                  | Is_Exist                | NonDebugMode,Vaild                      |
| Boolean | Is_대상_상태             | Is_Element_Exist        | Button_Changed, Window_Closed           |
| Boolean | Is_동작_상태             | Is_ProcrssName_Done     | Done, Succeed, Fail, Processing         |
| Boolean | Is_동작_대상_상태        | Is_Download_PDF_Succeed | Search_Data_HasError                    |
| Boolean | Has_대상                 | Has_Error               | SystemError, BussinessError, AttachFile |
| Boolean | Should_동작              | Should_Retry            | Start, Stop, Skip, Extract, Print       |
| Boolean | Should_동작_대상         | Should_Upload_Directory | Clear_Dir, Send_Email                   |
| Boolean | IsNot, HasNot, ShouldNot | -                       |

#### Naming Tip Integer
| 자료형  | 요령               | 예시                        |
| ------- | ------------------ | --------------------------- |
| Integer | 수열Max            | Int_RetryMax                |
| Integer | 횟수Cnt  (Count)   | Int_RetryCnt                |
| Integer | 위치Idx  (Index)   | Int_RowIdx                  |
| Integer | 배열Num  (Number)  | Int_TransActionNum          |
| Integer | 대상_속성          | Int_DT_Width                |
| Integer | 대상_속성_세부속성 | Int_Scrollbar_ClickOffset_X |



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
UI 구조가 a > b > c > d 이런식으로 되어 있을 때. Selector는     
a > d, a> c 이런식으로 중간 단계가 누락되어 있을 수 있다.  
셀렉터가 너무 불안정 할 때는 부모 selector를 확인해서 하위로 진입하는 식으로 잡는게 안정성 있다.


### Empty, Nothing, null
Empty : 변수 생성 후 초기화 하지 않음 (string, int 생성만 했을 때)
Nothing : 해당 변수가 참조하는 개체가 없음 (DataTable, Dictionary, List 등)
null : 알 수 없는 데이터(DataTable 생성 후 값을 입력하지 않음)
* Tostring은 에러를 배출하지 않는다.
* Nothing인 객체에 Tosting을 하면 에러가 발생한다. (참조개체가 없으므로 Tostring 매소드 호출할 수 없기 때문)
* System.Convert.ToString(Nothing)을 하게 되면 ""가 반환된다. Conver.ToString는 이미 정의되어 있고 null, Nothing 체크를 하기 때문
Nothing.Tostring = 에러 : 참조개체가 없어 "개체.ToString" 정의되지 않음
Convert.ToString(Nothing) = "" : ToString 함수는 Convert에서 정의 됨, null, Nothing 체크가능

### [엑셀]
pivotTable 
- 단축키 : Alt+D + p
- 경로 : 리본>삽입>피벗테이블
- 옵션 : 새 시트로 생성 > [필터, 열 레이블, 행 레이블, 값] 설정
- 수정 : 테이블 우클릭 > 피벗테이블 필드 표시
 
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

## 팝업 셀렉터 잡기
uiexplorer로 브라우저 팝업을 잡으려고 하면 Studio가 멈추는 경우가 있다.  
이 떄는 Selector를 수동으로 입력해서 셀렉터를 파악하여 개발해야한다.  

#### 요령
1. [Uipath 공식](https://docs.uipath.com/studio/docs/about-selectors)에서 셀렉터가 지원하는 테그 확인 
2. \<html>, \<wnd>, \<ctrl> 등 테그 속성을 확인하고, 적절한 값으로 셀렉터 찍기  
2.1. 팝업에 있는  Text는 name이나 title 속성에 들어있을 확률이 높다.  
2.2. \<wnd/> 에서 title, aaname 으로 보이는 글자를 넣어본다.  
2,3. \<ctrl/> 에서 role, name, text 등을 잡아본다.   
3. target > WaitForReady > None 넣어놓는다.(무한대기 방지)  
4. 예시 

##### UiAutomation.Activities 19.10?
```xml
<!-- Edge 팝업 확인 버튼 클릭 -->
<html app='msedge.exe' url='*' />
<ctrl role='dialog' />
<ctrl  role = 'push button' name='확인'/>
```

  ```xml
<!-- Edge 팝업 내 나가기 버튼 클릭 -->
<wnd app='msedge.exe' title='*나갈까요*' />
<ctrl name='*나가기*' />
 ```
  ```xml
<!-- GetText 크롬 팝업 내 텍스트 지정 -->
<html app='chrome.exe' title='*' />
<ctrl role='dialog' />
<ctrl role='text' name='*.*' />
 ```
```xml
<!-- Click 크롬 팝업 내 확인/계속 버튼 -->
<ctrl role='dialog' />
<ctrl  role = 'push button' name='계속'/>
```

##### UiAutomation.Activities 21.4.4

```xml
<!-- Edge Alert 텍스트 박스 - GetText -->
<wnd app='msedge.exe' title='*' />
<ctrl role='dialog' />
<ctrl idx='15' role='pane' />
```


```xml
<!-- Edge Alert/Confirm 1번째 버튼 - Click -->
<wnd app='msedge.exe' title='*' />
<ctrl role='dialog' />
<ctrl role='pane' idx='4' />
<ctrl role='push button' idx='1' />
<!-- 해당 UI의 버튼에 써있는 글자는 name 속성으로 - Get Attr -->
<!-- Confirm에서 취소버튼 등 2번째 버튼은 마지막 테그의 idx = '2' 입력 -->
```

###  문자열, 배열 내 중복 제거
```vb
str_tmp = join(split(str_tmp,vbNewLine).Distinct().ToArray,vbNewLine)
```
## DataSet 사용하는 방법
#### InvokeCode 사용 주의사항
##### MethodName 의 경우 대소문자를 구분한다.
"Add"로 써야할 것을 "add"로 쓸 경우 에러가 발생한다.
#### DataSet - Only Activity
```
0. 변수패널 : dt_tmp :DataSet, dt_tmp = DataTable
1. Assign  :  ds_tmp = new dataset
2. Assign  :  ds_tmp = new DataTable("테이블명")

3. Invoke Method : 
 - TargetType : (null)
 - TargetObject : ds_tmp.Tables
 - MethodName : Add
 - Parameters : in | DataTable | dt_tmp
 * MethodName에 "add"나 "ADD" 넣으면 오류 발생하니 주의
 
4. Add Data Column : ds_tmp.Tables("테이블명") <- "열이름1"

5. Log Message : ds_tmp.Tables("테이블명").Columns.Item(0).ColumnName
ㄴ 반환 : "열이름1"
```
#### DataSet - whith Build DataTable
```
0. 변수패널 : dt_tmp :DataSet, dt_tmp = DataTable
1. Assign  :  ds_tmp = new dataset
2. Build Data Table : out = dt_tmp
3. Assign : dt_tmp.TableName = "테이블명"

4. Invoke Method : 위와 동일
 
5. Log Message : ds_tmp.Tables("테이블명").Columns.Item(0).ColumnName
ㄴ 반환 : "열이름1"
```

#### DataSet - Only Inovk Code
```vb
변수패널 : ds_tmp : DataSet

Invoke Code : 
 - Argument : out_ds_DataSet | Out | DataSet | ds_tmp
 - 코드 내용
 	"""
	Dim dt_log As DataTable
	Dim dt_data1 As DataTable

	dt_log = New dataTable("Log")
	For Each col As String In {"성공여부", "비고"}
		dt_log.Columns.Add(col)
	Next
	
	dt_data1 = New dataTable("Data")
	For Each col As String In {"열1", "열2"}
		dt_data1.Columns.Add(col)
	Next

	out_ds_DataSet = New Dataset
	out_ds_DataSet .Tables.Add(dt_log)
	out_ds_DataSet .Tables.Add(dt_data1)
	"""

Log Message : ds_tmp.Tables("Log").Columns.Item(0).ColumnName =>  반환 : "성공여부"
Log Message : ds_tmp.Tables("Data").Columns.Item(0).ColumnName =>  반환 : "열1"
```


## 자주 쓰는 알고리즘

### LinQ 사용하여 특정 조건을 만족하는 row와 col만 추출하기
행 필터링, 열 필터링
[defaultView](https://newbiedev.tistory.com/24)
[Linq](https://www.vb-net.com/VB2015/Language/LINQ.%20Update,%20Combine,%20Custom%20func,%20LINQ%20Providers%20for%20Anonymous,%20Extension,%20Lambda,%20Generic,%20String,%20XML,%20Dataset,%20Arraylist,%20Assembly,%20FileSystem.pdf)

#### 행 필터링, 샘플링
```vb
' Row Sampling
drArr_tmp = dt_tmp.AsEnumerable.Where(Function(x) x("a").ToString.Contains("1")).ToArray
```

### 열 필터링, 선택
```vb
' Col selecting
if : drArr.count > 0 
dt_tmp = dt_tmp.DefaultView.ToTable(false, {"a","c","e"} )
dt_tmp = drArr_tmp.CopyToDataTable.DefaultView.ToTable(False, {"a","c","e"})

' CopyToDataTable은 count가 0일 때 에러가 발생하기 때문에 DataRow[] 를 사용하여 예외처리
' DefaultView : 첫번쨰 인자는 false로 해야한다. True로 할 경우 오류 발생(distinct 속성)
' 반환값은 "a,b,c" 총 3개의 열만 가진다.
' 여담으로 LinQ에서 IEnumerable 의 I는 interface의 약자다. array 대산 iEnumerable<DataRow> 사용 가능
```
##### 열 22개, 행 31800 개인 엑셀로 Test한 결과
- 수행시간 | 작업내역
- 0.11초 | LinQ 사용 : 행/열 모두 필터링
- 2.50초 | FiltterDataTable : 행/열 모두 필터링
- 1.48초 | InvokeCode : For Each - setField
- 2.00초 | ForEachRow : 액티비티 사용

### DataTable 값 update (Invoke code)

```vb
For Each row  As datarow In dt_tmp.AsEnumerable()
	row.SetField("ColName","Value")
Next
' argument dt_tmp는 in으로 주어도 정상적으로 수정됨
```

### DataTable n번째 행부터 m개 Row만 선택
행 필터링, 행 샘플링(sampling)
```vb
dt_tmp.AsEnumerable.Skip(int_n).Take(int_m).CopyToDataTable
' int_n + int_m > dt_tmp.rows.count : 일때 에러 발생
```

### python 에 range(n)을 uipath에서 배열로 만들기
```
Assign : arr_tmp = New String(n){}
ForEach : 속성{Value : arr_tmp , index : int_i, item : _ }
    Assign : arr_tmp(int_i) = int_i.Tosting

* New String(n){} = {"",""}  
```
#### string array 에서 Null,Empty,whiteSpace 항목 제거하는 방법
arr_tmp.Where(Function(x) not string.IsNullOrWhiteSpace(x)).ToArray
```
Assign : arr_tmp = New String(2){""," ","abc"}
Assign : arr_tmp = arr_tmp.Where(Function(x) not string.IsNullOrWhiteSpace(x)).ToArray
* arr_tmp : string[1] {"abc"}
```

#### 엑셀 읽어서 해더명에 공백 제거
```
ForEach : col in dt_tmp.Columns
    Assign : col.ColumnName = col.ColumnName.Replace(" ","")
    
* col 자료형 = System.Data.DataColumn
```


## 자주쓰는 명령어
cint(), cdbl(), .Tostring  
Split(txt , ": ") // as string array  
join(row.ItemArray," | ") // as string   
{"A","B","C"}.contains("A") // isin, has 함수 VB버전    
dic_tmp.ContainsKey("213") // dict에서 key 있는지 확인   
file.Exists(str_FilePath) // 경로에 파일 있는지 확인  
TypeName() // object로 케스팅된 string은 string으로 뜸

dt_tmp.Columns.Contains("Column1") # dt에 해당 열 있는지 확인   
dt_tmp.Columns(0).ColumnName = “newColumnName” # 열 이름 바꾸기   
System.Drawing.Color.Gray  # 엑셀 셀 책 체우기 할 떄 사용  
TimeSpan.FromMilliseconds(int_delayTime) # 딜레이 시간 넣을 떄 사용    
Asc("A") = 65  ,  Chr(65) = "A"
```
숫자 표시형식 : 1 -> 0001  
cint("1").ToString("0000")  
"1".PadLeft(4,cchar("0"))
```

#### 파일명, 폴더명 가져오기
str_targetPath = "폴더 경로"
Directory.GetFiles(str_targetPath) # string[] 형태로 경로 반환
Directory.GetDirectories(str_targetPath) # string[] 형태로 경로 반환
new FileInfo(str_targetPath) ## fileinfo 객채 선언
new System.IO.FileInfo("str_targetPath").LastWriteTime # 마지막 수정시간 얻기

#### 프로세스 작업시간 구하기
assign : dtm_ProcessStartTime = DateTime.Now   
delay : 00:01:30   
writeLine : "작업수행시간 : " + cint(DateTime.Now.Subtract(dtm_ProcessStartTime).TotalSeconds).ToString + " 초"   

#### 한글 날짜 요일 표시 방법 [출처](https://m.blog.naver.com/PostView.naver?isHttpsRedirect=true&blogId=elduque&logNo=120096308343)
1단계 : import 패널에서 System.Globalization 추가(CultureInfo 객체 사용을 위함)  
2단계 : writeLine 이나 LogMessage에서 출력값 확인하기목요일  
- DateTime.Today.ToString("dddd", CultureInfo.CreateSpecificCulture("ko-KR"))  #목요일
- DateTime.Today.ToString("ddd", CultureInfo.CreateSpecificCulture("ko-KR"))   #목
- Date.ParseExact("20210212", "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)  

in_TransactionItem.SpecificContent("WIID").ToString // 큐에서 특정값 호출

## 파워쉘로 작업/파일 실행시키는 방법
[스케줄러로 돌릴 때 참고](https://deje0ng.tistory.com/78)
[uipath 문서](https://docs.uipath.com/robot/docs/arguments-description)

```cmd
# 파워쉘 열기
1. window + X : 트레이 열기
2. a : PowerShell 관리자 권한으로 실행
3. cls 
4. (Get-PSReadlineOption).HistorySavePath
```

```cmd
# Uipath 경로로 이동
cd "C:\Program Files (x86)\UiPath\Studio\"

# 딜레이 시간 넣기
timeout 1 
Start-Sleep -Seconds 1


# 파일 실행
.\UiRobot.exe execute   --file "파일절대경로(xaml)"

# 작업 실행
.\UiRobot.exe execute  -p "작업이름"

# 예시.bat
cd "C:\Program Files (x86)\UiPath\Studio\"
.\UiRobot.exe execute  -process "KS출근" -input "{ 'str_code' : '178606' ,'str_ID' : 'wbpark'}"

```

#### 초기화 관련
New String(){"1","2"} #string array 생성및 할당   
New String(n){} #원소가 n개인 string array 생성

new List(of int32)  
new List(of string)   
new List(of string)(new string(){"가","나","다","라"})   
new List(of String) from {{"보험"},{"세금"}}   

New Dictionary(of string, int32)   
New Dictionary(of string,int32) from {{"red",50},{"yellow",10},{"green",80}}   
New Dictionary(Of String, string()) # 문자열 배열    
New Dictionary(of string, object) from {{"test1","50"},{"test2","10"},{"test3","80"}}   

#### 비밀번호 관련
(new Net.NetworkCredential("",Str)).SecurePassword // secureString 반환
(new Net.NetworkCredential("",Str)).Password // String 반환
new System.Net.NetworkCredential(string.Empty, secureStr).Password //SecureStr -> str
 
 #### linq 관련
 dt key값 겹치는 것 갱신

 ```
dt_destination : [key,열1,add열2,add열3]   
dt_sorce : [key,add열2,add열3]   
dataRow : System.Data>DataRow
 
 ForEachRow : row in dt_destination
    if : dt_sorce.AsEnumerable.Where(function(x) x("key").ToString.Trim=row("key").ToString).Count=1
        then : 
		assign : dataRow = dt_sorce.AsEnumerable.Where(function(x) x("key").ToString.Trim=row("key").ToString)(0)
		assign : row("add열2") = dataRow.Item("add열2")
		assign : row("add열3") = dataRow.Item("add열3")
	Else : (Do Nothing)
 ```

 
 dt 중복행 제거  
[출처 - 열 하나만](https://forum.uipath.com/t/delete-duplicate-row-based-on-one-column-duplicate-data/217700)  
[출처 - 열 둘이상](https://mpaper-blog.tistory.com/27?category=832250)   
- CopyToDataTable 쓸 때는 row 개수 확인 필수.

```
DT_input      // System.Data.DataTable
IEnum_DataRow // System.Collections.Generic.IEnumerable<System.Data.DataRow>
DT_output     // System.Data.DataTable
 
assgin : IEnum_DataRow = DT_input.AsEnumerable().GroupBy(Function(x) convert.ToString(x.Field(of object)("colName"))).SelectMany(function(gp) gp.ToArray().Take(1))
 
assgin : DT_output = IEnum_DataRow.CopyToDataTable

# 열 한개
assgin :
DT_output = DT_input.AsEnumerable().GroupBy(Function(x) convert.ToString(x.Field(of object)("colName"))).SelectMany(function(gp) gp.ToArray().Take(1)).CopyToDataTable
 
# 열 두개 
(From p In DT_input.AsEnumerable() Group By x = New With { Key.a =p.Item("A"), Key.b=p.Item("B")} Into Group Select Group(0)).ToArray().CopyToDataTable()

 ```
 dt 열 2개로 정렬
 ```
 Assign : dt_sorce =  
 (From x In dt_sorce.AsEnumerable() Order By convert.Tostring(x("colName1")),convert.ToString(x("colName2")) Select x).CopyToDataTable
 
# 정렬 우선순위1. colName1
# 정렬 우선순위2. colName2
실제 수행 : colName2로 정렬 수행 (최종적으로) colName1로 정렬
 ```
 
dt, Linq 관련 이슈
```
요약 : 
* EnumerableRowCollection<DataRow> 자료형에 dt.AsEnumerable.Where 값을 넣으면 호출시 오류가 발생함.
* 하지만 그냥 EnumerableRowCollection<DataRow> 자료형에 dt.AsEnumerable 넣어서 호출하는 건 괜찮음.
* 우회법으로는 array<DataRow> 자료형에dt.AsEnumerable.Where.toArray 넣는 것임.
* 추정컨데 원인은 변수에 where값 할당 시 값 대신 "힙 어딘가에 있는 임시 주소"가 들어가는 것 같음. 
  그냥 dt.AsEnumerable까지는 정상적인 주소가 할당되는데, where 연산 결과값은 힙 어딘가에서 바로 초기화 되어 주소를 찾을 수 없어는 것 같음.


변수 :
    dt_totalOutput : DataTable
    dt_resultStep2 : DataTable
    arr_DataRows: DataRow[]
    enum_DataRows : EnumerableRowCollection<DataRow>

정상1 : 바로 호출
    ForEach row in dt_totalOutput :
        LogMessage : 
             dt_resultStep2.AsEnumerable.Where(function(x) x("관리번호").ToString.Trim=row("관리번호").ToString).Count
    
정상2 : DataRow[] 저장 후 호출
    ForEach row in dt_totalOutput :
        Assign : 
	    arr_DataRows = dt_resultStep2.AsEnumerable.Where(function(x) x("관리번호").ToString=row("관리번호").ToString).ToArray
	LogMessage : 
	    arr_DataRows.count
	
정상3 : EnumerableRowCollection<DataRow> 저장 후 호출
    ForEach row in dt_totalOutput :
        Assign : 
	    enum_DataRows = dt_resultStep2.AsEnumerable
        ForEach dataRow in enum_DataRows :
	    if : dataRow("관리번호").ToString = row("관리번호").ToString
	       DoSomeThing
	LogMessage : SomeThing	      

###########
오류 : EnumerableRowCollection<DataRow> 저장 후 호출
    ForEach row in dt_totalOutput :
        Assign : 
	    enum_DataRows = dt_resultStep2.AsEnumerable.Where(function(x) x("관리번호").ToString=row("관리번호").ToString)
	LogMessage : 
	    enum_DataRows.count
>>> 오류문구 : 
  Log Message: Activity '1.9: VisualBasicValue<Object>' cannot access this public location reference because it is only valid for activity '1.14: VisualBasicValue<EnumerableRowCollection<DataRow>>'.  Only the activity which obtained the public location reference is allowed to use it.
```
dataTable 열이름 변경 : "Column1" -> "New Column"   
Assign : dt_tmp.Columns(dt_tmp.Columns.IndexOf("Column1")).ColumnName = "New Column"   
 
ForEachRow 액티비티에서 row 를 다른 테이블에 AddDataRow 를 할 경우  
“Add data row : This row already belongs to another table.” 오류 메시지가 나온다.   
이 떄는 AddDataRow에서 row를 array로 넘기면 해결된다. : row.ItemArray  

 
 #### cron 사용법
[cron 문법](https://www.leafcats.com/94)  
[cron 디버깅](http://www.cronmaker.com/;jsessionid=node0109oq4rr76ib71nhs60lrghk15443008.node0?0)  
```
Cron 예시 :
매년 매월 20일과 25일 13시 0분 0초
0 0 13 20,25 * ? *
(초, 분, 시, 일, 월, 요일, 년)


* : 매번
? : 모름 (일, 요일)에만 사용 가능하다
/ : 증가 (ex : 10/15 = 10분부터 시작해 매 15분마다
# : k#N 이달 N번째 K요일 (ex : 5#2 = 이달 두번째 목요일)
L : 마지막 (일,요일)에만 사용 가능 (ex : 6L = 이달 마지막 금요일)
W : 가장 가까운 평일 (ex : 10W = 이달 10일에서 가장 가까운 평일)
"-" , "," : 범위 (1-12 =1월-12월, "20,25" = 20일과 25일
```

##### 엑셀 읽기 오류 관련
[UIPATH 엑셀 StacOverFlow](https://stackoverflow.com/questions/2424718/how-to-know-if-a-cell-has-an-error-in-the-formula-in-c-sharp)  
[UIPATH 엑셀 오류 정리글](https://deokpals.tistory.com/12)  
```
    ErrDiv0 = -2146826281,
    ErrGettingData = -2146826245,
    ErrNA = -2146826246,
    ErrName = -2146826259,
    ErrNull = -2146826288,
    ErrNum = -2146826252,
    ErrRef = -2146826265,
    ErrValue = -2146826273
```

## 단축키 
### 인라인
 - 변수 추가 : Ctrl + K
 - 인수 추가 : Ctrl + M
 - 자동 완성 : Ctrl + space

### 액티비티 관련
 - 이름 변경 : F2
 - 설명 추가 : Shift + F2 (activity 설명)
 - 액티비티 찾기 : Ctrl + F (xaml 안에서 위치 찾아줌)
 - 액티비티 삭제 : Ctrl + E
 - 액티비티 주석 : Ctrl + D 
 - 액티비티 시도 : Ctrl + T (Try Catch)
 - 액티비티 추가 : Ctrl + Shift + T 
 - FlowChart set start node : 우클릭 + A
 
## sellector 변수처리
{{item}} : 이렇게 중괄호 2개로 덮히면 셀럭터 변수처리가 가능하다.
xml이기 때문에 주소 참조(변수호출)는 가능하지만 연산( {{ (cint(item)+2).Tostring }} )은 불가하다.
계산할 게 있다면 최종 결과물을 넣은 변수를 sellector xml에 넣어줘야한다. Tostring 미리 작업 해야만한다.
assign : temp = (cint(item)+2).Tostring , sellector edit : {{temp}} 호출
물론 와일드카드랑 같이 쓰면서 적당히 sellector를 조작하는 게 편하다.
셀렉터에 idx값이 필요한 순간도 있기는 한데, 일반적인 상황에서 웬만하면 idx값이 필요없게 짜는 걸 권장한다. 

## simulate 옵션
simulate click: 
 - True : 클릭 이벤트 호출 (실제 마우스 커스 안움직임) 
 - 장점 : 백그라운드에서 작업하기 때문에 마우스 사용 가능
 - 단점 : 가끔 element에 포커스가 안잡히는 문제가 생길 수 있다.
 - 권장 : 안정적인 작업을 위해서는 Flase 로 유지하는게 좋다. 
simulate type : 
 - True : 백그라운드에서 타이핑 이벤트 처리
 - 단점 : [key(enter)] 등 simulate key event 사용 불가

## Excel Activity
엑셀에서 alt+enter로 생성된 문자열은 chr(10)이다.
셀 안에 있는 줄바꿈으로 문자열을 나누려면 chr(10)으로 split하라.
Environment.NewLine.ToArray 으로 나눌 수도 있다.
```vb
str_test.split(chr(10)) 'is equivalent to
str_test.split(Environment.NewLine.ToArray)
```
### Excel 설치x 컴퓨터
.xlsx 파일만 작업이 가능하다.
**읽기** : 시스템.파일.통합문서.Read Range
**쓰기** : 시스템.파일.통합문서.Write Range
 - 수식 자동완성 안됨. "=SUM(A1:B1)"입력 시 문자열 그대로 들어감.

엑셀이 설치된 컴퓨터에서는 어플스코프 사용 권장함. 어플 스코프 사용시 범위 임력으로 수식 자동 체우기도 지원됨.

### Excel 설치된 컴퓨터 
#### Excel application scope 사용
엑셀 어플리케이션으로 파일을 직접 열어서 작업이 진행됨. 스코프 안에서는 현제 작업 중인 파일을 기준으로 제어됨.

**입력** : 앱통합.Excel.테이블.Read Range

**출력** : 앱통합.Excel.테이블.Write Range
- "=SUM(A1:B1)"입력시 상대위치를 통해 수식이 적용됨.

데이터 읽을 때는 반드시 header가 체크 되어있는지 확인할 것 (데이터 row가 밀려쓰기 될 수 있다.)

## 데이터 테이블
AddDataColumn : 새로운 열을 추가할 때는 반드시 열추가를 먼저 한다. 

For Each Row : 프로그래밍.데이터 테이블.For Each Row
- 지역변수 Row의 항목을 받아올 때는 Get Row Item 액티비티를 사용. 
- 지역변수를 받아올 때는 형변환이 복잡하므로 GenericValue로 받는 걸 권장.
item 변수를 추가적으로 선언하지 않고 row 데이터 조작할 때
- VB표현식으로 row(IndexNum)으로 호출.
- 정수형 데이터가 예상되는 경우 VB기준 integer.Parse(row(IndexNum).ToString)로 사용해야 한다.
- 굳이 string으로 변환 후 integer로 변환하는 이유는 Double를 바로 integer로 변환할 수 없기 때문.
- 변수선언 없이 데이터 조작예시 : 
```
Activity : UIPath.Excel.Activities.ExcelWrite Cell
Range : "C"+(DT2.Rows.IndexOf(row)+1).ToString
Value : (integer.Parse(row(0).ToString)+integer.Parse(row(1).ToString)).ToString
```

## UI 상호작용
### Input Dialog 
- UI 상으로 사용자가 입력한 값을 받는다.
  
### TypeInto
- 텍스트 내용 입력
- 콘트롤, 엔터 등 hotkey는 [k(enter)] 이런 식으로 이루어진다. 
- 그 외 window나 broswer에 키입력은 Sendkey 액티비티를 사용하여 전달한다.


## 기타

# 부록
``` 
TypeName() # 주의사항
UIPath에서 ForEach등 지역변수 설정 기본은 Object다.
별도의 설정을 하지 않으면 ForEach Array<String>의 item은 Object로 변환되어 사용된다.
TypeName(item)은 가장 구체적인 데이터 형을 표시하기 때문에 Object안에서 "String"을 반환한다.
WriteLine등 String 입력을 강제하는 속성에 item을 사용하면 Object->String 연산을 진행한다.
String에서 케스팅된 Object라도 암시적으로 String으로 변환하지 못하기 때문에 오류가 발생한다.

* ForEach에서 지역변수의 형태를 모호하게 설정하면 호출 부분에서 오류가 발생하니 주의한다.
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


## 비즈니즈 관점
원버튼 : 마우스, 키보드 조차 모른다고 생각하고 접근
 - 컴퓨터 전원을 켜면 자동으로 실행되는 실행파일 만들기 (UIPath 방법은 아님)
 - 전원이 켜진 상태라면 unattended 로봇으로 작성한 스캐줄대로 움직이거나
 - 사람이 Robot으로 실행버튼 정도는 누르는 식으로 하는 것 (이 번거로움을 문제로 상정할 수 있다.)

## 프로젝트 정리 팁
### 구조짜기
프레임웤 : Main.xaml 은 state muchine으로 지정
프로젝트 : 간단한 프로젝트는 Main을 flowchart으로 지정
 - Main 워크플로는 [Invoke Workflow] 만으로 구성 (사용할 것만 연결; 우클릭+A)
 - Invoke에 연결된 워크플로는 모두 flowchart로 구성 
 - 해당 flowchart는 오직 sequence와 flowDecision으로만 구성
 - 모둔 구현은 각각의 sequence 안에서 관리.

프로젝트 : [Main - InvokeWorkFlow(플로우차트 호출)] 
수행과제 : [FloswChart - Sequence(과제별 정리)]
구현내용 : [Sequence - step by step]

### Config 만들기
엑셀에서 세팅값 불러오기 // Dictionary 형식으로 Config 받아온다.
1. ReadRange // 초기화 Config = New Dictionary(Of String, Object) 
2. ForEachRow // Config(row("Name").Tostring) = row("Value")
3. MessageBox // Config("test1") // test1은 해당 엑셀파일 Name열에 있던 이름

팁 : config로 테이블 만들기
config("header") = "열1,열2,열3"
config("열1") = "값1,값2"
config("열2") = "값1,값2,값3"
config("열3") = "값1,값2,값3,값4"
ForEach : row, Split(config("header"),",")
   WriteLine : Join(Split(row,",")," | ")

### UIPath 개발 시 참고
Microsoft workflow에서 GUI 툴 그대로 가져와서 사용함.
미국에는 데스크탑 앱 개발 시 소프트웨어 접근성(시각/청각 장애우도 사용 가능한 기능)이 요구된다. 그 접근성 앱 개발을 위한 도구가 발전해서 RPA 프로그램이 된 것이다. [UIPath를 사용하지 않고 MS workflow로 트레킹하는 영상](https://ehpub.co.kr/category/%ED%94%84%EB%A1%9C%EA%B7%B8%EB%9E%98%EB%B0%8D-%EA%B8%B0%EC%88%A0/sw%EC%A0%91%EA%B7%BC%EC%84%B1-%EA%B8%B0%EC%88%A0-ui-%EC%9E%90%EB%8F%99%ED%99%94/) 따라서 UIPath 개발 시 MS의 WorkFlow 문서를 참고하는 것이 좋다. 여담으로 [MS office는 서버-side 개발을 권장하지 않는다.](https://support.microsoft.com/en-us/help/257757/considerations-for-server-side-automation-of-office?wa=wsignin1.0%3Fwa%3Dwsignin1.0) -by 이석원 프로님

### 테스트 케이스
[테스트 케이스 자동화](https://academy.uipath.com/learningpath-viewer/2234/1/155237/16)

# UIPath Advance 팁
- [1번 문제_해쉬코드](https://wooaoe.tistory.com/61)
- [2번 문제_연레포트](https://wooaoe.tistory.com/62)


# 윈도우 자격증명 쓰는법
1. 시작메뉴 - 자격 증명 검색 - 자격증명 관리자
2. windows 자격증명 - [일반 자격증명]에 추가
3. 인터넷 또는 네트워크 주소 : 해당 정보 관리할 이름으로 설정 ex : ACME-login
4. uipath 패키지 다운로드 (Uipath.Credentials.Activites 다운)
   - Get Secure Credential (windows 자격증명에 있는 값 가져옴)
   - CredenrialType : Generic
   - PersistanceType : Enterprise
   - Target : [일반 자격증명]에 있는 '인터넷 또는 네트워크 주소'값
비고 : 비밀번호 일반 텍스트로 출력하는 법 (sequre string -> string)
 String plainStr = new System.Net.NetworkCredential(string.Empty, secureStr).Password


# 문자열
message box
- "\n"이 먹히지 않아 vbCrLf 나 Environment.NewLine 을 써야한다.

for each row 에서 table 모두 출력 
WriteLine :
   join(row.ItemArray," | ") //row 가 string인 경우
   Join({row(0).ToString,row(1).ToString,row(2).ToString}," | " )

### split(str, Environment.NewLine) 안먹힐 때
app 스크래핑 중 \r\n과 \n이 혼합되는 경우 full test나 get text로는 split 처리가 안됨. 
visual 스크래핑을 해야 정상적으로 문자열을 자를 수 있다.
### full text scraping => DataTable
tab = Chr(9) // enter = Environment.NewLine
ForEach : row in Split(str_DataTable,Environment.NewLine)
WriteLine : join(Split(row,Chr(9)), " | ") 

{"January","February","March","April","May","June","July","August","September","October","November","December"}



# 오케 사용법

## 로봇 연결하기
conect robot to orchestrator
1. [오케 접속](https://cloud.uipath.com/koreabegmifx/DefaultTenant/)
2. MY FOLDERS - default - Robots - Standard Robot 생성 (=> 머신 생성 됨)
3. Tanent - Machines - Machine key 등 data 복사
4. Asistant - Preferance - Orchestrator Setting
   - Machine Key : 설정
   - Machine Name : 복붙
   - URL : https://cloud.uipath.com/koreabegmifx/DefaultTenant/
   - Machine Key : 복붙

## 설명
태넌트 : 계정에 포함된 서버
   - 감사 : 사용자의 모든 액션이 로그로 남는다.
   - 사용자 : 다른 사람을 추가할 수 있다. (커뮤니티 무료 버전을 불가)


Default폴더 : 
   - 로봇 관리 : standard (Developer=studio 사용하겠다. , Unattened = orche로 원격으로 쓰겠다. 스케줄링)
   - 환경 관리 : 로봇을 구룹지어놓은 것, 사람 인사관리할 때 부서 나누는 것처럼 로봇 그룹핑한 것이 환경
   - 자동화 : 프로세스 생성 및 스케줄링 가능. 로봇-편집-형식에서 로봇을 unattend로 돌린 후 사용해보자.
     - 프로세스 만들기 : 패키지(orche에 등록해놓은 것)연결, 우선순위(동시에 일 여러개 받을 때 처리순서)
     - 프로세스 실행 : 로봇 선택; 틍정로봇=(unattended로봇 선택 ), 동적할당=(환경 내 쉬고있는 로봇에게 일시킴)
     - 트리거 : 시간 or 큐 중 선택 가능
       - 시간 : 고급설정에 cron 표현식 있음. google에서 cron표현식 만들어주는 사이트 들어가서 세부설정 가능
       - 휴무일 : [테넌트-설정-휴무일] 에서 설정 가능. 저장해놓은 휴무일이 있으면 트리거 설정할 때 해당 휴무일에 쉴지 선택가능
       - 스케줄 : 특정 시간마다 반복되는 트리거, 설정 들어가서 꺼놓으면 해당 스케줄을 지우지 않고도 사용안함 가능

# 큐, 트렌젝션
큐는 Default 폴더에서 관리함.
   - studio에서 add하면 New상태의 큐item이 생성됨.
   - Get Transaction 하면 New 상태인 item 하나 가져옴. (item 상태 : In Process로 변경됨)
   - Get Queue 하면 특정 상태인 큐item을 가져올 수 있다. (어떤 상태든 가능하다.)
   - Set Transaction 하면 해당 큐item을 [성공 or 실패]로 설정할 수 있다.
   - 사용이 끝난 큐item은 [성공 or 실패]상태로 남겨둘지 Delete 큐item으로 완전 제거할지 결정하면 된다.
Transaction은 큐에서 가장 먼저 들어온 New 상태의 item을 의미한다고 생각하면 될 것 같다.

#### 다루기 
1. 패키지 배포 : studio에서 게시버튼 클릭 => 오케에 프로젝트명으로 패키지 업로드됨
2. 프로세스 생성 : 해당 패키지/수행환경 설정 후 저장
3. 스캐줄 관리 : 트리거에서 해당 프로세스/환경 설정 후 cron등으로 스캐줄 설정
4. 작업 수행 명령 : 작업 탭에서 새로만들기(시작버튼 "▶" 누르면 새로 만들기 나옴, 패키지/환경 설정)

## Datatable 
[row reverse 하는법](https://excelcult.com/how-to-reverse-a-datatable-in-uipath/)
```DT_tmp = DT_tmp.AsEnumerable.Reverse().CopyToDataTable```
데이터 필터링(abc열에서 값이 bcd인 행 찾기)
```DT_tmp = DT_tmp.AsEnumerable.where(Function(x) x("abc").TosTing = "bdc").ToArray ```

Convert Column in Data Table to Array
```DT_tmp.AsEnumerable().Select(Function (a) a.Field(of string)("columnname").ToString).ToArray()```

DT header to Array
```vb
list_header = new list(of String)
ForEach : item in DT_tmp
   Add To Collectoin <String> : item.ColumnName
arr_header = list_header.ToArray
```

### AddDataColumn : 열 추가
ForEachRow : DataTable의 행을 지역변수 Row에 담아 반복
- 형변환이 복잡해서 row는 GenericValue로 받는 걸 권장.
- 정수형 항목 꺼내기 : 
```
integer.Parse(row(IndexNum).ToString)
' row는 바로 형변환이 불가하여 문자열 변환을 거쳐 변환함
```

### DT 할당
보통 Build DataTable Activity를 사용하여 초기화한다.
Build DataTable로 재작한 더미 테이블(row,col = 0,0) 인스턴스는 Notiong이 아니다.
```
' Copy는 열 이름에 상관 없이 값을 복사 붙여넣기 한다.
DT_test = DT_tmp.Copy()

' Clone은 데이터는 복사하지 않고 Columns만 복사해서 넣는다.
DT_test = DT_tmp.Clone()
```

### DT 열 추가
보통 열 추가 후  ForEachRow로 초기화 한다.. 
```
Add DataTable Columns(dt_tmp, "new_col") 
ForEachRow(dt_temp) :
    Assgin : row.item("new_col") = "초기화 값"
```

### DT 행 추가 
dt_tmp1에 dt_tmp2의 데이터 추가
```
Merge DataTable :  (Activity) 
	Destination = dt_tmp1 
	Source = dt_tmp2
```

### Merge 열 이름 다를 때 
```
'| col1 | col 2 | 
'| tmp1 |       | 
'|      |  tmp2 | 
'이런 식으로 행이 이상하게 붙는다.
```

### Join 키 값으로 합치기
```
Join DataTable 액티비티 사용 
```


### Row Reverse 
```
DT_tmp = DT_tmp.AsEnumerable.Reverse().CopyToDataTable
```

### Filtering
abc열에서 값이 bcd인 행 모두 찾기
```
DT_tmp = DT_tmp.AsEnumerable.where(
    Function(x) x("abc").TosTing = "bdc").ToArray
```

### DT 값 호출
1행 1열 Table의 값 호출
dt_dumy(0)("dumy").ToString


## Process 확인
```
Get Processes : processes = processes
Assign : array = processes.AsEnumerable().Where(Function(x) x.ProcessName.Contains("OUTLOOK")).ToArray
if : array.Count >0 : 
	Write Line : "존재" + array(0).ProcessName
	
타입 : 
processes = System.Collections.ObjectModel.Collection<System.Diagonotics.Process>
array = System.Dianotics.Pcrocess[]
```

### sendMassage 주의사항

### 복사 붙여넣기 주의사항
메일 내 이미지 복사 붙여넣기시 주의사항
(1) [Background] outlook.Message 사용
    결과 : 문자열 추출 가능 (표 서식 및 이미지 깨짐)
    의견 : 데이터 추출용으로 사용할 수 있을 것 같습니다.
(2) [Forground] 웹브라우저로 메일함 들어가서 복사 붙여넣기
    결과 : 표 서식 및 이미지 정보 정상적으로 복제 가능
    의견 : 브라우저를 사용하여 웹에서 호출가능한 이미지 경로와 표 서식이 복사됩니다.
(3) [Forground] outlook App 에서 복사 붙여넣기
    결과 : 이미지 경로 깨짐으로 [x박스] 생성됨
    의견 : 이미지 경로가 메일서버에서 로컬PC로 다운받은 경로로 ITMS 웹에서 인식할 수 없습니다.
(4) [Forground] Uipath : Set Clipboard / Get From Clipboard 액티비티 사용
    결과 : Set Clipboard 과정에서 테이블 서식과 이미지 정보가 모두 깨집니다.
    의견 : 표와 이미지가 들어간 데이터는 UiPath Clipboard 액티비티를 사용할 수 없습니다. 
    
### SendMassage 주의사항 
입력값 필드에 shift 체크를 풀어놓고 "A"를 입력할 떄, 실제로 입력되는 정보는 shift + 'a'다. 
ctrl+c를 할 경우 c를 대문자로 입력하게 되면 ctrl+shift+c가 실행되므로 주의바람.

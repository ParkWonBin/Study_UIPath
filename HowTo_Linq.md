
##### Linq 수행 속도
열 22개, 행 31800 개인 엑셀로 Test한 결과
|수행시간 | 작업내역
|---|---|
|0.11초 | LinQ 사용 : 행/열 모두 필터링
|2.50초 | FiltterDataTable : 행/열 모두 필터링
|1.48초 | InvokeCode : For Each - setField
|2.00초 | ForEachRow : 액티비티 사용

#### String Array
uipath에서 사용할 수 있는 split 함수는 3가지 종류다.
1. [Strings.Split(Of String, String)][Split1]
2. [String.Split(Of Char)][Split2]
3. [System.Text.Regularexpressions.Regex.Split(Of String, String)][Split3]
참고 : 정규식 공부
[Split1]:https://docs.microsoft.com/ko-kr/dotnet/api/microsoft.visualbasic.strings.split?view=net-6.0
[Split2]:https://docs.microsoft.com/ko-kr/dotnet/api/system.string.split?view=net-6.0
[Split3]:https://docs.microsoft.com/ko-kr/dotnet/api/system.text.regularexpressions.regex.split?view=net-6.0

```vb

' 1. Strings.Split
' Strings.Split(Expression As String, Delimiter As String, ... )
StrArr = split("1,2,3",",") 
' StrArr : {"1","2","3"} 'Split(Str_source, Str_Seperator)
' StrArr = split("1 2 3") '기본 Seperator는 " " 이다.

' 2. String.Split
' String.Split(Seperator As Char(), ... )
StrArr = "1,2,3".split(","c) 
StrArr = "1,.2 3;".Split(" ,.;:".ToCharArray, StringSplitOptions.RemoveEmptyEntries )
StrArr = "1,.2 3;".Split(New Char() {" "c, ","c, "."c, ";"c, ":"c}, StringSplitOptions.RemoveEmptyEntries )
' StrArr : {"1","2","3"}

'3. Regex.Split
StrArr = System.Text.Regularexpressions.Regex.Split("1,2@3#4test5","[,@#]|test")
' StrArr : {"1","2","3","4"} 
' 두번째 인자는 구분자에 해당하는 패턴을 정규식으로 입력한다. 

' strings.join
Str_Result = join(Split("1 2 3"), "|")
' Str_Result : "1|2|3" ' join(StrArr_source, Str_Seperator)
' Str_Result = join({"1","2"}) '기본 Seperator는 " " 이다.
```

* 참고 : UiPath에서 기본 split, join은 Strings 라이브러리의 것이다.
* "문자열".split , {"String Array",""}.join 은 strings.split, strings.join과 다른 함수이다.

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

'람다식에 변수 2개 넣어줄 시 첫번째 변수는 값, 2번째 변수는 index를 의미함
StrArr_tmp = Split("가 나 다").Select(function(x,i) x & i.tostring ).ToArray
'StrArr_tmp : {"가0" , "나1" , "다2" }

' 쿼리형 : From Where Select 맛보기
StrArr_tmp = (From x In Split("1 2 3 4 5 6") Where (2<Cint(x) AndAlso Cint(x)<5)  Select "["+x+"]").ToArray
'StrArr_tmp : {"[3]","[4]"}

' 람다형 : where(function() ).select(function())
StrArr_tmp = Split("1 2 3 4 5 6").Where(function(x) (2<Cint(x) AndAlso Cint(x)<5) ).Select(function(x) "["+x+"]").ToArray
'StrArr_tmp : {"[3]","[4]"}

```

#### DataTable과 함께 사용
```vb
'Extract DataColumn Names From DataTable
StrArr_ColNames = Enumerable.Range(0,dt_tmp.Columns.Count-1).Select(function(x) dt_tmp.Columns.Item(x).ColumnName).ToArray 

'DT 열 하나만 뽑아서 Array로 추출
StrArr_ColData = dt_tmp.AsEnumerable.Select(function(x) x.item("ColName").ToString).ToArray
```
#### 람다식에 인수 넣어주기
```vb
Console.WriteLine(((Function(num As Integer) num + 1)(5)).ToString)
' 그냥 람다식에 () 치고 바로 뒤에 (인수) 넣어주면 됨.
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

##### GroupBy 통한 DataTable 만들기 코드 예시
```vb
Dim arr_groupby_BusinessNumber As System.Linq.IGrouping<System.String, System.Data.DataRow>[]

'DT 를 "사업자 등록번호"열의 값을 기준으로 Grouping 합니다.
arr_groupby_BusinessNumber = DT_Source.AsEnumerable.GroupBy(Function(row) row("사업자 등록번호").ToString).ToArray

'해당 Group의 사업자 등록번호를 모두 가져옵니다.
arr_Key_gpName_by_BSNum = arr_groupby_BusinessNumber.Select(function(gp) gp.key).ToArray

'사업자 번호별 공급가액, 세액의 부분합계를 가져옵니다.
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

### LinQ 사용하여 특정 조건을 만족하는 row와 col만 추출하기
행 필터링, 열 필터링
[defaultView](https://newbiedev.tistory.com/24)
[Linq](https://www.vb-net.com/VB2015/Language/LINQ.%20Update,%20Combine,%20Custom%20func,%20LINQ%20Providers%20for%20Anonymous,%20Extension,%20Lambda,%20Generic,%20String,%20XML,%20Dataset,%20Arraylist,%20Assembly,%20FileSystem.pdf)

### DataColumn Filtering
```vb
```
### DataRow Filtering
행 필터링, 행 샘플링(sampling)
```vb
dt_tmp.AsEnumerable.Skip(int_n).Take(int_m).CopyToDataTable
' int_n + int_m > dt_tmp.rows.count : 일때 에러 발생
```

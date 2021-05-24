## 자주쓰는 명령어
cint(), cdbl(), .Tostring  
Split(txt , ": ") // as string array  
join(row.ItemArray," | ") // as string   
{"A","B","C"}.contains("A") // isin, has 함수 VB버전    
dic_tmp.ContainsKey("213") // dict에서 key 있는지 확인   
file.Exists(str_FilePath) // 경로에 파일 있는지 확인  
TypeName() // object로 케스팅된 string은 string으로 뜸  
new list(of string)  
new dictionary(of string, int32)  
New Dictionary(Of String, string()) # 문자열 배열  
New String(){"1","2"} #string array 생성및 할당   

dt_tmp.Columns.Contains("Column1") # dt에 해당 열 있는지 확인   
dt_tmp.Columns(0).ColumnName = “newColumnName” # 열 이름 바꾸기   
System.Drawing.Color.Gray  # 엑셀 셀 책 체우기 할 떄 사용  
TimeSpan.FromMilliseconds(int_delayTime) # 딜레이 시간 넣을 떄 사용    
Asc("A") = 65
Chr(65) = "A"
```
숫자 표시형식 : 1 -> 0001  
cint("1").ToString("0000")  
"1".PadLeft(4,cchar("0"))
```

dataTable 열이름 변경 : "Column1" -> "New Column"   
Assign : dt_tmp.Columns(dt_tmp.Columns.IndexOf("Column1")).ColumnName = "New Column"   

dtm_tmp = Date.ParseExact("20210212", "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)  
 
ForEachRow 액티비티에서 row 를 다른 테이블에 AddDataRow 를 할 경우  
“Add data row : This row already belongs to another table.” 오류 메시지가 나온다.   
이 떄는 AddDataRow에서 row를 array로 넘기면 해결된다. : row.ItemArray  
 
 
 dt 중복행 제거 [출처](https://forum.uipath.com/t/delete-duplicate-row-based-on-one-column-duplicate-data/217700)  
 ```
DT_input      // System.Data.DataTable
IEnum_DataRow // System.Collections.Generic.IEnumerable<System.Data.DataRow>
DT_output     // System.Data.DataTable
 
assgin : IEnum_DataRow = DT_input.AsEnumerable().GroupBy(Function(x) convert.ToString(x.Field(of object)("colName"))).SelectMany(function(gp) gp.ToArray().Take(1))
 
assgin : DT_output = IEnum_DataRow.CopyToDataTable

# 한줄 코드
assgin :
DT_output = DT_input.AsEnumerable().GroupBy(Function(x) convert.ToString(x.Field(of object)("colName"))).SelectMany(function(gp) gp.ToArray().Take(1)).CopyToDataTable
 
 ```
 
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

 
 #### 셀렉터로 크롬창 팝업 잡기
 팝업창 선택할 떄 페이지 로드가 멈추는 곳이 있다.   
 target > WaitForReady > None 넣어놓기   
 팝업창 내용 스크랩용  
 ```xml
<html app='chrome.exe' title='*' />
<ctrl role='dialog' />
<ctrl role='text' name='*.*' />
 ```
 팝업창 버튼 클릭용
```xml
<ctrl role='dialog' />
<ctrl  role = 'push button' name='계속'/>
```
#### 엑셀 시트명 갖고오기
excel scope에서 output workbook에 변수 만들기(wb)  
엑셀 시트명 확인 : if : wb.GetSheets.Contains(str_sheetName)

#### Linq 관련
```
### convert dt to Dictionary
DT_tmp.AsEnumerable.ToDictionary(Of String, Object)(Function (row) row("key").toString, Function (row) row("value").toString)

데이터 필터링(abc열에서 값이 bcd인 행 찾기)
DT_tmp = DT_tmp.AsEnumerable.where(Function(x) x("abc").TosTing = "bdc").ToArray()

Convert Column in Data Table to Array
DT_tmp.AsEnumerable().Select(Function (a) a.Field(of string)("columnname").ToString).ToArray()

### Row Reverse 
DT_tmp = DT_tmp.AsEnumerable.Reverse().CopyToDataTable

### Filtering
abc열에서 값이 bcd인 행 모두 찾기
DT_tmp = DT_tmp.AsEnumerable.where(Function(x) x("abc").TosTing = "bdc").ToArray

Copy는 열 이름에 상관 없이 값을 복사 붙여넣기 한다.
DT_test = DT_tmp.Copy()

Clone은 데이터는 복사하지 않고 Columns만 복사해서 넣는다.
DT_test = DT_tmp.Clone()
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

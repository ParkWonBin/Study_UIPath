
# UIPath Advance 팁
- [1번 문제_해쉬코드](https://wooaoe.tistory.com/61)
- [2번 문제_연레포트](https://wooaoe.tistory.com/62)


#### 프로세스 작업시간 구하기
assign : dtm_StartTime = DateTime.Now   
delay : 00:01:30   
LogMessate : DateTime.Now.Subtract(dtm_StartTime).TotalSeconds.ToString("0.00") + " 초"   


#### 한글 날짜 요일 표시 방법 [출처](https://m.blog.naver.com/PostView.naver?isHttpsRedirect=true&blogId=elduque&logNo=120096308343)
1단계 : import 패널에서 System.Globalization 추가(CultureInfo 객체 사용을 위함)  
2단계 : writeLine 이나 LogMessage에서 출력값 확인하기목요일  
- DateTime.Today.ToString("dddd", CultureInfo.CreateSpecificCulture("ko-KR"))  #목요일
- DateTime.Today.ToString("ddd", CultureInfo.CreateSpecificCulture("ko-KR"))   #목
- Date.ParseExact("20210212", "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)  
in_TransactionItem.SpecificContent("WIID").ToString // 큐에서 특정값 호출


### Uipath 단축키
 - 변수 추가 : Ctrl + K
 - 인수 추가 : Ctrl + M
 - 자동 완성 : Ctrl + space
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

### ForEach Activity 사용 관련
\TypeName() # 주의사항
UIPath에서 ForEach등 지역변수 설정 기본은 Object다.
별도의 설정을 하지 않으면 ForEach Array<String>의 item은 Object로 변환되어 사용된다.
TypeName(item)은 가장 구체적인 데이터 형을 표시하기 때문에 Object안에서 "String"을 반환한다.
WriteLine등 String 입력을 강제하는 속성에 item을 사용하면 Object->String 연산을 진행한다.
String에서 케스팅된 Object라도 암시적으로 String으로 변환하지 못하기 때문에 오류가 발생한다.
* ForEach에서 지역변수의 형태를 모호하게 설정하면 호출 부분에서 오류가 발생하니 주의한다.

### TypeInto
- 텍스트 내용 입력
- 콘트롤, 엔터 등 hotkey는 [k(enter)] 이런 식으로 이루어진다. 
- 그 외 window나 broswer에 키입력은 Sendkey 액티비티를 사용하여 전달한다.

### SendMassage 주의사항 
입력값 필드에 shift 체크를 풀어놓고 "A"를 입력할 떄, 실제로 입력되는 정보는 shift + 'a'다. 
ctrl+c를 할 경우 c를 대문자로 입력하게 되면 ctrl+shift+c가 실행되므로 주의바람.

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
#### [VBA : Excel Range -> HTML](https://stackoverflow.com/questions/54033321/excel-vba-convert-range-with-pictures-and-buttons-to-html)

#### 엑셀 시트명 갖고오기
excel scope에서 output workbook에 변수 만들기(wb)  
엑셀 시트명 확인 : if : wb.GetSheets.Contains(str_sheetName)

엑셀 어플리케이션으로 파일을 직접 열어서 작업이 진행됨. 스코프 안에서는 현제 작업 중인 파일을 기준으로 제어됨.

**입력** : 앱통합.Excel.테이블.Read Range

**출력** : 앱통합.Excel.테이블.Write Range
- "=SUM(A1:B1)"입력시 상대위치를 통해 수식이 적용됨.

데이터 읽을 때는 반드시 header가 체크 되어있는지 확인할 것 (데이터 row가 밀려쓰기 될 수 있다.)

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


## 비즈니즈 관점
원버튼 : 마우스, 키보드 조차 모른다고 생각하고 접근
 - 컴퓨터 전원을 켜면 자동으로 실행되는 실행파일 만들기 (UIPath 방법은 아님)
 - 전원이 켜진 상태라면 unattended 로봇으로 작성한 스캐줄대로 움직이거나
 - 사람이 Robot으로 실행버튼 정도는 누르는 식으로 하는 것 (이 번거로움을 문제로 상정할 수 있다.)

### UIPath 개발 시 참고
Microsoft workflow에서 GUI 툴 그대로 가져와서 사용함.
미국에는 데스크탑 앱 개발 시 소프트웨어 접근성(시각/청각 장애우도 사용 가능한 기능)이 요구된다. 그 접근성 앱 개발을 위한 도구가 발전해서 RPA 프로그램이 된 것이다. [UIPath를 사용하지 않고 MS workflow로 트레킹하는 영상](https://ehpub.co.kr/category/%ED%94%84%EB%A1%9C%EA%B7%B8%EB%9E%98%EB%B0%8D-%EA%B8%B0%EC%88%A0/sw%EC%A0%91%EA%B7%BC%EC%84%B1-%EA%B8%B0%EC%88%A0-ui-%EC%9E%90%EB%8F%99%ED%99%94/) 따라서 UIPath 개발 시 MS의 WorkFlow 문서를 참고하는 것이 좋다. 여담으로 [MS office는 서버-side 개발을 권장하지 않는다.](https://support.microsoft.com/en-us/help/257757/considerations-for-server-side-automation-of-office?wa=wsignin1.0%3Fwa%3Dwsignin1.0) -by 이석원 프로님

### 테스트 케이스
[테스트 케이스 자동화](https://academy.uipath.com/learningpath-viewer/2234/1/155237/16)


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


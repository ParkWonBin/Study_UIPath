  ### 공사중!
1. 시스템 이슈 및 Howto 작성 => Issues 통해 정리중
2. 함수 및 app 정리 => 프로그램 제작 중.. Function 에 중간결과 나옴
3. 추가로 알아볼것

열 추가와 동시에 기본값 넣기. eval 함수처럼 넣을 수 있음
```vb
 DT_result.columns.Add("Process Code", gettype(string), string.Format("substring({0}, 1,8)", "ReleaseName")) 
```

  ## 레퍼런스 모음
- #### [Markdown 사용법](https://gist.github.com/ihoneymon/652be052a0727ad59601)  
- #### [YAML 사용법](https://luran.me/397) , [YAML 공식문서](https://yaml.org/)
- #### [.NET 공식문서](https://docs.microsoft.com/ko-kr/dotnet/api/?view=net-6.0) (CODE짤 때 틈틈히 읽을 것)
   - [System.Linq](https://docs.microsoft.com/ko-kr/dotnet/api/system.linq?view=net-6.0), [System.Data](https://docs.microsoft.com/ko-kr/dotnet/api/system.data?view=net-6.0) ,  [System.IO](https://docs.microsoft.com/ko-kr/dotnet/api/system.io?view=net-6.0)
   - System. [Reflection, Diagnostics, Net, Runtime, Xml, Text, Security ]
   - [System.Windows.Automation](https://docs.microsoft.com/ko-KR/dotnet/api/system.windows.automation?view=windowsdesktop-6.0)
   - 자동화 관련 [.NET UI 자동화](https://docs.microsoft.com/ko-kr/dotnet/framework/ui-automation/ui-automation-overview)


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
 

#### UIPath 개발 시 참고
Uipath는 Microsoft workflow에서 GUI 툴 그대로 가져와서 사용한다. [MS workflow로 트레킹 하는 영상][MS_WF]   
RPA 프로그램은 미국 데스크탑 앱 개발 SW 접근성(시각/청각 장애우도 사용 가능해야 함) 도구가 발전해서 만들어졌다. 따라서 UIPath 개발 시 MS의 WorkFlow 문서를 참고하는 것이 좋다. 여담으로 [MS office는 Server-Side 개발을 권장하지 않는다.][MS_ref2] -by 이석원 프로님
[테스트 케이스 자동화](https://academy.uipath.com/learningpath-viewer/2234/1/155237/16)

[MS_WF]:https://youtu.be/pPnpFvM02HA
[MS_ref2]:https://support.microsoft.com/en-us/help/257757/considerations-for-server-side-automation-of-office?wa=wsignin1.0%3Fwa%3Dwsignin1.0

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

### Empty, Nothing, null
Empty : 변수 생성 후 초기화 하지 않음 (string, int 생성만 했을 때)
Nothing : 해당 변수가 참조하는 개체가 없음 (DataTable, Dictionary, List 등)
null : 알 수 없는 데이터(DataTable 생성 후 값을 입력하지 않음)
* Tostring은 에러를 배출하지 않는다.
* Nothing인 객체에 Tosting을 하면 에러가 발생한다. (참조개체가 없으므로 Tostring 매소드 호출할 수 없기 때문)
* System.Convert.ToString(Nothing)을 하게 되면 ""가 반환된다. Conver.ToString는 이미 정의되어 있고 null, Nothing 체크를 하기 때문
Nothing.Tostring = 에러 : 참조개체가 없어 "개체.ToString" 정의되지 않음
Convert.ToString(Nothing) = "" : ToString 함수는 Convert에서 정의 됨, null, Nothing 체크가능
 
### OutLook 사진첨부
Attach 로 이미지 파일 첨부하고, 
메일 본문을 html형식으로 설정한 이후 <img> 테그를 사용하여 보낼 수 있습니다.참
```html
<!--  첨부파일 이미지가 "123.png"라면 -->
<img src='cid:123.png' width='300' height='300' >
```
[참고](https://stackoverflow.com/questions/29369862/outlook-email-picture-attachment-not-showing-when-i-displaying-outlook-html-ema?rq=1)
위와 같이 이미지를 첨부하고 크기를 설정할 수 있습니다.
 

 #### cron 사용법
[cron 문법](https://www.leafcats.com/94)  
[cron 디버깅](http://www.cronmaker.com/;jsessionid=node0109oq4rr76ib71nhs60lrghk15443008.node0?0)  
```text
Cron 예시 :
  - 매년 매월 20일과 25일 13시 0분 0초
  - 0 0 13 20,25 * ? *
  - (초, 분, 시, 일, 월, 요일, 년)


* : 매번
? : 모름 (일, 요일)에만 사용 가능하다
/ : 증가 (ex : 10/15 = 10분부터 시작해 매 15분마다
# : k#N 이달 N번째 K요일 (ex : 5#2 = 이달 두번째 목요일)
L : 마지막 (일,요일)에만 사용 가능 (ex : 6L = 이달 마지막 금요일)
W : 가장 가까운 평일 (ex : 10W = 이달 10일에서 가장 가까운 평일)
"-" , "," : 범위 (1-12 =1월-12월, "20,25" = 20일과 25일
```

# String 관련
#### Integer To String
```vb
'날짜 표시 형식
cint("1").ToString("0000") '= 0001  
"1".PadLeft(4,"0"c) '= 00001
```

#### DateTime To String
[한글 요일 표시](https://m.blog.naver.com/PostView.naver?isHttpsRedirect=true&blogId=elduque&logNo=120096308343)
- 1단계 : import 패널에서 System.Globalization 추가(CultureInfo 객체 사용을 위함)  
- 2단계 : writeLine 이나 LogMessage에서 출력값 확인하기목요일  

```vb
'날짜 표시 형식
now.ToString("yyyy_MM_dd")
'한글 날짜 표시
DateTime.Today.ToString("dddd", CultureInfo.CreateSpecificCulture("ko-KR"))  #목요일
DateTime.Today.ToString("ddd", CultureInfo.CreateSpecificCulture("ko-KR"))   #목
Date.ParseExact("20210212", "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)  
```

#### String To Charactor
```vb
' 문자열 & 아스키코드
Asc("A") '= 65 
Chr(65) '= "A"

' ShortCode : 문자열 - 아스키코드 번호
join(str_tmp.ToCharArray.Select(function(x) string.Format("{0} : {1}",x,asc(x).ToString) ).ToArray, vbNewLine)

' CSV 열 구분 : chr(44) = ','
' CSV 행 구분 : chr(13)+chr(10) = \r\n
' 엑셀 셀 내부 줄바꿈 : chr(10) 
```

' 파일명 제어
StrArr = System.IO.Directory.GetFiles("절대경로") '각 파일의 절대경로 얻음
StrArr = System.IO.Directory.GetFiles(Environment.CurrentDirectory) '프로젝트 경로파일 얻음

#### Linq Functions
```vb
'메일 참조에서
StrArr_ReceiveMail_CC = System.Text.RegularExpressions.Regex.split(Mail_ReceiveMail.cc.ToString,"[^\w@.-]+").Where(function(x) x.Contains("@")).ToArray 
```

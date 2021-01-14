### 색갈 테스트
<span style="color:red"> 강조점 </span>
<span style="color:darkgreen"> 액티비티 이름</span>
<span style="color:darkorange"> 인라인 코드 </span>

## 단축키 
### 인라인
 - 변수 추가 : Ctrl + K
 - 자동 완성 : Ctrl + space
  
### 액티비티 관련
 - 이름 변경 : F2
 - 설명 추가 : Shift + F2 (activity 설명)
 - 액티비티 삭제 : Ctrl + E
 - 액티비티 주석 : Ctrl + D 
 - 액티비티 시도 : Ctrl + T (Try Catch)
 - 액티비티 추가 : Ctrl + Shift + T 

## Exsel Activity
### Exsel 설치x 컴퓨터
.xlsx 파일만 작업이 가능하다.
**읽기** : 시스템.파일.통합문서.Read Range
**쓰기** : 시스템.파일.통합문서.Write Range
 - 수식 자동완성 안됨. "=SUM(A1:B1)"입력 시 문자열 그대로 들어감.

엑셀이 설치된 컴퓨터에서는 어플스코프 사용 권장함. 어플 스코프 사용시 범위 임력으로 수식 자동 체우기도 지원됨.

### Exsel 설치된 컴퓨터 
#### Exsel application scope 사용
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
### 문자열
message box
- "\n"이 먹히지 않아 vbCrLf 나 Environment.NewLine 을 써야한다.

### UIPath 개발 시 참고
Microsoft workflow에서 GUI 툴 그대로 가져와서 사용함.
미국에는 데스크탑 앱 개발 시 소프트웨어 접근성(시각/청각 장애우도 사용 가능한 기능)이 요구된다. 그 접근성 앱 개발을 위한 도구가 발전해서 RPA 프로그램이 된 것이다. [UIPath를 사용하지 않고 MS workflow로 트레킹하는 영상](https://ehpub.co.kr/category/%ED%94%84%EB%A1%9C%EA%B7%B8%EB%9E%98%EB%B0%8D-%EA%B8%B0%EC%88%A0/sw%EC%A0%91%EA%B7%BC%EC%84%B1-%EA%B8%B0%EC%88%A0-ui-%EC%9E%90%EB%8F%99%ED%99%94/) 따라서 UIPath 개발 시 MS의 WorkFlow 문서를 참고하는 것이 좋다.


여담으로 MS office는 [서버-side 개발을 권장하지 않는다.](https://support.microsoft.com/en-us/help/257757/considerations-for-server-side-automation-of-office?wa=wsignin1.0%3Fwa%3Dwsignin1.0) -by 이석원 프로님


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

## exsel activity
### exsel application scope 사용 안함
Read Range : 시스템.파일.통합문서.Read Range
 - 간단하게 DataTable 만들 때 사용.

Write Range : 시스템.파일.통합문서.Write Range
 - 수식 자동완성 안됨. "=SUM(A1:B1)"입력시 문자열 그대로 들어감.

데이터 조작이 필요한 작업을 할 때는 어플스코프 사용 권장함.

### exsel application scope 사용
- 엑셀 파일을 열어서 작업을 진행함.

Read Range : 앱통합.Excel.테이블.Read Range
- exsel application scope에서 작업 중인 파일의 데이터를 읽음
- 반드시 header가 체크 되어있는지 확인할 것 (데이터 row가 밀려쓰기 될 수 있다.)

Write Range : 앱통합.Excel.테이블.Write Range
- 수식 자동완성 됨. "=SUM(A1:B1)"입력시 상대위치를 통해 수식이 적용됨.

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
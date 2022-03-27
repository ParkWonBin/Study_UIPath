### 엑셀 읽기 오류 관련
[UIPATH 엑셀 StacOverFlow](https://stackoverflow.com/questions/2424717/how-to-know-if-a-cell-has-an-error-in-the-formula-in-c-sharp)  
[UIPATH 엑셀 오류 정리글](https://deokpals.tistory.com/11)  

```yaml
Excel ReadRange - Value of Errors: 
  - ErrNA : -2146826247,
  - ErrNum : -2146826253,
  - ErrRef : -2146826266,
  - ErrNull : -2146826289,
  - ErrName : -2146826260,
  - ErrDiv-0 : -2146826281,
  - ErrValue : -2146826274
  - ErrGettingData : -2146826246,
```


### Excle 함수
```yaml
조건부 서식에 쓰기 좋은 함수:
  - =Not(IsBlank($A1))        # 해당 행의 A열의 값이 비어있지 않으면 서식 적용
  - =ISNUMBER(SEARCH("1",A1)) # 해당 위치에 있는 문자열이 '1'을 포함하고 있으면 True

Excel Template 만들 떄 좋은 함수 : 
  - =TEXTJOIN(",",TRUE, B3,B4,...) # 빈셀무시=True, 구분자=",", String Join
  - =vlookup()
  
Pivot Table 설정 : 
  - 단축키 : Alt+D + p
  - 경로 : 리본>삽입>피벗테이블
  - 옵션 : 새 시트로 생성 > [필터, 열 레이블, 행 레이블, 값] 설정
  - 수정 : 테이블 우클릭 > 피벗테이블 필드 표시
```

### Excel Activities 사용 관련 
[VBA Excel Range -> HTML](https://stackoverflow.com/questions/54033321/excel-vba-convert-range-with-pictures-and-buttons-to-html)

```yaml
공백문자 관련 : 
  - 공백문자 생성 : 셀 입력창 > alt+enter
  - 공백문자 종류 : 줄바꿈 문자는 chr(10)이다. (chr(13)="\n"과 다른 문자다.)
  - 배열 만들기 : 
    - str_test.split(chr(10)) 
    - str_test.split(Environment.NewLine.ToArray)

Workbook Activity : 
  - 장점 : MS office Excel이 다운로드 되어 있지 않아도 사용이 가능하다.  
  - 특징 : Write Range로 "=SUM(A1:B1)"입력 시 문자열로 해당 값이 입력된다. 

Excel application scope :
  - 실제로 Excel 프로그램을 백그라운드에서 실행하여 작업을 시도한다.  
  - MS office Excel 이 다운로드 되어 있는 컴퓨터에서만 사용이 가능하다.  
  - 수식("SUM(A1:B1)")을 입력할 경우 해당 수식의 값이 계산되어 입력된다. 
  - DRM이 설치된 환경이라면 Kill Process와 함께 쓰는 것을 권장한다.
  - 엑셀이 이미 실행되고 있는 경우 에러를 발생하기 때문이다.

```


### 엑셀 Activities Tip
```yaml
시트명 갖고오기:
    - excel scope > output : wb As WorkBook
    - StrArr_SheetNames = wb.GetSheets

Read Range 확인사항 : 
    - header가 체크여부 확인

Uipath > Excel에 CopyPasteRange : 
    - 장점 : 병합된 셀의 서식까지 모두 붙여넣어진다.
    - 주의 : WriteRange 시 Data 병합 여부는 제대로 갱신되지 않음

UiPath > Refresh Pivot Table  :
    - 피봇 테이블 사용 시 Read Range / Save WorkBook 전 확인요망.

```

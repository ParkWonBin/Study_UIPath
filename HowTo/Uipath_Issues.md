# 특이사항
알고 싶지 않았으나 알게된 정보를 정리합니다.   

#### UiPath 라이브러리 올릴 떄 참고
```yaml
주석관련 : 
  - Activity에 커서올리면 바로 볼 수 있는 주석 설정하기 :
    - 방법 : Project패널 > 해당 xaml 우클릭 > Properties > tooltip 설정 > Save
    - 비고 : 최상위 activity에 주석 넣어놓으면, tooltip 설정할 때 자동으로 주석 내용을 넣어준다.
             tooltip 을 적용하면 xaml에 <sap2010:Annotation.AnnotationText> 항목이 생긴다. 
에러관련 : 
  - 배포 후 NameSpace 컴파일 에러 : 
    - Invoke Code에서 liblaray 함수/객체 명을 strict하게 쓰지 않았을 때 주로 생김 (특힉 workbook)
    - 다른 라이브러리와 객체/함수 명이 같아 생기는 오류. 함수/객체 호출 시 System부터 쭉 경로 다 써줘야 예방 가능
  - 배포 후 잘못된 형식 관련 에러 : 
    - Library 안에서 State Machine 사용했을 떄, 트랜젝션 Triger에 Element Exist 등 넣어두면 해당 오류 생김.
    - 정확한 원인은 모르겠으나. 트렌젝션에서 처리하는 내용을 entry나 exist 마지막 부분으로 이동시키면 해결 가능.
```

#### Kill Process 액티비티
```yaml
Description : 
  UiPath.Core.System.Activity 버전이 18.4 -> 19.4 로 변경될 때 kill process의 대상이
  이미 꺼져있는 경우(process 목록에서 찾지 못한 경우) 에러를 발생하도록 변경되었다.
ToDo:
  이전과제에 적용된 Kill Process를 모두 찾아 ContinueOnError=True를 적용하여 해결
After : 
  kill process 를 Uipath 액티비티로 사용하기 보다는 Custom Activity나 library를 만들어 사용하는 게 좋을 것
```

## cmd로 Uipath 실행 관련
[스케줄러 참고](https://deje0ng.tistory.com/78)
[UiPath 문서 참고](https://docs.uipath.com/robot/docs/arguments-description)
```cmd 
# 해당 방법은 Uipath 2019.10 버전 이하에서 작동했던 방법입니다.
# 2020 이후 버전부터는 Uipath Assistant를 경유해서 동작하도록 업데이트 되었습니다.
# Ubot 라이센스를 구매하지 않고 MS스케줄러를 사용하는 편법을 막기 위한 조치로 보입니다. 

# KS출근.bat
cd "<C:\Program Files (x86)\UiPath\Studio\>"
.\UiRobot.exe execute  -process "KS출근" -input "{ 'str_workTime' : '8-17' , 'str_ID' : 'wbpark' }"
```

##### Invoke Code 사용할 경우 MethodName 의 경우 대소문자를 구분 필수
"Add"로 써야할 것을 "add"로 쓸 경우 에러가 발생한다.

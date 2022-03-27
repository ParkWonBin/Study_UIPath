
####  작업할 때 참고할 것

```vb
' Foreground
Microsoft.VisualBasic.Interaction
System.Windows.Forms
' System
System.IO
System.Linq
System.File
```

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

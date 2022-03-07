# 특이사항
알고 싶지 않았으나 알게된 정보를 정리합니다.   

## [시작] 메뉴 패널 안열릴 떄 window app 사용하기
시작버튼, window 키를 눌러도 패널이 열리지 않을 때 Window 기본 app 사용 방법  
윈도우 탐색기가 열린다면 시스템 경로 이동하여 수동으로 app 실행이 가능합니다.  

| 이름 | 경로,이름 | 
|:---:|---|
경로이동 | C:\Windows\System32
원격접속 | mstsc.exe
그림판 | mspaint.exe

## System32 경로 내 존재하지 않는 앱이 System32 경로에서 실행되는 경우
| 폴더명 | 경로 위치 | 
|:---:|---|
System32 | C:\Windows\System32
SysWOW64 | C:\Windows\SysWOW64

###### 출처: [몽키의 IT개발 노트](https://pung77.tistory.com/23) 
```txt
32bit 프로세스는 SysWOW64 시스템 폴더의 내용을 참조하여 실행된다.
64bit 프로세스는 System 32 시스템 폴더의 내용을 참조하여 실행된다.

Windows는 File System Redirector라는 기능을 지원하여,   
Sytem32 폴더에 접근하여 사용하려 하여도 SysWOW64 폴더로 리다이렉트 시켜 자동으로 SysWOW64 폴더의 내용을 참조한다.   
https://msdn.microsoft.com/ko-kr/library/windows/desktop/aa384187(v=vs.85).aspx   

예를들어 64bit os에서 
32bit 프로세스가  LoadLibrary(C:\windows\System32\Kernel32.dll) 을 호출하여 Kernel32.dll을 로딩하려 하여도 
실제로는 리다이렉트되어 C:\windows\SysWOW64\Kernel32.dll 경로의 Kernel32.dll을 참조한다.

32bit 프로세스가 System32 폴더에 접근하고 싶다면 Wow64EnableWow64FsRedirection API를 사용해 리다이렉트 기능을 끄고 강제로 접근하면된다.  
https://msdn.microsoft.com/ko-kr/library/windows/desktop/aa365744(v=vs.85).aspx
```


##### 동일한 문자열이 고장날 떄
눈에 보이지 않는 공백문자가 포함되어 있을 수 있다.  
글자의 가로 길이가 0인 특수문자가 포함되어 있는 문자다.  

```vb
A = "C:\Users\H2109941\Desktop\tmp\test.xlsx"
B = "‪C:\Users\H2109941\Desktop\tmp\test.xlsx"

System.IO.File.Exists(A) | True | window 탐색기 주소창에서 복사
System.IO.File.Exists(B) | False| 해당 파일 속성>보안>개체이름 에서 왼쪽에서 오른쪽으로 드레그하여 복사

A.Length | 39
B.Length | 40 | 글자의 폭이 0 인 문자가 끼어있음

```

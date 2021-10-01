: [ README ] :::::::::::::
: 사용법       : 실행시키고자 하는 VB파일의 경로를 targetPath에 등록합니다.
: 주의사항1   : set 문장이 "=" 좌우에 공백문자가 들어가면 안됩니다. 
: 주의사항2   : set 문장 마지막에 글자에 공백문자가 들어가면 안됩니다.
: 주의사항3   : 경로를 입력할 때 따옴표 " 를 입력하면 안됩니다.
: 주의사항4   : 변수 호출은 %변수명%으로 이루어지며, 문자열이 replace되어 명령으로 전달되는 느낌입니다.
:::::::::: [ Cmd 명령 구간 ]::::::::::::::

::: VB script 위치 입력
set targetPath=C:\test.vb

::: Window에 기본적으로 내장된 VB Compiler 입력
set vbcCompiler=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\vbc.exe

::: VB Compiler로 script 실행
%vbcCompiler% %targetPath%

::: VB Compiler가 만든 execute 파일 실행
%targetPath:~0,-2%exe

::: Console.WriteLine 을 확인하기 위해 CMD창 멈춤
PAUSE 

:::::::::: [ 주석 : CMD 문법 ] :::::::::::
: echo %targetPath%            ::: 그냥 변수 호출
: echo %targetPath:~0,4%      ::: substring(0,4)
: echo %targetPath:~-2%       ::: 뒤에서 2글자만 출력
: echo %targetPath:~0,-2%     ::: 처음부터 출력, 뒤에서 2글자 삭제
: %FolderPath%\%FileName%      ::: 문자열 합치기 = 그냥 이어서 쓰면 됨
:
:::::::::::::::::: [ 참고문헌 ] ::::::::::::::::::
: cmd 문법 관련
: https://m.blog.naver.com/PostView.naver?isHttpsRedirect=true&blogId=novajini&logNo=220158528197
:
: substring 관련
: https://m.blog.naver.com/PostView.naver?isHttpsRedirect=true&blogId=21centmen&logNo=220610919732
:
: 시간지연 
: https://trustall.tistory.com/32

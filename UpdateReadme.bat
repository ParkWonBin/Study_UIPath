echo off

::: 같은 이름의 VB파일로 경로명 변경
set targetPath=%0
set targetPath=%targetPath:~1,-4%vb

echo
echo [ Run VB File ]
echo  %targetPath%
echo

::: Window에 기본적으로 내장된 VB Compiler 입력
set vbcCompiler=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\vbc.exe

::: VB Compiler로 script 실행
%vbcCompiler% %targetPath%

::: VB Compiler가 만든 execute 파일 실행
%targetPath:~0,-2%exe

::: Console.WriteLine 을 확인하기 위해 CMD창 멈춤
PAUSE 
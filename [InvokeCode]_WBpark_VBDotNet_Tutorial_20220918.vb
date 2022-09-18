'# UiPath Invoke Code 사용 팁
'### System 외 다른 라이브러이 사용할 떄 binding 오류 생기면 CType으로 중간중간 형변환 해줘야 한다.
'### Runtime Binding 오류발생 시 조치방법 : 해당 라인에서 변수들마다 Ctype으로 형변환 적용할 것
'### xaml파일의 레퍼런스 참조와 충돌을 피하기 위해 System, Microsoft 등 소스코드에 해당 라이브러리 전체 경로 표시 추천
'==========================================================
'# 조건문
'==========================================================
Dim Str_TestCase As String = "Case1"

'## 삼항 연상자
Console.WriteLine( If( Str_TestCase.Equals("Case1"), "Scenario : True", "Scenario : False") )

'## IF
If Str_TestCase.Equals("Case1")  Then
	Console.WriteLine("IF.Scenario : Case1")
ElseIf Str_TestCase.equals("Case2")
	Console.WriteLine("IF.Scenario : Case2")
Else 
	Console.WriteLine("IF.Scenario : Else")
End If

'## Select Case (Switch) 예시1 | 값일치
Select Case Str_TestCase
	Case Is = "CASE1"
		Console.WriteLine("SelectCase.Senario : CASE1")
	Case Is = "CASE2"
		Console.WriteLine("SelectCase.Senario :CASE2")
	Case Is = "CASE3"
		Console.WriteLine("SelectCase.Senario :CASE3")
	Case Else
		Console.WriteLine("SelectCase.Senario :Else")
End Select

'## Select Case (Switch) 예시2 | 범위
Dim Int_TestCase As Integer = 1
Select Case Int_TestCase
	Case 1 To 5
		Console.WriteLine("SelectCase.Senario : 1~5 사이 값")
	Case 6, 8, 10
		Console.WriteLine("SelectCase.Senario : 6, 8, 10 중 하나")
	Case Is >=11 , <=20
		Console.WriteLine("SelectCase.Senario : 11과 20사이 숫자")
	Case Else
		Console.WriteLine("SelectCase.Senario : Else")
End Select
'=========================================================
' # [Loop] For , For Each 
' For , ForEach 끝날 때 Next 뒤에 로컬변수 명시 권장.
' 문법적으로 Next 뒤에 변수 안써줘도 되지만. 코드 가독성을 위해 권장
'----------------------------------
' ## integer, TO 사용하여 범위 지정
For int_i As Integer = 0 To 3
	console.writeline("1.For: int_i = "+int_i.tostring)
Next int_i
' 0 1 2 3
'----------------------------------
' ## integer, TO, Step 사용하여 범위 지정
For int_i As Integer = 3 To 0 Step -1
	console.writeline("2.For: int_i = "+int_i.tostring)
Next int_i
' 3 2 1 0
'----------------------------------
' ## integer, Array 사용하여 범위 지정
For Each  int_i As Integer In {1, 3, 5}
	console.writeline("3.For Each: int_i = "+int_i.tostring)
Next int_i
' 1 3 5
'=========================================================
' # [Loop] While, Do While, Until
Dim int_WhileCnt As Integer
Dim int_DoWhileCnt As Integer
Dim int_UntilCnt As Integer
'=========================================================
int_WhileCnt = 0
While int_whileCnt < 3
	int_whileCnt = int_whileCnt +1
 	console.writeline("1.While: int_whileCnt<3 | "+int_whileCnt.tostring)
End While
' 1 2 3
'----------------------------------
int_DoWhileCnt = 0
Do While int_DoWhileCnt < 3
	int_DoWhileCnt = int_DoWhileCnt +1
 	console.writeline("2.Do While: int_whileCnt<3 | "+int_DoWhileCnt.tostring)
Loop
' 1 2 3
'----------------------------------
int_DoWhileCnt = 0
Do 
	int_DoWhileCnt = int_DoWhileCnt + 1
 	console.writeline("3.Do While: int_DoWhileCnt < 3 | "+int_DoWhileCnt.tostring)
Loop While int_DoWhileCnt < 3
' 1 2 3
'----------------------------------
int_UntilCnt = 0
Do Until int_UntilCnt >= 3
	int_UntilCnt = int_UntilCnt + 1
 	console.writeline("4.Do Until: int_UntilCnt < 0 |"+int_UntilCnt.tostring)
Loop 
'1 2 3
'----------------------------------
int_UntilCnt = 0
Do 
	int_UntilCnt = int_UntilCnt + 1
 	console.writeline("5.Do Until: int_UntilCnt < 0 |"+int_UntilCnt.tostring)
Loop Until int_UntilCnt >= 3
'1 2 3
'=========================================================
' Try Catch 관련
' https://docs.microsoft.com/ko-kr/dotnet/visual-basic/language-reference/statements/try-catch-finally-statement
'=========================================================
Try
   Console.WriteLine("[Try] Make some exception")
   Throw New Exception("Exception Message")
   Throw New SystemException("SystemException Message")
   Throw New ApplicationException("ApplicationException Message")
   Throw New NullReferenceException("NullReferenceException Message")
   Throw New IndexOutOfRangeException("IndexOutOfRangeException Message")
   Throw New ArgumentException("ArgumentException Message")
   Throw New ArgumentNullException("ArgumentNullException Message")
   Throw New ArgumentOutOfRangeException("ArgumentOutOfRangeException Message")
   Throw New ExternalException("ExternalException Message")
Catch ex As Exception
	Console.WriteLine("[Catch] " & ex.Message & vbCrLf & "Stack Trace: " & vbCrLf & ex.StackTrace)
Finally
    Console.WriteLine("[Finally] end try")
End Try
'-----------------------------------------------
Dim int_try_i As Integer  = 5
Try
    Throw New ArgumentException()
Catch e As OverflowException When int_try_i = 5
    Console.WriteLine("First handler")
Catch e As ArgumentException When int_try_i = 4
    Console.WriteLine("Second handler")
Catch When int_try_i = 5
    Console.WriteLine("Third handler")
End Try
'=========================================================
' List of key-value pairs 생성
Dim list_KeyValuePair As List(Of KeyValuePair(Of String, Integer)) =New List(Of KeyValuePair(Of String, Integer))
list_KeyValuePair.Add(New KeyValuePair(Of String, Integer)("dot", 1))
list_KeyValuePair.Add(New KeyValuePair(Of String, Integer)("net", 2))
list_KeyValuePair.Add(New KeyValuePair(Of String, Integer)("Codex", 3))
' List of key-value pairs 값 호출
For Each pair As KeyValuePair(Of String, Integer) In list_KeyValuePair
	Console.WriteLine("KeyValuePair : key={0}, Value={1}", pair.Key, pair.Value)
Next
'=========================================================
' 함수와 SUB 선언 및 호출에 대하여
' Invoke Code 액티비티 내 Sub,Function 생성 불가하여,
' Action이나 Func변수에 Lambda함수, Lambda프로시저 를 할당하고 호출.
' Lambda 함수/프로시저 에는 Optional 변수를 사용할 수 없습니다.
'=========================================================
'test. 인수 없는 함수 한줄로 선언 및 호출
Dim func_test As System.Func(Of String) = Function() "test"
 Console.WriteLine(func_test)   'out: "System.Func`1[System.String]"
 Console.WriteLine(func_test()) 'out: "test"

'1. Func 인수 없는 함수
Dim func_A As System.Func(Of String) = Function() As String
    Return "func_A : No inArgument"
End Function
 Console.WriteLine(func_A())  'out: "func_A : No inArgument"

'2. Func 인수 1개 함수
Dim func_B As System.Func(Of String, String) = Function(x As String) As String 
     Return "func_B : "+ x
End Function
Console.WriteLine(func_B("First Argumnet"))  'out: "func_B : First Argumnet"

'3. Func 인수 1개 함수
Dim func_C As System.Func(Of String, String, String) = Function(x As String, y As String) As String 
     Return String.format("func_C : x={0}, y={1}",x,y)
End Function
Console.WriteLine(func_C("1","2")) 'out: "func_B : x=1, y=2"
'----------------------------------------------
'test1. 인수 없는 프로시저 한줄로 선언 및 호출
Dim Action_test_1 As System.Action = Sub() Console.WriteLine("Action_test_1")
Action_test_1() ' out: "Action_test_1"

'test2. 다른 프로시저를 받아서 호출
Dim Action_test_2 As System.Action(Of String) = AddressOf System.Console.WriteLine
Action_test_2("Action_test_2") ' out: "Action_test_2"

'1. 인수 없는 프로시저
Dim Action_A As System.Action = Sub() 
	Console.WriteLine("Action_A : No inArgument")
End Sub
Action_A() ' out:Action_A : No inArgument

'2. 인수 1개 프로시저
Dim Action_B As System.Action(Of String) = Sub(x As String) 
	Console.WriteLine("Action_B : {0}", x)
End Sub
Action_B("First Argument")

'3. 인수 2개 프로시저
Dim Action_C As System.Action(Of String, String) = Sub(x As String, y As String) 
	Console.WriteLine("Action_C : x={0} y={1}", x, y)
End Sub
Action_C("1","2")

'=========================================================
' Function, Action 실전 활용법
'=========================================================
' 특정 함수가 실행될 때마다 로그 찍기
Dim Act_logger As System.Action(Of String, System.Action(Of String), String) = Sub( _Marker As String, _Act As Action(Of String), _arg As String)
	Dim _startTime As System.DateTime = System.DateTime.Now
    console.writeLine(_marker+" 작업 시작!")
	_Act(_arg)	
	console.writeLine(_marker+" 작업 종료! 소요시간 : "+ Now.Subtract(_startTime).TotalSeconds.ToString("0.000 초"))
End Sub
Act_logger("[101] 모듈 | Line 1", AddressOf console.writeLine, "test")
'-------------------------------
' Retry 로직을 적용하여 성공여부를 변수로 받아볼 수 있도록 설정함.
Dim Fnc_RetryProcess As System.Func(Of String, String) = Function( _Marker As String ) As String
    Dim Bln_ContinueOnError As Boolean = False
	Dim Bln_is_Succeed As Boolean = False
	Dim Int_RetryMax As Integer = 5
	Dim Int_RetryCnt As Integer = 0
	Dim Str_outMsg As String = ""
	While Not Bln_is_Succeed AndAlso Int_RetryCnt < Int_RetryMax
		Try 
			Dim _startTime As System.DateTime = System.DateTime.Now
			console.writeLine(_Marker+" 작업 시작" +If(Int_RetryCnt>0 ," | 재시도 "+Int_RetryCnt.ToString,"") )
			'Process Start---------
			'Do Something 
			Bln_is_Succeed = True ' 작업 성공 여부 체크 조건 설정
			'------------------------
			console.writeLine(_Marker+" 작업 종료 | 소요시간 : "+ Now.Subtract(_startTime).TotalSeconds.ToString("0.000 초"))
			Str_outMsg = If(Bln_is_Succeed,"Succeed",Str_outMsg)
		Catch e As System.Exception
			Str_outMsg = String.format("{1}{0}{2}{0}{3}",vbnewline, e.TargetSite.tostring, e.message, e.source.tostring)
			console.writeLine(Str_outMsg)
		Finally
			Int_RetryCnt = Int_RetryCnt+1
		End Try
	End While
	
	If Not Bln_is_Succeed AndAlso Not Bln_ContinueOnError Then
		' 프로세서 수행에 실패한 경우
		Throw New System.ApplicationException(Str_outMsg)
	End If 
	
	Return Str_outMsg '성공 시 "Succeed", 실패 시 에러문구 들어있음
End Function

' 해당 함수 성공여부 추출
console.writeLine( Fnc_RetryProcess("test") )

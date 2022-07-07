'1. Func 한줄 선언. 자료형 명시X
Dim func_A As System.Func(Of String, String) = Function(x) "func_A : "+ x

'2. Func 한줄 선언. 자료형 명시O
Dim func_B As System.Func(Of String, String) = Function(x As String) "func_B : "+ x

'3. Func 블럭 선언. 자료형 명시X
Dim func_C As System.Func(Of String, String) = Function(x) 
	Return "func_C : "+ x
End Function

'4. Func 블럭 선언. 자료형 명시O
Dim func_D As System.Func(Of String, String) = Function(x As String) As String 
	Return "func_D : "+ x
End Function

'5. Action 한줄 선언. 자료형 명시X
Dim Action_A As System.Action(Of String) = Sub(x) Console.WriteLine("Action_A : {0}", x)

'6. Action 한줄 선언. 자료형 명시O
Dim Action_B As System.Action(Of String) = Sub(x As String) Console.WriteLine("Action_B : {0}", x)

'7. Action 블럭 선언. 자료형 명시X
Dim Action_C As System.Action(Of String) = Sub(x) 
	Console.WriteLine("Action_C : {0}", x)
End Sub

'8. Action 블럭 선언. 자료형 명시O
Dim Action_D As System.Action(Of String) = Sub(x As String) 
	Console.WriteLine("Action_D : {0}", x)
End Sub

Console.WriteLine(func_A("test1"))
Console.WriteLine(func_B("test2"))
Console.WriteLine(func_C("test3"))
Console.WriteLine(func_D("test4"))
Action_A("test5")
Action_B("test6")
Action_C("test7")
Action_D("test8")
        
        'UiPath Invoke Code는 입력받은 문자열을 EVAL 해서 sub으로 실행시켜준다.
'따라서 InvokeCode 안에서 VB.NET을 쓸 떄 특이한 제약이 많이 생긴다.
'(1). sub 생성 불가, 단, Action으로 생성은 가능하다.
'(2). System 외부의 객체(VBA등)를 다룰 때는 script가 Type구분을 잘 못한다. 
    'System 외부의 객체 사용시 매소드,맴버를 호출할 떄는 CType 필수.(런타임 바인딩 오류, 형변환 오류)
'(3). 어셈블리 참조 신경쓸 것(System Or Microsoft에서부터 쭉 쓰는거 권장 ㅠㅠ)

' [VBA <-> VB.NET 코드 변환시 유의사항]
'(1) Runtime Binding 오류발생 시 조치방법 : OLEObject 매소드(Add)를 사용하기 전에 Ctype 적용할 것
'(2) VBA에서는 줄바꿈 할 때 '_'가 강제되지만 VB.NET은 선택사항
'(3) VB.NET에서는 선언과 동시에 초기화 하는 문법이 사용 가능하지만 [Dim ~ As = ~] VBA에서는 불가
'(4) VBA에서 기본적으로 사용하는 True/False는 Microsoft.Office.Core.MsoTriState 에 정의된걸 써야 인식가능
'(5) VBA함수가 받는 숫자 중에는 CType으로 Single로 형변환해야 들어가는 것들 있음
'(6) VBA객체의 매소드를 사용할 시, 매소드 호출 전 해당 객체를 CType으로 묶어줘야 호출이 가능함. (RunTime Binding은 Ctype으로 대부분 해결가능)
'(7) VBA객에에서 특정 매소드/맴버가 없다는 오류문구가 UiPath에서 나는 경우. Import 패널에 최대한 가까운 어셈플리 Import 해주면 해결될 때 있음.

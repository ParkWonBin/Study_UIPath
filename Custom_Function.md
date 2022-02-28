# Custom Function
#### Uipath 에서 사용자 정의 함수 만드는 방법
### 1. invoke code를 통해 함수를 정의하고 out으로 반환하여 사용한다. 
- Edit Arguments 통해서 정의한 함수를 밖으로 출력한다.
- 자료형은 System.Func<T,...TResult> 로 되어있는 것으로 한다.
- T의 경우 인수로 입력받을 변수의 Type을 지정하는 설정이며, 마지막 TResult는 함수의 반환값의 Type이다.

```vb
Fnc_test = Function(str_test As String) As String()
 Return "input : " & str_test
End Function
```


### 2. invoke code 내에서 정의한 함수를 해당 invoke code 내에서 재사용 한다.
```vb
Dim Fnc_test As System.Func(Of String, String()) = Function(str_test As String) As String()
 Return "input : " & str_test
End Function

console.WriteLine( Fnc_test("Test") )
```

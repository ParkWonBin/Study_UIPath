### 함수 만들기 요렁

```yaml
Uipath 에서 사용자 정의 함수 만들기 요령 :
- Inovke Code를 통해 VB로 정의한 함수를 Out 변수로 뺴낸다. 
- 함수의 자료형은 System.Func<T,...TResult> 
- 인수 설명 : 
  - T : 인수로 입력받을 변수의 Type
  - TResult : 함수의 반환값의 Type
```

```vb
Fnc_test = Function(str_test As String) As String
 Return "input : " & str_test
End Function
```

### 람다 함수식에 변수 넣어주기
```vb
' 그냥 람다식에 () 치고 바로 뒤에 (인수) 넣어주면 됨.
Console.WriteLine( ( (Function(num As Integer) num + 0)(5) ).ToString ) '=> 6
```

## 내 함수 목록

```vb
Dim Fnc_Get_All_Files As System.Func(Of String, String()) = Function(str_path As String) As String
 Dim list_str_dir As New System.Collections.Generic.List(Of String) 
 Dim list_str_file As New System.Collections.Generic.List(Of String)

 list_str_dir.Add(str_path)
 While list_str_dir.Count <> 0
  str_path = list_str_dir.Last
  list_str_dir.RemoveAt(list_str_dir.Count -1)
  list_str_dir.AddRange(System.IO.Directory.GetDirectories(str_path))
  list_str_file.AddRange(System.IO.Directory.GetFiles(str_path))
    End While
  Return list_str_file
End Function
```


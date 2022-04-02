### Fnc_UI_Get_DirPath.vb
```vb
Dim Fnc_UI_Get_DirPath As System.Func(Of String, String) = Function(str_Desc As String) As String
  '2022.04.02|wbpark|폴더 선택용 UI
  Dim Dlg_FolderBrowser As System.Windows.Forms.FolderBrowserDialog = New System.Windows.Forms.FolderBrowserDialog() 
  Dlg_FolderBrowser.Description = str_Desc
  Dlg_FolderBrowser.ShowNewFolderButton = True
  Dlg_FolderBrowser.RootFolder = Environment.SpecialFolder.MyComputer
  
  Dim result As System.Windows.Forms.DialogResult = Dlg_FolderBrowser.ShowDialog()
    If result = System.Windows.Forms.DialogResult.OK Then
    Return Dlg_FolderBrowser.SelectedPath
  End If 
  Return ""
End Function
```

### Fnc_Ui_MsgBox.vb
```vb
Dim Fnc_Ui_MsgBox As System.Func(Of String, String, Boolean) = Function(Str_Massege As String, Str_title As String) As Boolean
  '2022.04.02|wbpark|message와 title을 입력받아 확인/취소 여부를 bool로 입력받습니다.
  Dim Result As System.Windows.Forms.DialogResult = System.Windows.Forms.MessageBox.Show(Str_Massege, Str_title, MessageBoxButtons.YesNo)
  Return If(Result = System.Windows.Forms.DialogResult.Yes,True,False)
End Function
```

### Fnc_Get_All_Files.vb
```vb
Dim Fnc_Get_All_Files As System.Func(Of String, String()) = Function(str_path As String) As String()
  '2022.04.02|wbpark|숨김파일 및 운영체제에 의해 숨겨진 파일까지 모두 파악.
  Dim StrList_file As New System.Collections.Generic.List(Of String)
  Dim StrList_dir As New System.Collections.Generic.List(Of String)
 
  StrList_dir.Add(str_path)
  While StrList_dir.Count <> 0
    str_path = StrList_dir.Last
    StrList_dir.RemoveAt(StrList_dir.Count -1)
    StrList_dir.AddRange(System.IO.Directory.GetDirectories(str_path))
    StrList_file.AddRange(System.IO.Directory.GetFiles(str_path))
  End While
  Return StrList_file.ToArray
End Function
```

### Fnc_Text_Replace.vb
```vb
Dim Fnc_Text_Replace As System.Func(Of String, String(), String(),String) = Function(Str_Source As String, Replace_Before As String(), Replace_After As String()) As String
  '2022.04.02|wbpark|입력된 문자열 일괄 replace
  For i As Integer = 0 To Replace_Before.length-1
    Str_Source=Str_Source.replace(Replace_Before(i),Replace_After(i))
  Next 
  Return Str_Source
End Function
```

### Fnc_Regex_Replace_Selector.vb
```vb
Dim Fnc_Regex_Replace_Selector As System.Func(Of String, String, String) = Function(Str_Source As String, Str_ptn As String) As String
  ' 2022.04.02|wbpark|셀렉터 내 변수 부분 {{}}로 바꿔주기
  For Each x As System.Text.RegularExpressions.Match In System.Text.RegularExpressions.Regex.Matches(Str_Source,Str_ptn) 
    If Not (x.Tostring.ToUpper.Contains(".TOSTRING") OrElse x.ToString.Contains("(") OrElse  x.ToString.Contains(")") ) Then
      Dim Str_replaceTo As String = x.Tostring.replace("[&quot;","").replace("&quot;+","{{").replace("+&quot;","}}").replace("&quot;]","")
      Str_Source=Str_Source.replace(x.tostring,Str_replaceTo)
    End If
  Next
  Return Str_Source
End Function
```

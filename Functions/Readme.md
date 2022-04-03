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

### App_Convertion_IE_TO_Edge.vb
```vb
Dim App_Convertion_IE_TO_Edge As System.Func(Of String(), String(), String) = Function(StrArr_Before As String(), StrArr_After As String()) As String
  '2022.04.04|wbpark|UiPath UIAutomation=21.4를 기준으로 IE 셀렉터를 Edge로 변환합니다.
  'UIAutomation의 버전상이할 경우, 버전이 높든 낮든 상관 없이 다른 버전에서 Validate 한 셀렉터를 제대로 인식하지 못하는 문제가 있습니다.  

  ' ui 작업
  Dim Str_Desc_msg As String = String.format("###########{0}UIAutomation=21.4를기준으로 동작합니다{0}패키지 업데이트는 수동으로 진행하셔야 합니다.(버전 확인 필수){0}파일명에 %20이 있는 경우 ' '으로 바꿔줍니다.(Nuget에서 추출한 경우){0}###########{0}{0}Attach Browser의 BrowserType을 IE에서 Edge로 바꿉니다.{0}Selector 내 변수가 사용된 경우 {1} {2}로 바꾸어줍니다.{0}(셀렉터 '.tostring','(',')' 미포함 시에만 작동){0}{0}before -> After{0}",VbNewLine,"{{","}}")
  Fnc_Ui_MsgBox(Str_Desc_msg+Join(StrArr_Before.Select(Function(x,i) x+" => "+StrArr_After(i) ).ToArray,VbNewLine),"App 설명")
  Dim Str_Dir_Source As String = Fnc_UI_Get_DirPath("IE=>Edge 변환할 Project의 폴더 선택")
  Dim Str_Dir_Result As String = Str_Dir_Source+"_Replaced"
  
  ' 파일 작성
  Dim int_Result_Cnt As Integer = 0
  Dim Str_Error_Log As String = ""
  For Each Str_FilePath As String In Fnc_Get_All_Files(Str_Dir_Source)
    Try
      '결과 폴더를 따로 만들기
      Dim Str_File_ResultPath As String = Str_FilePath.replace(Str_Dir_Source,Str_Dir_Result)    
      System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(Str_File_ResultPath))
      
      'xaml은 일괄적으로 변경 후 저장
      If System.IO.Path.GetExtension(Str_FilePath).ToUpper = ".XAML"
        ' 파일 읽기 및 Replace
        Dim Str_FileContent As String = System.IO.File.ReadAllText(Str_FilePath)
        Str_FileContent = Fnc_Regex_Replace_Selector(Str_FileContent,"Selector=""\[[^]]+\]""")
        Str_FileContent = Fnc_Text_Replace(Str_FileContent,StrArr_Before,StrArr_After)
        
        ' 파일 작성
        System.IO.File.WriteAllText(Str_File_ResultPath.replace("%20"," "), Str_FileContent, System.Text.Encoding.UTF8)
        System.Console.WriteLine(int_Result_Cnt.Tostring("00")+" | "+Str_File_ResultPath)
      Else
        System.IO.File.Copy(Str_Dir_Source, Str_File_ResultPath)
      End If
    Catch ex As Exception
      Str_Error_Log=Str_Error_Log+Str_FilePath+" | "+ex.message+vbnewline
      console.writeline(Str_FilePath+vbnewline+ex.message)
    End Try
  Next
  System.Diagnostics.Process.Start("explorer.exe", Str_Dir_Result)
  return Str_Error_Log
End Function
```

### App_Update_Functions_And_Readme.vb
```vb
Dim App_Update_Functions_And_Readme As System.Func(Of String, String) = Function(Str_DirPath As String) As String
  '2022.04.04|wbpark|Readme갱신 및 Function 기록관리
  
  ' 입력값 미존재 시 UI 입력 받기
  if String.isNullOrWhiteSpace(Str_DirPath) Then
    Str_DirPath = Fnc_UI_Get_DirPath("App 폴더 선택 : "+vbNewLine+"해당 폴더 내 vb파일을 읽고 Custom 함수를 기록관리 합니다.")
  End if

  ' Extract Data
  Dim Str_SavePath_Fnc As String = System.IO.Path.combine(System.IO.Path.GetDirectoryName(Str_DirPath),"Functions")
  Dim Dic_ReadMe As New System.Collections.Generic.Dictionary(Of String, String)
  Dim Dic_Ref As New System.Collections.Generic.Dictionary(Of String, System.Collections.Generic.List(Of String))
  Dim Str_Error_Log As String = ""

  ' Write File
  System.IO.Directory.CreateDirectory(Str_SavePath_Fnc)
  For Each Str_AppPath As String In Fnc_Get_All_Files(Str_DirPath).Where(Function(x) System.IO.Path.GetExtension(x).ToLower=".vb")
    Dim Str_AppContent As String = System.IO.File.ReadAllText(Str_AppPath)
    Dim StrArr_Contents As String() = System.Text.RegularExpressions.Regex.Split(Str_AppContent,"'"+"-+").Where(Function(x) x.Contains("End "+"Function")).ToArray
    Dim StrArr_FileNames As String() = StrArr_Contents.Select(Function(x) x.Substring(x.IndexOf("Dim ")+4,x.IndexOf(" As ")-x.IndexOf("Dim ")-4)+".vb").ToArray
    For int_i As Integer = 0 To StrArr_FileNames.length-1
      Try
        Dim Str_SaveFilePath As String = System.IO.Path.Combine(Str_SavePath_Fnc,StrArr_FileNames(int_i))
        System.IO.File.WriteAllText(Str_SaveFilePath, StrArr_Contents(int_i).Trim)
        Dic_ReadMe(StrArr_FileNames(int_i)) = StrArr_Contents(int_i).Trim
        If Dic_Ref.keys.Contains(StrArr_FileNames(int_i)) Then
            Dic_Ref(StrArr_FileNames(int_i)).Add(System.IO.Path.GetFileNameWithoutExtension(Str_AppPath))
        Else
            Dic_Ref(StrArr_FileNames(int_i))={System.IO.Path.GetFileNameWithoutExtension(Str_AppPath)}.ToList
        End If 
      Catch ex As Exception
        Str_Error_Log=Str_Error_Log+Str_AppPath+" | "+ex.message+vbnewline
        console.writeline(Str_AppPath+vbnewline+ex.message)
      End Try
    Next
  Next
  Dim Str_ReadMeContent As String = Join(Dic_ReadMe.keys.Select(Function(k,i) String.Format("### {1}{0}```vb{0}{2}{0}```{0}",vbNewLine,k,Dic_ReadMe(k))).ToArray ,vbNewLine)
  Dim Str_ReadMelog As String = Join(Dic_Ref.keys.Select(Function(k) k+" : "+vbNewLine+"  - Count : "+Dic_Ref(k).Count.Tostring+vbNewLine+Join(Dic_Ref(k).Select(Function(x) "  - "+x).ToArray,vbNewLine)).ToArray,vbNewLine)
  System.IO.File.WriteAllText(Str_SavePath_Fnc+"\Readme.md",Str_ReadMeContent)
  System.IO.File.WriteAllText(Str_SavePath_Fnc+"\Readme.yaml",Str_ReadMelog)
  
  ' Open Result Directory
  System.Diagnostics.Process.Start("explorer.exe", Str_SavePath_Fnc)
  System.Diagnostics.Process.Start("notepad.exe", Str_SavePath_Fnc+"\Readme.yaml")
  Return Str_Error_Log
End Function
```

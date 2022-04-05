### Fnc_UI_Get_DirPath.vb

```yaml
Dereference :
  - 1. App_Convertion_IE_TO_Edge
  - 2. App_Update_Functions_And_Readme
```

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

### Fnc_Get_All_Files.vb

```yaml
Dereference :
  - 1. App_Convertion_IE_TO_Edge
  - 2. App_Update_Functions_And_Readme
```

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

### Fnc_Ui_MsgBox.vb

```yaml
Dereference :
  - 1. App_Convertion_IE_TO_Edge
```

```vb
Dim Fnc_Ui_MsgBox As System.Func(Of String, String, Boolean) = Function(Str_Massege As String, Str_title As String) As Boolean
  '2022.04.02|wbpark|message와 title을 입력받아 확인/취소 여부를 bool로 입력받습니다.
  Dim Result As System.Windows.Forms.DialogResult = System.Windows.Forms.MessageBox.Show(Str_Massege, Str_title, MessageBoxButtons.YesNo)
  Return If(Result = System.Windows.Forms.DialogResult.Yes,True,False)
End Function
```

### Fnc_Text_Replace.vb

```yaml
Dereference :
  - 1. App_Convertion_IE_TO_Edge
```

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

```yaml
Dereference :
  - 1. App_Convertion_IE_TO_Edge
```

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

```yaml
Dereference :
  - 1. App_Convertion_IE_TO_Edge
```

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

```yaml
Dereference :
  - 1. App_Update_Functions_And_Readme
```

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
  Dim Str_ReadMeContent As String = Join(Dic_ReadMe.keys.OrderByDescending(Function(k) Dic_Ref(k).Count).Select(Function(k) String.Format("### {1}{0}{0}```yaml{0}Dereference :{0}{3}{0}```{0}{0}```vb{0}{2}{0}```{0}",vbNewLine,k,Dic_ReadMe(k),Join(Dic_Ref(k).Select(Function(x,i) "  - "+(i+1).Tostring+". "+x).ToArray,vbNewLine) )).ToArray ,vbNewLine)
  Dim Str_Dereference As String = Join(Dic_Ref.keys.OrderByDescending(Function(k) Dic_Ref(k).Count).Select(Function(k) k+" : "+vbNewLine+Join(Dic_Ref(k).Select(Function(x,i) "  - "+(i+1).Tostring+". "+x).ToArray,vbNewLine)).ToArray,vbNewLine)
  System.IO.File.WriteAllText(Str_SavePath_Fnc+"\Readme.md",Str_ReadMeContent)
  System.IO.File.WriteAllText(Str_SavePath_Fnc+"\Dereference.yaml",Str_Dereference)
  ' Open Result Directory
  System.Diagnostics.Process.Start("explorer.exe", Str_SavePath_Fnc)
  System.Diagnostics.Process.Start("notepad.exe", Str_SavePath_Fnc+"\Dereference.yaml")
  Return Str_Error_Log
End Function
```

### Fnc_Kill_Process_By_Name.vb

```yaml
Dereference :
  - 1. Test_Module
```

```vb
Dim Fnc_Kill_Process_By_Name As System.Func(Of String, String) = Function(ProcessName As String) As String
  '2022.04.04|wbpark|DRM 있는 엑셀 종료를 위해 제작
  Dim Arr_process As System.Diagnostics.Process() = System.Diagnostics.Process.GetProcesses.Where(Function(x) x.ProcessName.ToUpper = ProcessName.ToUpper).toArray
    While Arr_process.Count>0
        For Each p As System.Diagnostics.Process In Arr_process
            p.kill()
        Next
        System.Threading.Thread.Sleep(1000)	'1초 딜레이
        Arr_process=System.Diagnostics.Process.GetProcesses.Where(Function(x) x.ProcessName.ToUpper = ProcessName.ToUpper).toArray
    End While 
    Return ProcessName+" 종료"
End Function
```

### Fnc_Back_Config.vb

```yaml
Dereference :
  - 1. Test_Module
```

```vb
Dim Fnc_Back_Config As System.Func(Of System.Collections.Generic.Dictionary(Of String, String),String) = Function(dic_config As System.Collections.Generic.Dictionary(Of String, String)) As String
  '2022.04.04|wbpark|Config Value에 null값이 있으면 에러가 발생합니다.
  Dim Str_Result = String.Format("{0}{0}{3}{0}{2}",vbNewLine,"New Dictionary(Of String,String) From {","}", Join( Dic_Config.Keys.Select(Function(key) String.Format("{0} {2}{3}{2} , {2}{4}{2} {1}", "{","}", chr(34), key, System.Convert.ToString(Dic_Config(key)).Replace(vbNewLine," ").Replace(chr(10)," ").Replace(chr(34),"'") ) ).ToArray, ","+vbNewLine) )
  System.IO.File.WriteAllText(System.IO.Path.Combine(System.Environment.CurrentDirectory,"Config.log",str_result))
  Return Str_Result
  'Fnc_Back_Config( New System.Collections.Generic.Dictionary(Of String, String) From {{"key","value"},{"k2","v2"}}  )}
End Function
```

### Fnc_Set_Env_With_Config.vb

```yaml
Dereference :
  - 1. Test_Module
```

```vb
Dim Fnc_Set_Env_With_Config As System.Func(Of System.Collections.Generic.Dictionary(Of String, Object), String) = Function(Dic_Config As System.Collections.Generic.Dictionary(Of String, Object) ) As String
  '2022.04.04|wbpark|Config값 환경변수로 넣기
  Dim Str_Error As String =""
  For Each Key As String In Dic_Config.Keys
    Try
      System.Environment.SetEnvironmentVariable(Key,in_Dic_Config(Key).tostring )
    Catch ex As System.Exception
		Str_Error=Str_Error+ex.message+vbnewline
    End Try
  Next
  System.Console.WriteLine(Str_Error)
  Return Str_Error
  'Fnc_Set_Env_With_Config(in_Dic_Config)
End Function
```

### Fnc_Bake_UiElementAttar.vb

```yaml
Dereference :
  - 1. Test_Module
```

```vb
Dim Fnc_Bake_UiElementAttar As System.Func(Of UiPath.Core.UiElement, String) = Function(ui_tmp As UiPath.Core.UiElement) As String
  '2022.04.04|wbpark|Uipath Activity와 함께 써야하는 함수. Invoke Code 내에서 Uipath Core 함수 호출 방법 불명.
  Dim Str_result As String  = ""
  Dim Str_fileName As String = Environment.CurrentDirectory+"\UiElementAttar.log"
  Dim dic_attar As dictionary(Of String, String) = ui_tmp.GetNodeAttributes(False)
  For Each key As String In dic_attar.keys()
    Str_result = String.format("{0}{2} : {3}{1}", Str_result, vbnewline,key,dic_attar(key))
  Next
  Str_result  = String.format("{0}{2} : {3}{1}", Str_result, vbnewline,"Selector",ui_tmp.Selector.Text)
  Str_result  = String.format("{0}{2} : {3}{1}", Str_result, vbnewline,"SelectorStrategy",ui_tmp.SelectorStrategy.ToString)
  Str_result  = String.format("{0}{2} : {3}{1}", Str_result, vbnewline,"ParentSelector",ui_tmp.Parent.Selector.Text)
  Str_result  = String.format("{0}{2} : {3}{1}", Str_result, vbnewline,"TopParent",ui_tmp.TopParent().Selector.Text)
  System.IO.File.WriteAllText(Str_fileName ,Str_result)
  Return Str_result
End Function
```

### Fnc_Mail_GetAttr.vb

```yaml
Dereference :
  - 1. Test_Module
```

```vb
Dim Fnc_Mail_GetAttr As System.Func(Of System.Net.Mail.MailMessage,String,String) = Function(msg_mail  As System.Net.Mail.MailMessage, Str_AttarName As String) As String
  ' 2022.04.05|wbpark|MailMessage 속성별 호출 방법 통일 시키기. 대소문자 구분 안함
  Str_AttarName=Str_AttarName.Trim.ToUpper
  If Str_AttarName = "SENDER" Then
    Return msg_mail.Sender.Address
  Else If Str_AttarName = "SUBJECT" Then
    Return msg_mail.Subject
  Else If Str_AttarName = "BODY" Then 
    Return msg_mail.Body
  Else If Str_AttarName = "TO" Then
    Return String.Join( ";", msg_mail.To.Select(Function(x) x.Address).ToArray)
  Else If Str_AttarName = "CC" Then
    Return String.Join( ";", msg_mail.CC.Select(Function(x) x.Address).ToArray)
  Else If Str_AttarName = "BCC" Then
    Return String.Join( ";", msg_mail.Bcc.Select(Function(x) x.Address).ToArray)
  Else If Str_AttarName = "ATTACH" Then
    Return String.Join(vbNewLine, msg_mail.Attachments.Select(Function(x) x.Name).ToArray)         
  Else
    Dim StrArr_HeaderKeys = "Uid|Date|DateCreated|DateRecieved|Size|Body|HtmlBody|PlainText".ToUpper().Split("|"c)
    Dim IndexHeader As Integer = Array.IndexOf(StrArr_HeaderKeys, Str_AttarName.ToUpper.Trim)
    If IndexHeader <> -1 Then
      Return msg_mail.Headers(StrArr_HeaderKeys(IndexHeader))
    Else
      Return "No AttarName : "+Str_AttarName
    End If 
  End If 
End Function
```

### Fnc_MailList_To_DataTable.vb

```yaml
Dereference :
  - 1. Test_Module
```

```vb
Dim Fnc_MailList_To_DataTable As System.Func(Of List(Of System.Net.Mail.MailMessage), String(), System.Data.DataTable)= Function(List_MailBox As List(Of System.Net.Mail.MailMessage), StrArr_Columns As String()) As System.Data.DataTable
  ' 2022.04.05|wbpark|메일함 데이터 DataTable로 만들어 저장
  Dim Dt_MailData As New DataTable
  For Each colName As String In StrArr_Columns
    Dt_MailData.Columns.Add(colName.Trim, System.Type.GetType("System.String"))
  Next
  For Each mail As System.Net.Mail.MailMessage In List_MailBox
    Dim StrArr_NewRow As String() = StrArr_Columns.Select(Function(x) Fnc_Mail_GetAttr(mail,x) ).ToArray
    Dt_MailData.Rows.Add(StrArr_NewRow)
  Next
  Return Dt_MailData
  ' 예시 : StrArr_ColNames = Sender|Subject|To|Cc|Bcc|Attach|Uid|Date|DateCreated|DateRecieved|Size|Body|HtmlBody|PlainText".Split("|"c).Distinct.toArray
  ' out_dt = Fnc_Convert_MailList_To_DataTable(in_list_mail,StrArr_ColNames)
End Function
```

### Fnc_UI_CustomDialog.vb

```yaml
Dereference :
  - 1. Test_UI_CustomDialog
```

```vb
Dim Fnc_UI_CustomDialog As System.Func(Of String,String,String,String) = Function(caption As String, text As String, selStr As String) As String
  Dim prompt As New System.Windows.Forms.Form With {.Width = 280, .Height = 200, .Text = caption}
  Dim textLabel As New System.Windows.Forms.Label With { .Left = 16, .Top = 20, .Width = 240, .Text = text }
  Dim textBox As New System.Windows.Forms.TextBox With { .Left = 16, .Top = 50, .Width = 240, .TabIndex = 0, .TabStop = True }
  Dim selLabel As New System.Windows.Forms.Label With { .Left = 16, .Top = 130, .Width = 88, .Text = selStr }
  Dim cmbx As New System.Windows.Forms.ComboBox With { .Left = 112, .Top = 130, .Width = 144}
  cmbx.Items.Add("Dark Grey")
  cmbx.Items.Add("Orange")
  cmbx.Items.Add("None")
  cmbx.SelectedIndex = 0
  Dim confirmation As New System.Windows.Forms.Button With { .Text = "In Ordnung!", .Left = 16, .Width = 80, .Top = 88, .TabIndex = 1, .TabStop = True }
  AddHandler confirmation.Click, Sub(sender, e) prompt.Close()
  prompt.Controls.Add(textLabel)
  prompt.Controls.Add(textBox)
  prompt.Controls.Add(selLabel)
  prompt.Controls.Add(cmbx)
  prompt.Controls.Add(confirmation)
  prompt.AcceptButton = confirmation
  prompt.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
  prompt.TopMost=True
  prompt.ShowDialog()
  Return String.Format("{0};{1}", textBox.Text, cmbx.SelectedItem.ToString)
End Function

'Dim Str_tmp As String = Fnc_UI_CustomDialog("caption","text","selStr")
'console.WriteLine(Str_tmp)
              
' New With  :
' https://docs.microsoft.com/en-us/dotnet/visual-basic/programming-guide/language-features/objects-and-classes/object-initializers-named-and-anonymous-types
' code : 
' https://stackoverflow.com/questions/5427020/prompt-dialog-in-windows-forms
' https://docs.microsoft.com/ko-kr/dotnet/api/system.windows.forms.combobox.text?view=windowsdesktop-6.0&viewFallbackFrom=dotnet-plat-ext-6.0
```

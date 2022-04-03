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
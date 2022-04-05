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
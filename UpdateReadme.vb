Module RunVB
    Public Sub Main()
'-----------------------------------------------
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
'-----------------------
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
'-----------------------
' App
'-----------------------
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
'-----------------------
' Main
'-----------------------
' Input
Dim in_Str_DirPath As String = ""
in_Str_DirPath = System.IO.Path.Combine(System.Environment.CurrentDirectory,"Apps")
App_Update_Functions_And_Readme(in_Str_DirPath)
'-----------------------------------------------
    end Sub
end Module 

if string.IsNullOrWhiteSpace(in_Str_ProjectCode)
>>> input dialog
>>> "Convertion IE -> Edge"
>>> String.Format("in_Str_ProjectCode 가 입력되지 않았습니다.{0}IE->Edge 변환할 폴더를 선택해주세요.{0}{0}과제코드 입력시 :{0}폴더 설정{0}에서 파일을 추출합니다.{0}{0} 절대경로 입력시 해당 경로의 프로젝트를 변환합니다.",vbNewLine)
>>> in_Str_ProjectCode
End if
Str_SourceDir = if(system.io.Directory.Exists(in_Str_ProjectCode),in_Str_ProjectCode, System.IO.Path.Combine("폴더 설정",in_Str_ProjectCode))

if not System.IO.Directory.Exists( Str_SourceDir )
    throw new BusinessRuleException( Str_SourceDir + " 폴더가 존재하지 않습니다." )
End if
    
if string.IsNullOrWhiteSpace(in_Str_Result_FolderPath)
        >>> "in_Str_Result_FolderPath 가 입력되지 않았습니다."+vbNewLine+" 변환 결과 파일을 확인할 폴더를 입력하세요."
        >>> selectfolder :  in_Str_Result_FolderPath
end if
Str_ResultDir = System.IO.Path.Combine(in_Str_Result_FolderPath, if(system.io.Directory.Exists(in_Str_ProjectCode) ,system.io.Path.GetFileNameWithoutExtension(in_Str_ProjectCode), in_Str_ProjectCode) )
'----------------------------------------
'----------------------------------------        
'----------------------------------------
'----------------------------------------
Dim Str_Dir_Source As String = in_Str_SourceDir
Dim Str_Dir_Result As String = in_Str_ResultDir
Dim int_Result_Cnt As Integer = 0

'-----------------------------------------------
' 폴더 내 모든 파일 조회
Dim Fnc_Get_All_Files As System.Func(Of String, String()) = Function(str_path As String) As String()
    Dim list_str_dir As New System.Collections.Generic.List(Of String) 
    Dim list_str_file As New System.Collections.Generic.List(Of String)
    
    list_str_dir.Add(str_path)
    While list_str_dir.Count <> 0
    str_path = list_str_dir.Last
    list_str_dir.RemoveAt(list_str_dir.Count -1)
    list_str_dir.AddRange(System.IO.Directory.GetDirectories(str_path))
    list_str_file.AddRange(System.IO.Directory.GetFiles(str_path))
    End While
    Return list_str_file.ToArray
End Function
'-----------------------------------------------
' main
Dim StrArr_Before As String() = {"&lt;html html","&lt;html title","&lt;html idx","&lt;html app='iexplore.exe'","BrowserType=""IE"""}
Dim StrArr_After As String() = {"&lt;html app='msedge.exe' html","&lt;html app='msedge.exe' title","&lt;html app='msedge.exe' idx","&lt;html app='msedge.exe'","BrowserType=""Edge"""}
Dim Str_ptn As String = "Selector=""\[[^]]+\]"""

Dim Str_Msg_title As String = "과제설명"
Dim Str_Desc_msg As String = String.format("###########{0} 패키지 업데이트는 수동으로 진행하셔야 합니다.{0}파일명에 %20이 있는 경우 ' '으로 바꿔줍니다.{0}###########{0}{0}Attach Browser의 BrowserType을 IE에서 Edge로 바꿉니다.{0}",VbNewLine)
Str_Desc_msg=Str_Desc_msg+"Selector 내 변수가 사용된 경우 {{ }}로 바꾸어줍니다."+VbNewLine+VbNewLine+"before -> After"+VbNewLine
Str_Desc_msg=Str_Desc_msg+Join(StrArr_Before.Select(Function(x,i) x+" => "+StrArr_After(i) ).ToArray,VbNewLine)
microsoft.visualbasic.interaction.msgbox(Str_Desc_msg,,Str_Msg_title)

System.Console.WriteLine("변환 완료 : ")
For Each Str_FilePath As String In Fnc_Get_All_Files(Str_Dir_Source)
    ' 파일 작성
	Dim Str_File_ResultPath As String = Str_FilePath.replace(Str_Dir_Source,Str_Dir_Result)    
	Str_File_ResultPath=Str_File_ResultPath.replace("%20"," ")
	System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(Str_File_ResultPath))
    
    'xaml은 일괄적으로 변경
    If System.IO.Path.GetExtension(Str_FilePath).ToUpper = ".XAML"
        Dim Str_FileContent As String = System.IO.File.ReadAllText(Str_FilePath)
        
        ' 셀렉터 내 변수 부분 {{}}로 바꿔주기
        For Each x As System.Text.RegularExpressions.Match In System.Text.RegularExpressions.Regex.Matches(Str_FileContent,Str_ptn)
            Dim Str_replaceTo As String = x.Tostring.replace("[&quot;","").replace("&quot;+","{{").replace("+&quot;","}}").replace("&quot;]","").replace("&quot;","""")
            Str_FileContent=Str_FileContent.replace(x.tostring,Str_replaceTo)
        Next

        ' app 및 Attach Edge,IE 일괄 변경
        For int_i As Integer = 0 To StrArr_Before.length-1
            Str_FileContent=Str_FileContent.replace(StrArr_Before(int_i),StrArr_After(int_i))
        Next 

        System.IO.File.WriteAllText(Str_File_ResultPath, Str_FileContent, System.Text.Encoding.UTF8)
    Else
        System.IO.File.Copy(Str_FilePath, Str_File_ResultPath)
    End If
    System.Console.WriteLine(int_Result_Cnt.Tostring("00")+" | "+Str_File_ResultPath)
Next
System.Diagnostics.Process.Start("explorer.exe", Str_Dir_Result)

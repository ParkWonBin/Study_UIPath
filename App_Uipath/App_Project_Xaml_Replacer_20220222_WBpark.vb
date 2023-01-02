'2022.02.22. 박원빈 프로
'UiPath 프로젝트 내 특정 키워드를 Replace하는 모듈입니다.
'WBpark__App_Xaml_Replacer_Converion_IE2Edge.vb
'-----------------------------------------------
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
Dim StrArr_Before As String() = {"&lt;html url","&lt;html html","&lt;html title","&lt;html idx","&lt;html app='iexplore.exe'","BrowserType=""IE""","BrowserType=""{x:Null}""","ProcessName=""iexplorer"">","ProcessName=""iexplore"">","ProcessName=""iexplorer"" />","ProcessName=""iexplore"" />","ProcessName=""iexplorer""/>","ProcessName=""iexplore""/>"}
Dim StrArr_After As String() = {"&lt;html app='msedge.exe' url","&lt;html app='msedge.exe' html","&lt;html app='msedge.exe' title","&lt;html app='msedge.exe' idx","&lt;html app='msedge.exe'","BrowserType=""Edge""","BrowserType=""Edge""","ProcessName=""msedge"">","ProcessName=""msedge"">","ProcessName=""msedge"" />","ProcessName=""msedge"" />","ProcessName=""msedge""/>","ProcessName=""msedge""/>"}
Dim Str_ptn As String = "Selector=""\[[^]]+\]"""

Dim Str_Msg_title As String = "과제설명"
Dim Str_Desc_msg As String = String.format("###########{0} 패키지 업데이트는 수동으로 진행하셔야 합니다.{0}파일명에 %20이 있는 경우 ' '으로 바꿔줍니다.{0}###########{0}{0}Attach Browser의 BrowserType을 IE에서 Edge로 바꿉니다.{0}",VbNewLine)
Str_Desc_msg=Str_Desc_msg+"Selector 내 변수가 사용된 경우 {{ }}로 바꾸어줍니다."+VbNewLine+VbNewLine+"before -> After"+VbNewLine
Str_Desc_msg=Str_Desc_msg+Join(StrArr_Before.Select(Function(x,i) x+" => "+StrArr_After(i) ).ToArray,VbNewLine)
microsoft.visualbasic.interaction.msgbox(Str_Desc_msg,vbSystemModal,Str_Msg_title)

System.Console.WriteLine("변환 완료 : ")
For Each Str_FilePath As String In Fnc_Get_All_Files(Str_Dir_Source)
    ' 파일 작성
	Dim Str_File_ResultPath As String = Str_FilePath.replace(Str_Dir_Source,Str_Dir_Result)    
	Str_File_ResultPath=Str_File_ResultPath.replace("%20"," ")
	System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(Str_File_ResultPath))
    
    'xaml은 일괄적으로 변경
    If System.IO.Path.GetExtension(Str_FilePath).ToUpper = ".XAML" Then
        Dim Str_FileContent As String = System.IO.File.ReadAllText(Str_FilePath)
        
       ' 셀렉터 내 변수 부분 {{}}로 바꿔주기
        For Each x As System.Text.RegularExpressions.Match In System.Text.RegularExpressions.Regex.Matches(Str_FileContent,Str_ptn) 
		    If Not (x.Tostring.ToUpper.Contains(".TOSTRING") OrElse x.ToString.Contains("(") OrElse  x.ToString.Contains(")") ) Then
	            Dim Str_replaceTo As String = x.Tostring.replace("[&quot;","").replace("&quot;+","{{").replace("+&quot;","}}").replace("&quot;]","")
	            Str_FileContent=Str_FileContent.replace(x.tostring,Str_replaceTo)C
			End If
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
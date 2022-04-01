'Convertion_IE_TO_Edge.vb
'UIAutomation의 버전상이할 경우, 버전이 높든 낮든 상관 없이 다른 버전에서 Validate 한 셀렉터를 제대로 인식하지 못하는 문제가 있습니다.
'이 앱은 UIAutomation=21.4를기준으로 동작합니다.
Imports System
Imports Microsoft
Module Convertion_IE_TO_Edge
    Public Sub Main()

'-----------------------------------------------
' 폴더 입력받기 UI
Dim Fnc_Get_DirPath As System.Func(Of String, String) = Function(str_Desc As String) As String
    Dim folderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog = New System.Windows.Forms.FolderBrowserDialog() 

    folderBrowserDialog1.Description = str_Desc
    folderBrowserDialog1.ShowNewFolderButton = False
    folderBrowserDialog1.RootFolder = Environment.SpecialFolder.MyComputer

    Dim result As System.Windows.Forms.DialogResult = folderBrowserDialog1.ShowDialog()
    If result = System.Windows.Forms.DialogResult.OK Then
            Return FolderBrowserDialog1.SelectedPath
    End If 
    Return ""
End Function
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
Dim StrArr_Before As String() = {"&lt;html html","&lt;html title","&lt;html idx","&lt;html app='iexplore.exe'","BrowserType=""IE""","BrowserType=""{x:Null}""","ProcessName=""iexplore"" />"}
Dim StrArr_After As String() = {"&lt;html app='msedge.exe' html","&lt;html app='msedge.exe' title","&lt;html app='msedge.exe' idx","&lt;html app='msedge.exe'","BrowserType=""Edge""","BrowserType=""Edge""","ProcessName=""msedge"" />"}
Dim Str_ptn As String = "Selector=""\[[^]]+\]"""
Dim Str_Description As String = "###############\n#IE -> Edge 변환 :#\n변환할 Project의 폴더를 선택해주세요.".replace("\n",VbNewLine)

Dim Str_Msg_title As String = "과제설명"
Dim Str_Desc_msg As String = String.format("###########{0} 패키지 업데이트는 수동으로 진행하셔야 합니다.{0}파일명에 %20이 있는 경우 ' '으로 바꿔줍니다.{0}###########{0}{0}Attach Browser의 BrowserType을 IE에서 Edge로 바꿉니다.{0}",VbNewLine)
Str_Desc_msg=Str_Desc_msg+"Selector 내 변수가 사용된 경우 {{ }}로 바꾸어줍니다."+VbNewLine+"(셀렉터 '.tostring','(',')' 미포함 시에만 작동)"+VbNewLine+VbNewLine+"before -> After"+VbNewLine
Str_Desc_msg=Str_Desc_msg+Join(StrArr_Before.Select(Function(x,i) x+" => "+StrArr_After(i) ).ToArray,VbNewLine)
microsoft.visualbasic.interaction.msgbox(Str_Desc_msg,vbSystemModal,Str_Msg_title)

Dim Str_Dir_Source As String = Fnc_Get_DirPath(Str_Description)
Dim Str_Dir_Result As String = Str_Dir_Source+"_Replaced"
Dim int_Result_Cnt As Integer = 0

System.Console.WriteLine("변환 완료 : ")
For Each Str_FilePath As String In Fnc_Get_All_Files(Str_Dir_Source)
    '결과 폴더를 따로 만들기
    Dim Str_File_ResultPath As String = Str_FilePath.replace(Str_Dir_Source,Str_Dir_Result)    
    System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(Str_File_ResultPath))

    'xaml은 일괄적으로 변경 후 저장
    If System.IO.Path.GetExtension(Str_FilePath).ToUpper = ".XAML"
        Dim Str_FileContent As String = System.IO.File.ReadAllText(Str_FilePath)
        
       ' 셀렉터 내 변수 부분 {{}}로 바꿔주기
	For Each x As System.Text.RegularExpressions.Match In System.Text.RegularExpressions.Regex.Matches(Str_FileContent,Str_ptn) 
		If Not (x.Tostring.ToUpper.Contains(".TOSTRING") OrElse x.ToString.Contains("(") OrElse  x.ToString.Contains(")") ) Then
			Dim Str_replaceTo As String = x.Tostring.replace("[&quot;","").replace("&quot;+","{{").replace("+&quot;","}}").replace("&quot;]","")
			Str_FileContent=Str_FileContent.replace(x.tostring,Str_replaceTo)
		End If
	Next

        ' app 및 Attach Edge,IE 일괄 변경
        For int_i As Integer = 0 To StrArr_Before.length-1
            Str_FileContent=Str_FileContent.replace(StrArr_Before(int_i),StrArr_After(int_i))
        Next 

        ' 파일 작성
        Str_File_ResultPath=Str_File_ResultPath.replace("%20"," ")
        System.IO.File.WriteAllText(Str_File_ResultPath, Str_FileContent, System.Text.Encoding.UTF8)
    Else
        System.IO.File.Copy(Str_Dir_Source, Str_File_ResultPath)
    End If
    System.Console.WriteLine(int_Result_Cnt.Tostring("00")+" | "+Str_File_ResultPath)
Next
System.Diagnostics.Process.Start("explorer.exe", System.IO.Path.GetDirectoryName(Str_Dir_Result))

    end Sub
end Module 
